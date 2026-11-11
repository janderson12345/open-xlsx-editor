# Workbooks saved with a password

Excel's **File > Info > Protect Workbook > Encrypt with Password** doesn't save a zip. It saves a compound file
(MS-CFB, `src/ooxml/cfb.js`) holding the encrypted .xlsx, as MS-OFFCRYPTO's agile encryption describes
(`src/ooxml/encryption.js`). The compound file has these streams:

- `EncryptionInfo`: the settings, salts and wrapped keys, as XML after an 8-byte header (version 4.4)
- `EncryptedPackage`: the .xlsx's size, then its bytes in AES-CBC, 4,096 bytes at a time, each segment with its own IV
- four `\x06DataSpaces` streams that name the transform

The password opens a random key that encrypts the package. It is the salt and the password (UTF-16) hashed, then
hashed again `spinCount` times (100,000 in Excel's files). An HMAC over the package shows that nobody has changed it.

The hashing and AES come from `@noble/hashes` and `@noble/ciphers` (MIT, audited, no dependencies of their own). They
are synchronous, so opening and saving stay synchronous.

## What works

- **Opening.** `Workbook.fromBytes(bytes, { password })` and `openWorkbook(bytes, { password })` open a protected
  workbook. In the editor, `api.open(bytes, name)` shows Excel's Password box ("'name' is protected.") and returns a
  promise. A wrong password gets Excel's message ("The password you supplied is not correct...") and nothing opens,
  as in Excel.
- **Refusals.** Without a password, with a wrong one, with a package that fails its integrity check, or with an older
  encryption, opening throws a `PasswordError` with `reason` set to `needed`, `wrong`, `damaged` or `unsupported`.
- **Saving.** A workbook opened with a password is saved encrypted with the same password. Every salt and key is new
  each time, as Excel's are, and the data-spaces streams it came with are kept. `setPassword(pw)` sets a new password,
  or encrypts a workbook that had none (as Encrypt with Password does). `setPassword(null)` saves it as a plain .xlsx.
- **What's read.** Agile encryption (Excel 2010 and later) with SHA-1, SHA-256, SHA-384 or SHA-512, and AES-128, 192
  or 256. **What's written** is what Excel 365 writes: AES-256, SHA-512, 100,000 spins.

Not handled: the older "standard" encryption of Excel 2007 (`unsupported`), and .xls files.

## Measured against Excel (10 November 2026)

A hidden Excel saved `basic.xlsx` and `formulas.xlsx` with a test password (`SaveAs` with a password). Those files
stay outside the repository, because Excel writes its user's name into them.

| Check | Result |
|---|---|
| Excel's encrypted files opened here | Both, with the right password. No password and a wrong one refused. |
| The data-spaces streams written here | The same bytes as Excel's, all four (also a unit test). |
| Excel's `basic.xlsx` decrypted and encrypted again here, then opened in Excel | Opened with the password, all 3 sheets; a wrong password refused. |
| `formulas.xlsx` encrypted fresh here (new data spaces), then opened in Excel | Opened with its password, all 3 sheets; a wrong password refused. |
| Excel's `formulas.xlsx` opened here, two cells edited, saved, opened in Excel | Opened with the password: Data!B2 = 7, H20 = "edited here", the table's 8 rows intact. |

Decrypting takes about a second in Node, most of it the 100,000 spins, done three times: once each for the verifier
input, the verifier hash and the key.

Tests:

- `test/encryption.test.js`:
  - compound files, including a 9 MB stream that needs DIFAT sectors
  - the data spaces against Excel's bytes
  - round trips, including a password beyond ASCII
  - the four refusals
  - an encrypted workbook edited and saved encrypted, its password changed and then removed
- `test/editing-more.mjs`: the Password box, a wrong password, then the right one, and the workbook saved encrypted again.
- `test/a11y.mjs`: axe on the Password box, in light and dark.
