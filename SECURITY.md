# Security

open-xlsx-editor runs in the browser and opens files people give it. So it is a security problem if a workbook can
make it run script, read data it shouldn't, or hang the page. So is anything that lets a saved file carry something
its author didn't put there.

## Reporting a problem

Please don't open a public issue for it. Use GitHub's private reporting instead: on the repository's **Security** tab,
choose **Report a vulnerability**. Include:

- the workbook, or a small one that shows the problem;
- the browser;
- what happened.

## What it does to keep workbooks safe

- **Nothing in a workbook is run.** An .xlsm's macros (its VBA project) are kept as bytes and written back, never
  executed. There are no scripts and no remote content. External links and data connections are kept in the file,
  but they are not refreshed.
- **The editor fetches nothing while it runs.** It sends nothing anywhere and loads no fonts or code from another
  host. Fonts come only from a folder the host page names (README, Fonts).
- **Only safe links are followed.** A hyperlink is followed only to an `http:`, `https:`, `mailto:` or `ftp:`
  address, or to a place in the workbook. Any other address (`javascript:`, `data:`, `file:`...) stays in the file as
  it was, but clicking it shows Excel's "Cannot open the specified file." A host page can take the address itself
  with the `onLink` option.
- **The workbook's text is never treated as HTML.** Its XML is read as data, and its text is drawn on a canvas or put
  in the page as text. Charts and shapes are drawn from the editor's own code.
- **Saving writes back only what was there.** The saved file holds the parts that were in it, plus only the ones an
  edit needs.
- **Passwords stay in the browser.** A password-protected workbook is decrypted and encrypted with the password given
  (docs/ENCRYPTION.md). The password isn't stored or sent.
- **The MCP server stays inside its folder.** It reads and writes only .xlsx and .xlsm files inside the folder it is
  given, and makes no network requests (docs/AGENT.md).

## Supported versions

Fixes go into the latest release.
