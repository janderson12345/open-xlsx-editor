// The cases of Sort measured in Excel (excel-steps.ps1: Worksheet.Sort, as Excel's Sort box and its A to Z and Z to A
// buttons apply it): the order of numbers, text, TRUE and FALSE, errors and blanks; how text compares (case,
// punctuation, digits); ties; whether the first row is taken as headers; formulas, formats and several keys; left to
// right; the range found around one cell; merged cells and arrays.
// usage: node tools/excel-compare/sort-cases.mjs > tools/excel-compare/sort-cases.json
const col = (c, from, values) => Object.fromEntries(values.map((v, i) => [`${c}${from + i}`, v]).filter(([, v]) => v !== ''));
// (Each row's B says where it started, to see ties and what moves with the row.)
const tagged = (values, from = 2) => ({ ...col('A', from, values), ...col('B', from, values.map((_, i) => `r${from + i}`)) });

const mixed = ['3', 'b', 'TRUE', '=NA()', '', 'a', '-1', 'FALSE', 'B', "'10", '=1/0', '2.5', '=""', "' ", '=""', '0'];
const words = ['a', 'B', 'ab', 'a-b', 'a b', "a'b", 'a_b', 'Ab', 'z', '1a', "'10", "'2", '_x', '-x', '#x', '!x', 'x1', 'x10', 'x2', ' a', 'A', 'aa', 'é', 'e', 'f', 'co-op', 'coop', 'cop', '(x)', '[x]', '~x', 'x.y', 'xy'];
const cases = [];
const add = (name, setup, steps, read) => cases.push({ name, setup, steps, read });

add('mixed types ascending', { cells: { A1: 'Item', B1: 'Row', ...tagged(mixed) } }, [{ do: 'sort', range: 'A1:B17', keys: [{ key: 'A1' }], header: 'yes' }], 'A1:B18');
add('mixed types descending', { cells: { A1: 'Item', B1: 'Row', ...tagged(mixed) } }, [{ do: 'sort', range: 'A1:B17', keys: [{ key: 'A1', order: 'desc' }], header: 'yes' }], 'A1:B18');
add('text ascending', { cells: { A1: 'Word', B1: 'Row', ...tagged(words) } }, [{ do: 'sort', range: `A1:B${words.length + 1}`, keys: [{ key: 'A1' }], header: 'yes' }], `A1:B${words.length + 2}`);
add('text descending', { cells: { A1: 'Word', B1: 'Row', ...tagged(words) } }, [{ do: 'sort', range: `A1:B${words.length + 1}`, keys: [{ key: 'A1', order: 'desc' }], header: 'yes' }], `A1:B${words.length + 2}`);
add('text, match case', { cells: { A1: 'Word', B1: 'Row', ...tagged(['b', 'A', 'a', 'B', 'ab', 'Ab', 'aB', 'AB', 'a', 'A']) } }, [{ do: 'sort', range: 'A1:B11', keys: [{ key: 'A1' }], header: 'yes', matchCase: true }], 'A1:B12');
add('text, ignoring case: ties kept in order', { cells: { A1: 'Word', B1: 'Row', ...tagged(['b', 'A', 'a', 'B', 'ab', 'Ab', 'aB', 'AB', 'a', 'A']) } }, [{ do: 'sort', range: 'A1:B11', keys: [{ key: 'A1' }], header: 'yes' }], 'A1:B12');
add('numbers stored as text, as numbers', { cells: { A1: 'N', B1: 'Row', ...tagged(["'10", '9', "'2", '1', "'1.5", 'x', "'-3"]) } }, [{ do: 'sort', range: 'A1:B8', keys: [{ key: 'A1', textAsNumbers: true }], header: 'yes' }], 'A1:B9');
add('dates and times', { cells: { A1: 'When', B1: 'Row', ...tagged(['3/1/2024', '1/15/2024', '9:30', '12/31/2023', '45000', '1/15/2024 8:00']) } }, [{ do: 'sort', range: 'A1:B7', keys: [{ key: 'A1' }], header: 'yes' }], 'A1:B8');
// (Each printable ASCII sign, and a few others, before an x: their order among themselves and against digits.)
const signs = [...' !"#$%&\'()*+,-./:;<=>?@[\\]^_`{|}~', '\u2019', '\u2013', '\u00a0', '\u00a7', '\u00b0'].map((ch) => (ch === ' ' ? ' x' : `${ch}x`));
const signed = [...signs.map((t) => (/^[=+@-]/.test(t) || t.startsWith("'") ? `'${t}` : t)), '0x', 'x', 'Ax', 'ñ', 'n', 'o', 'ß', 'ss', 'st', 'æ', 'ae', 'af'];
for (const order of ['asc', 'desc']) {
  add(`signs ${order === 'asc' ? 'ascending' : 'descending'}`, { cells: { A1: 'Sign', B1: 'Row', ...tagged(signed) } },
    [{ do: 'sort', range: `A1:B${signed.length + 1}`, keys: [{ key: 'A1', order }], header: 'yes' }], `A1:B${signed.length + 2}`);
}

// Headers: whether Excel takes the first row as headers when not told (xlGuess, as the A to Z button does).
add('header guessed: text over text and numbers', { cells: { A1: 'Name', B1: 'Qty', A2: 'c', A3: 'a', A4: 'b', B2: '3', B3: '1', B4: '2' } }, [{ do: 'sort', range: 'A1:B4', keys: [{ key: 'A1' }], legacy: true }], 'A1:B5');
add('header guessed: text over dates', { cells: { A1: 'When', A2: '3/1/2024', A3: '1/1/2024', A4: '2/1/2024' } }, [{ do: 'sort', range: 'A1:A4', keys: [{ key: 'A1' }], legacy: true }], 'A1:B5');
add('header guessed: text over text and a number', { cells: { A1: 'm', A2: 'c', A3: '5', A4: 'b' } }, [{ do: 'sort', range: 'A1:A4', keys: [{ key: 'A1' }], legacy: true }], 'A1:B5');
add('header guessed: a number over text', { cells: { A1: '7', A2: 'c', A3: 'a', A4: 'b' } }, [{ do: 'sort', range: 'A1:A4', keys: [{ key: 'A1' }], legacy: true }], 'A1:B5');
// (Which columns count: one of four, two of four, one of three, a column blank in both rows, text over a blank.)
const guess = (name, cells, formats) => add(`header guessed: ${name}`, { cells, formats }, [{ do: 'sort', range: 'A1', region: true, keys: [{ key: 'A1' }], legacy: true }], 'A1:E7');
guess('one column of four differs', { A1: 'Name', B1: 'Qty', C1: 'Cat', D1: 'Note', A2: 'c', B2: '3', C2: 'x', D2: 'y', A3: 'a', B3: '1', C3: 'w', D3: 'v' });
guess('two columns of four differ', { A1: 'Name', B1: 'Qty', C1: 'Amt', D1: 'Note', A2: 'c', B2: '3', C2: '30', D2: 'y', A3: 'a', B3: '1', C3: '10', D3: 'v' });
guess('one column of three differs', { A1: 'Name', B1: 'Qty', C1: 'Cat', A2: 'c', B2: '3', C2: 'x', A3: 'a', B3: '1', C3: 'w' });
guess('two columns of three differ', { A1: 'Name', B1: 'Qty', C1: 'Amt', A2: 'c', B2: '3', C2: '30', A3: 'a', B3: '1', C3: '10' });
guess('a column blank in both rows', { A1: 'Name', B1: 'Qty', A2: 'c', B2: '3', A3: 'a', B3: '1', C3: 'z' });
guess('text over a blank', { A1: 'Name', B1: 'Qty', A2: 'c', A3: 'a', B3: '1', A4: 'b', B4: '2' });
guess('the diagonal cells taken away', { A1: 'Name', B1: 'Qty', A2: 'c', B2: '3', A3: 'a', B3: '1', A4: 'q' });
guess('a border under the first row', { A1: 'm', B1: 'x', A2: 'c', A3: 'a', A4: 'b', B2: 'r2', B3: 'r3', B4: 'r4' }, [{ range: 'A1:B1', border: true }]);
guess('one bold cell of two', { A1: 'm', B1: 'x', A2: 'c', A3: 'a', A4: 'b', B2: 'r2', B3: 'r3', B4: 'r4' }, [{ range: 'A1', bold: true }]);
guess('one bold cell of four', { A1: 'm', B1: 'x', C1: 'y', D1: 'z', A2: 'c', A3: 'a', B2: 'r2', B3: 'r3', C2: 'p', C3: 'q', D2: 's', D3: 't' }, [{ range: 'A1', bold: true }]);
guess('the second row bold, not the first', { A1: 'm', B1: 'x', A2: 'c', A3: 'a', A4: 'b', B2: 'r2', B3: 'r3', B4: 'r4' }, [{ range: 'A2:B2', bold: true }]);
guess('two columns blank in both rows', { A1: 'Name', B1: 'Qty', A2: 'c', B2: '3', A3: 'a', B3: '1', C3: 'z', D3: 'w' });
guess('two columns differ, two blank in both rows', { A1: 'Name', B1: 'Qty', C1: 'Amt', A2: 'c', B2: '3', C2: '30', A3: 'a', B3: '1', C3: '10', D3: 'z', E3: 'w' });
guess('three columns blank in both rows', { A1: 'Name', B1: 'Qty', A2: 'c', B2: '3', A3: 'a', B3: '1', C3: 'z', D3: 'w', E3: 'v' });
guess('bold, two columns blank in both rows', { A1: 'm', B1: 'x', A2: 'c', B2: 'r2', A3: 'a', B3: 'r3', C3: 'z', D3: 'w' }, [{ range: 'A1:B1', bold: true }]);
guess('two blanks in the first row only', { C1: 'Amt', A2: '3', B2: '30', C2: '300', A3: '1', B3: '10', C3: '100' });
guess('two blanks in the second row only', { A1: 'Name', B1: 'Qty', C1: 'Amt', D1: 'Cat', A2: 'c', B2: '3', A3: 'a', B3: '1', C3: '5', D3: 'x' });
guess('two blanks in the first row of six', { A1: 'Name', B1: 'Qty', C1: 'Amt', D1: 'Cat', A2: 'c', B2: '3', C2: '30', D2: 'p', E2: '7', F2: '8', A3: 'a', B3: '1', C3: '10', D3: 'q', E3: '9', F3: '6' });
guess('one blank in each of the two rows', { A1: 'Name', B1: 'Qty', A2: 'c', C2: '30', A3: 'a', B3: '1', C3: '10' });
guess('text over numbers, the third row text', { A1: 'Qty', A2: '3', A3: 'x', A4: '1' });
add('header guessed: Sort object', { cells: { A1: 'Qty', B1: 'Amt', A2: '3', A3: '1', A4: '2', B2: '30', B3: '10', B4: '20' } }, [{ do: 'sort', range: 'A1:B4', keys: [{ key: 'A1' }] }], 'A1:B5');
const nums = { A2: '3', A3: '1', A4: '2', B2: '30', B3: '10', B4: '20' };
add('header guessed: text over numbers', { cells: { A1: 'Qty', B1: 'Amt', ...nums } }, [{ do: 'sort', range: 'A1:B4', keys: [{ key: 'A1' }], legacy: true }], 'A1:B5');
add('header guessed: numbers over numbers', { cells: { A1: '5', B1: '50', ...nums } }, [{ do: 'sort', range: 'A1:B4', keys: [{ key: 'A1' }], legacy: true }], 'A1:B5');
add('header guessed: text over text', { cells: { A1: 'm', B1: 'x', A2: 'c', A3: 'a', A4: 'b', B2: 'r2', B3: 'r3', B4: 'r4' } }, [{ do: 'sort', range: 'A1:B4', keys: [{ key: 'A1' }], legacy: true }], 'A1:B5');
add('header guessed: bold text over text', { cells: { A1: 'm', B1: 'x', A2: 'c', A3: 'a', A4: 'b', B2: 'r2', B3: 'r3', B4: 'r4' }, formats: [{ range: 'A1:B1', bold: true }] }, [{ do: 'sort', range: 'A1:B4', keys: [{ key: 'A1' }], legacy: true }], 'A1:B5');
add('header guessed: text over numbers in one column only', { cells: { A1: 'Qty', B1: '50', ...nums } }, [{ do: 'sort', range: 'A1:B4', keys: [{ key: 'A1' }], legacy: true }], 'A1:B5');
add('header guessed: text over numbers in the other column', { cells: { A1: '5', B1: 'Amt', ...nums } }, [{ do: 'sort', range: 'A1:B4', keys: [{ key: 'A1' }], legacy: true }], 'A1:B5');
add('header guessed: a blank over numbers', { cells: { B1: 'Amt', ...nums } }, [{ do: 'sort', range: 'A1:B4', keys: [{ key: 'A1' }], legacy: true }], 'A1:B5');
add('header guessed: number format differs', { cells: { A1: '5', B1: '50', ...nums }, formats: [{ range: 'A1:B1', numFmt: '0.00' }] }, [{ do: 'sort', range: 'A1:B4', keys: [{ key: 'A1' }], legacy: true }], 'A1:B5');
add('header guessed: fill differs', { cells: { A1: '5', B1: '50', ...nums }, formats: [{ range: 'A1:B1', color: 65535 }] }, [{ do: 'sort', range: 'A1:B4', keys: [{ key: 'A1' }], legacy: true }], 'A1:B5');
add('header guessed: one row', { cells: { A1: 'b', B1: 'a' } }, [{ do: 'sort', range: 'A1:B1', keys: [{ key: 'A1' }], legacy: true }], 'A1:B2');
add('no header', { cells: { A1: 'Qty', B1: 'Amt', ...nums } }, [{ do: 'sort', range: 'A1:B4', keys: [{ key: 'A1' }], header: 'no' }], 'A1:B5');

// Formulas and formats travel with their rows.
const sheet = {
  cells: {
    A1: 'Name', B1: 'Qty', C1: 'Twice', D1: 'Fixed', E1: 'Next', F1: 'Label',
    A2: 'c', B2: '3', C2: '=B2*2', D2: '=$B$2', E2: '=B3', F2: '=A2&"!"',
    A3: 'a', B3: '1', C3: '=B3*2', D3: '=$B$2+1', E3: '=B4', F3: '=A3&"!"',
    A4: 'b', B4: '2', C4: '=B4*2', D4: '=B$2', E4: '=B5', F4: '=A4&"!"',
    B6: '=SUM(B2:B4)', H2: '=B2', H3: '=C3', H4: '=SUM(B2:B3)',
  },
  formats: [{ range: 'A2:B2', bold: true }, { range: 'B3', color: 65535, numFmt: '0.00' }, { range: 'C4', border: true }],
  sheet2: { A1: '=Sheet1!B2' },
  names: { First: '=Sheet1!$B$2' },
};
add('formulas and formats move with their rows', sheet, [{ do: 'sort', range: 'A1:F4', keys: [{ key: 'A1' }], header: 'yes' }], 'A1:H7');
add('sorted by a formula column', sheet, [{ do: 'sort', range: 'A1:F4', keys: [{ key: 'F1', order: 'desc' }], header: 'yes' }], 'A1:H7');

// Several keys, left to right.
const team = { A1: 'Team', B1: 'Score', C1: 'Row', A2: 'red', B2: '5', C2: 'r2', A3: 'blue', B3: '7', C3: 'r3', A4: 'red', B4: '9', C4: 'r4', A5: 'blue', B5: '7', C5: 'r5', A6: 'red', B6: '5', C6: 'r6', A7: 'green', B7: '1', C7: 'r7' };
add('two keys: team A to Z, score largest first', { cells: team }, [{ do: 'sort', range: 'A1:C7', keys: [{ key: 'A1' }, { key: 'B1', order: 'desc' }], header: 'yes' }], 'A1:C8');
add('three keys', { cells: team }, [{ do: 'sort', range: 'A1:C7', keys: [{ key: 'B1' }, { key: 'A1', order: 'desc' }, { key: 'C1', order: 'desc' }], header: 'yes' }], 'A1:C8');
add('left to right', { cells: { A1: 'Key', B1: '3', C1: 'b', D1: '1', E1: 'a', A2: 'Row', B2: 'c2', C2: 'c3', D2: 'c4', E2: 'c5' } }, [{ do: 'sort', range: 'B1:E2', keys: [{ key: 'B1' }], header: 'no', leftToRight: true }], 'A1:F3');

// The range around one cell (CurrentRegion), as the A to Z button on one cell sorts it.
const region = { B3: 'Name', C3: 'Qty', B4: 'c', C4: '3', B5: 'a', C5: '1', B6: 'b', C6: '2', B8: 'z', C8: '0', E4: 'apart', E5: 'kept', A1: 'title' };
add('region around one cell', { cells: region }, [{ do: 'sort', range: 'C5', region: true, keys: [{ key: 'B5' }], legacy: true }], 'A1:F9');
add('region with a blank cell inside', { cells: { A1: 'Name', B1: 'Qty', A2: 'c', B2: '3', A3: 'a', A4: 'b', B4: '2' } }, [{ do: 'sort', range: 'A1', region: true, keys: [{ key: 'B1' }], legacy: true }], 'A1:C5');
add('region touching diagonally', { cells: { A1: 'Name', B1: 'Qty', A2: 'c', B2: '3', A3: 'a', B3: '1', C4: 'd', D4: 'x', A5: 'q' } }, [{ do: 'sort', range: 'A2', region: true, keys: [{ key: 'A1' }], legacy: true }], 'A1:E6');
add('blank rows in the range go last', { cells: { A1: 'Name', B1: 'Qty', A2: 'c', B2: '3', A4: 'a', B4: '1', A6: 'b', B6: '2' } }, [{ do: 'sort', range: 'A1:B7', keys: [{ key: 'A1' }], header: 'yes' }], 'A1:B8');

// Refusals.
add('merged cells of different sizes', { cells: { A1: 'Name', A2: 'c', A3: 'a', A4: 'b', B2: 'x' }, merges: ['B2:C2'] }, [{ do: 'sort', range: 'A1:C4', keys: [{ key: 'A1' }], header: 'yes' }], 'A1:C5');
add('merged cells all the same size', { cells: { A1: 'Name', A2: 'c', A3: 'a', A4: 'b', B2: 'x2', B3: 'x3', B4: 'x4' }, merges: ['B2:C2', 'B3:C3', 'B4:C4'] }, [{ do: 'sort', range: 'A2:C4', keys: [{ key: 'A2' }], header: 'no' }], 'A1:C5');
add('every cell in a merged range of one size', { cells: { A2: 'c', A3: 'a', A4: 'b' }, merges: ['A2:B2', 'A3:B3', 'A4:B4'] }, [{ do: 'sort', range: 'A2:B4', keys: [{ key: 'A2' }], header: 'no' }], 'A1:C5');
add('a merged range outside the sorted rows', { cells: { A1: 'Name', A2: 'c', A3: 'a', A4: 'b', A6: 'm' }, merges: ['A6:B6'] }, [{ do: 'sort', range: 'A1:B4', keys: [{ key: 'A1' }], header: 'yes' }], 'A1:C7');
add('part of an array', { cells: { A1: 'N', A2: '3', A3: '1', A4: '2' }, arrays: { 'B2:B3': '=A2:A3*2' } }, [{ do: 'sort', range: 'A1:B4', keys: [{ key: 'A1' }], header: 'yes' }], 'A1:B5');

process.stdout.write(JSON.stringify(cases, null, 1) + '\n');
