// The function screen tip, as Excel shows it under a cell being edited: inside a function's parentheses, its syntax
// ("SUM(number1, [number2], ...)") with the argument the caret is in in bold. Only for the functions Insert Function
// lists (calc/catalog.js).
import { functionInfo } from '../calc/catalog.js';

/** The function call the caret at `caret` is inside in formula `text`, innermost first: { name, arg } (arg from 0),
 *  or null. Text in quotes, sheet names in single quotes and array constants are passed over. */
export function callAt(text, caret) {
  if (!text.startsWith('=')) return null;
  const stack = [];
  for (let i = 1; i < caret; i++) {
    const ch = text[i];
    if (ch === '"' || ch === "'") {
      // (A doubled quote is a quote inside: the string goes on.)
      let j = i + 1;
      while (j < text.length && !(text[j] === ch && text[j + 1] !== ch)) j += text[j] === ch ? 2 : 1;
      if (j >= caret) return null;
      i = j;
    } else if (ch === '(') {
      const name = /([A-Za-z_][A-Za-z0-9_.]*)$/.exec(text.slice(1, i))?.[1] || null;
      stack.push({ name: name ? name.replace(/^_xlfn\.(_xlws\.)?/i, '').toUpperCase() : null, arg: 0 });
    } else if (ch === '{') stack.push({ name: null, arg: 0, array: true });
    else if (ch === ')' || ch === '}') stack.pop();
    else if (ch === ',' && stack.length) stack[stack.length - 1].arg += 1;
  }
  for (let k = stack.length - 1; k >= 0; k--) if (stack[k].array) return null; else if (stack[k].name) return { name: stack[k].name, arg: stack[k].arg };
  return null;
}

/** Which of a function's arguments (Excel's names) the `arg`th one typed is: repeated ones ("number1, [number2],
 *  ...") go on with the last named. */
export function argIndex(args, arg) {
  if (arg < args.length && args[arg] !== '...') return arg;
  const dots = args.indexOf('...');
  if (dots < 0) return -1;
  // ("value1, [value2], ...": past the list, the last named one; IFS's pairs go on two by two.)
  const named = args.slice(0, dots);
  const pair = named.length >= 2 && /1\]?$/.test(named[named.length - 2]) && /1\]?$/.test(named[named.length - 1]) ? 2 : 1;
  return named.length - pair + ((arg - named.length) % pair + pair) % pair;
}

/** The tip under `textarea` (in `overlay`): shown, updated and hidden as the text and caret change. */
export function attachFunctionTip(textarea, overlay) {
  const tip = document.createElement('div');
  tip.className = 'oxe-fntip';
  tip.setAttribute('role', 'tooltip');
  tip.id = `oxe-fntip-${Math.random().toString(36).slice(2, 8)}`;
  const update = () => {
    const call = textarea.isConnected ? callAt(textarea.value, textarea.selectionStart) : null;
    const f = call && functionInfo(call.name);
    if (!f) { tip.remove(); textarea.removeAttribute('aria-describedby'); return; }
    const at = argIndex(f.args, call.arg);
    tip.replaceChildren(`${f.name}(`, ...f.args.flatMap((a, k) => {
      const part = k === at ? Object.assign(document.createElement('b'), { textContent: a }) : a;
      return k ? [', ', part] : [part];
    }), ')');
    tip.style.left = textarea.style.left || `${textarea.offsetLeft}px`;
    tip.style.top = `${textarea.offsetTop + textarea.offsetHeight + 2}px`;
    if (!tip.isConnected) overlay.appendChild(tip);
    textarea.setAttribute('aria-describedby', tip.id);
  };
  for (const ev of ['input', 'keyup', 'click', 'focus']) textarea.addEventListener(ev, update);
  textarea.addEventListener('blur', () => tip.remove());
  return { update, el: tip };
}
