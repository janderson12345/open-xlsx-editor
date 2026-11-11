// The editor's small DOM helpers: CSS custom properties set in pixels, a div with a class (and a role).
export const px = (n) => `${n}px`;
export const setVars = (el, vars) => { for (const k in vars) el.style.setProperty(k, px(vars[k])); };
export const div = (cls, role) => {
  const el = document.createElement('div');
  el.className = cls;
  if (role) el.setAttribute('role', role);
  return el;
};
