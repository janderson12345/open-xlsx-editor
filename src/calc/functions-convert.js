// CONVERT: a number from one unit to another of the same kind, with Excel's units and factors. Each factor is Excel's
// own, read to the last bit from CONVERT(1,unit,base) in a hidden Excel; which units take a prefix (k, m, ki, ...) was
// measured the same way. Registered into functions.js's table; evaluate.js imports this file.
import { E, textToNumber } from './values.js';
import { lifted, one } from './functions.js';

// Each kind: its base unit first, then units as "name factor" (several names for one factor: "cwt|shweight"); a * marks
// the units that take a prefix. Area and volume units carry their power (2 or 3), which a prefix is raised to.
const KINDS = {
  mass: `g*:1 sg:14593.902937206363 lbm:453.59237 u*:1.660538782E-24 ozm:28.349523125 grain:0.06479891
    cwt|shweight:45359.237 uk_cwt|lcwt:50802.34544 stone:6350.29318 ton:907184.74 uk_ton|LTON|brton:1016046.9088`,
  distance: `m*:1 mi:1609.344 Nmi:1852 in:0.0254 ft:0.3048 yd:0.9144 ang*:1E-10 ell:1.143 ly*:9460730472580800
    parsec*|pc*:30856775812815532 Pica|Picapt:0.0003527777777777778 pica:0.004233333333333334
    survey_mi:1609.3472186944373`,
  time: 'sec*|s*:1 yr:31557600 day|d:86400 hr:3600 mn|min:60',
  pressure: 'Pa*|p*:1 atm*|at*:101325 mmHg*:133.322 psi:6894.757293168362 Torr:133.32236842105263',
  force: 'N*:1 dyn*|dy*:1E-05 lbf:4.4482216152605005 pond*:0.00980665',
  energy: `J*:1 e*:1E-07 c*:4.184 cal*:4.1868 eV*|ev*:1.602176487E-19 HPh|hh:2684519.537696173 Wh*|wh*:3600
    flb:1.3558179483314003 BTU|btu:1055.05585262`,
  power: 'W*|w*:1 HP|h:745.6998715822702 PS:735.49875',
  magnetism: 'T*:1 ga*:0.0001',
  temperature: 'K*|kel*:1 C|cel:0 F|fah:0 Rank:0 Reau:0',
  volume: `l*|L*|lt*:1 tsp:0.00492892159375 tspm:0.005 tbs:0.01478676478125 oz:0.0295735295625 cup:0.2365882365
    pt|us_pt:0.473176473 uk_pt*:0.56826125 qt:0.946352946 uk_qt:1.1365225 gal:3.785411784 uk_gal:4.54609
    ang3*|ang^3*:1E-27 barrel:158.987294928 bushel:35.23907016688 ft3|ft^3:28.316846592 in3|in^3:0.016387064
    ly3|ly^3:8.467866646237152E+50 m3*|m^3*:1000 mi3|mi^3:4168181825440.5796 yd3|yd^3:764.554857984
    Nmi3|Nmi^3:6352182208000 Pica3|Pica^3|Picapt3|Picapt^3:4.39039566186557E-08 GRT|regton:2831.6846592
    MTON:1132.67386368`,
  area: `m2*|m^2*:1 uk_acre:4046.8564224 us_acre:4046.8726098742522 ang2*|ang^2*:1E-20 ar*:100 ft2|ft^2:0.09290304
    ha:10000 in2|in^2:0.00064516 ly2|ly^2:8.95054210748189E+31 Morgen:2500 mi2|mi^2:2589988.110336
    Nmi2|Nmi^2:3429904 Pica2|Pica^2|Picapt2|Picapt^2:1.2445216049382715E-07 yd2|yd^2:0.83612736`,
  information: 'bit*:1 byte*:8',
  speed: 'm/s*|m/sec*:1 m/h*|m/hr*:0.0002777777777777778 mph*:0.44704 kn:0.5144444444444445 admkn:0.5147733333333333',
};
/** name -> { kind, factor, prefixed, power } */
const UNITS = new Map();
for (const [kind, list] of Object.entries(KINDS)) {
  for (const entry of list.trim().split(/\s+/)) {
    const [names, factor] = entry.split(':');
    for (const n of names.split('|')) {
      const name = n.replace(/\*$/, '');
      const power = kind === 'area' && /2$/.test(name) ? 2 : kind === 'volume' && /3$/.test(name) ? 3 : 1;
      UNITS.set(name, { kind, factor: Number(factor), prefixed: n.endsWith('*'), power });
    }
  }
}
// (A prefix is squared or cubed for a length squared or cubed, "km2" 1E+06 m2, and not otherwise: "kar" is 1000 ar, "cl" 0.01 l.)
const PREFIXES = {
  Y: 1e24, Z: 1e21, E: 1e18, P: 1e15, T: 1e12, G: 1e9, M: 1e6, k: 1e3, h: 1e2, da: 1e1, e: 1e1, d: 1e-1, c: 1e-2,
  m: 1e-3, u: 1e-6, n: 1e-9, p: 1e-12, f: 1e-15, a: 1e-18, z: 1e-21, y: 1e-24,
};
/** The binary prefixes, for bits and bytes only. */
const BINARY = { Yi: 2 ** 80, Zi: 2 ** 70, Ei: 2 ** 60, Pi: 2 ** 50, Ti: 2 ** 40, Gi: 2 ** 30, Mi: 2 ** 20, ki: 2 ** 10 };

/** A unit's name as { kind, scale, base }: a unit as named, or one that takes a prefix with one (one only: "kmm" is
 *  #N/A). Names are exact: "M", "FT" and " m" are #N/A. */
function unit(name) {
  const plain = UNITS.get(name);
  if (plain) return { ...plain, scale: plain.factor, name };
  for (const [table, binary] of [[PREFIXES, false], [BINARY, true]]) {
    for (const p of Object.keys(table)) {
      if (!name.startsWith(p)) continue;
      const rest = name.slice(p.length);
      const u = UNITS.get(rest);
      if (!u || !u.prefixed || (binary && u.kind !== 'information')) continue;
      return { ...u, scale: table[p] ** u.power * u.factor, name: rest };
    }
  }
  throw E.na;
}

/** Kelvin from a temperature in a unit, and back. */
const TO_K = { C: (x) => x + 273.15, cel: (x) => x + 273.15, F: (x) => (x - 32) / 1.8 + 273.15, fah: (x) => (x - 32) / 1.8 + 273.15,
  Rank: (x) => x / 1.8, Reau: (x) => x * 1.25 + 273.15 };
const FROM_K = { C: (k) => k - 273.15, cel: (k) => k - 273.15, F: (k) => (k - 273.15) * 1.8 + 32, fah: (k) => (k - 273.15) * 1.8 + 32,
  Rank: (k) => k * 1.8, Reau: (k) => (k - 273.15) / 1.25 };

lifted('CONVERT', (a, ctx) => {
  const v = one(a[0], ctx);
  if (typeof v === 'boolean') throw E.value;
  const x = typeof v === 'number' ? v : v == null ? 0 : textToNumber(String(v));
  if (x == null) throw E.value;
  const from = unit(String(one(a[1], ctx) ?? ''));
  const to = unit(String(one(a[2], ctx) ?? ''));
  if (from.kind !== to.kind) throw E.na;
  if (from.kind === 'temperature') {
    const k = TO_K[from.name] ? TO_K[from.name](x) : x * from.scale;
    return TO_K[to.name] ? FROM_K[to.name](k) : k / to.scale;
  }
  return x * from.scale / to.scale;
});
