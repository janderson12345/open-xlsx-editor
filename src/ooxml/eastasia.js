// The East Asian built-in number formats (ids 27-36 and 50-58): each means a different code in Japanese, Chinese
// (Simplified and Traditional) and Korean Excel, and a file gives only the id. The codes are ECMA-376 Part 1's table
// (18.8.30, numFmt); how Excel draws them (eras, the Republic of China's years, 上午/下午) was measured from Excel
// 365 over COM (tools/excel-compare/excel-eastasia.ps1, test/eastasia-excel.json).

const JA = {
  27: '[$-411]ge.m.d', 28: '[$-411]ggge"年"m"月"d"日"', 29: '[$-411]ggge"年"m"月"d"日"', 30: 'm/d/yy',
  31: 'yyyy"年"m"月"d"日"', 32: 'h"時"mm"分"', 33: 'h"時"mm"分"ss"秒"', 34: 'yyyy"年"m"月"', 35: 'm"月"d"日"',
  36: '[$-411]ge.m.d', 50: '[$-411]ge.m.d', 51: '[$-411]ggge"年"m"月"d"日"', 52: 'yyyy"年"m"月"', 53: 'm"月"d"日"',
  54: '[$-411]ggge"年"m"月"d"日"', 55: 'yyyy"年"m"月"', 56: 'm"月"d"日"', 57: '[$-411]ge.m.d',
  58: '[$-411]ggge"年"m"月"d"日"',
};
const ZH_CN = {
  27: 'yyyy"年"m"月"', 28: 'm"月"d"日"', 29: 'm"月"d"日"', 30: 'm-d-yy', 31: 'yyyy"年"m"月"d"日"', 32: 'h"时"mm"分"',
  33: 'h"时"mm"分"ss"秒"', 34: '上午/下午h"时"mm"分"', 35: '上午/下午h"时"mm"分"ss"秒"', 36: 'yyyy"年"m"月"',
  50: 'yyyy"年"m"月"', 51: 'm"月"d"日"', 52: 'yyyy"年"m"月"', 53: 'm"月"d"日"', 54: 'm"月"d"日"',
  55: '上午/下午h"时"mm"分"', 56: '上午/下午h"时"mm"分"ss"秒"', 57: 'yyyy"年"m"月"', 58: 'm"月"d"日"',
};
const ZH_TW = {
  27: '[$-404]e/m/d', 28: '[$-404]e"年"m"月"d"日"', 29: '[$-404]e"年"m"月"d"日"', 30: 'm/d/yy', 31: 'yyyy"年"m"月"d"日"',
  32: 'hh"時"mm"分"', 33: 'hh"時"mm"分"ss"秒"', 34: '上午/下午hh"時"mm"分"', 35: '上午/下午hh"時"mm"分"ss"秒"',
  36: '[$-404]e/m/d', 50: '[$-404]e/m/d', 51: '[$-404]e"年"m"月"d"日"', 52: '上午/下午hh"時"mm"分"',
  53: '上午/下午hh"時"mm"分"ss"秒"', 54: '[$-404]e"年"m"月"d"日"', 55: '上午/下午hh"時"mm"分"',
  56: '上午/下午hh"時"mm"分"ss"秒"', 57: '[$-404]e/m/d', 58: '[$-404]e"年"m"月"d"日"',
};
const KO = {
  27: 'yyyy"年" mm"月" dd"日"', 28: 'mm-dd', 29: 'mm-dd', 30: 'mm-dd-yy', 31: 'yyyy"년" mm"월" dd"일"',
  32: 'h"시" mm"분"', 33: 'h"시" mm"분" ss"초"', 34: 'yyyy-mm-dd', 35: 'yyyy-mm-dd', 36: 'yyyy"年" mm"月" dd"日"',
  50: 'yyyy"年" mm"月" dd"日"', 51: 'mm-dd', 52: 'yyyy-mm-dd', 53: 'yyyy-mm-dd', 54: 'mm-dd', 55: 'yyyy-mm-dd',
  56: 'yyyy-mm-dd', 57: 'yyyy"年" mm"月" dd"日"', 58: 'mm-dd',
};

/** The codes by language: ja, zh-CN (also zh-SG, zh-Hans), zh-TW (also zh-HK, zh-MO, zh-Hant), ko. */
export const EAST_ASIAN_FORMATS = { ja: JA, 'zh-CN': ZH_CN, 'zh-TW': ZH_TW, ko: KO };

/** The ids that are East Asian built-ins. */
export const EAST_ASIAN_IDS = Object.keys(JA).map(Number);

/** Which table a locale tag reads (null for one that isn't Japanese, Chinese or Korean). */
export function eastAsianTable(tag) {
  const t = String(tag || '').toLowerCase();
  if (t.startsWith('ja')) return JA;
  if (t.startsWith('ko')) return KO;
  if (/^zh-(tw|hk|mo|hant)/.test(t)) return ZH_TW;
  if (t.startsWith('zh')) return ZH_CN;
  return null;
}
