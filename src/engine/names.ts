export function levenshtein(a: string, b: string): number {
  const m = a.length, n = b.length;
  if (m === 0) return n;
  if (n === 0) return m;
  const prev = new Array(n + 1);
  const curr = new Array(n + 1);
  for (let j = 0; j <= n; j++) prev[j] = j;
  for (let i = 1; i <= m; i++) {
    curr[0] = i;
    for (let j = 1; j <= n; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      curr[j] = Math.min(prev[j] + 1, curr[j - 1] + 1, prev[j - 1] + cost);
    }
    for (let j = 0; j <= n; j++) prev[j] = curr[j];
  }
  return prev[n];
}

export function similarityRatio(a: string, b: string): number {
  if (a === b) return 1;
  const maxLen = Math.max(a.length, b.length);
  if (maxLen === 0) return 1;
  return 1 - levenshtein(a, b) / maxLen;
}

// Known nickname groups (normalized, lowercase). Extend as needed.
const NICKNAME_GROUPS: string[][] = [
  ['robert', 'rob', 'bob', 'bobby'],
  ['william', 'will', 'bill', 'billy'],
  ['richard', 'rich', 'rick', 'dick'],
  ['margaret', 'maggie', 'peggy', 'meg'],
  ['elizabeth', 'liz', 'beth', 'betty', 'eliza'],
  ['james', 'jim', 'jimmy'],
  ['john', 'jack', 'johnny'],
  ['michael', 'mike', 'mick'],
  ['edward', 'ed', 'eddie', 'ted'],
  ['charles', 'charlie', 'chuck'],
];

function sameNicknameGroup(a: string, b: string): boolean {
  return NICKNAME_GROUPS.some((g) => g.includes(a) && g.includes(b));
}

function firstNamesMatch(a: string, b: string): boolean {
  if (!a || !b) return true; // missing first name should not block a surname match
  if (a === b) return true;
  if ((a.length === 1 && b.startsWith(a)) || (b.length === 1 && a.startsWith(b))) return true;
  const [short, long] = a.length <= b.length ? [a, b] : [b, a];
  if (short.length >= 3 && long.startsWith(short)) return true; // chris -> christopher
  if (sameNicknameGroup(a, b)) return true;
  return similarityRatio(a, b) >= 0.5;
}

function tokens(name: string): string[] {
  return name.split(' ').filter(Boolean);
}

export function namesMatchFuzzy(a: string, b: string): boolean {
  const ta = tokens(a), tb = tokens(b);
  if (ta.length === 0 || tb.length === 0) return false;
  const aFirst = ta[0], aLast = ta[ta.length - 1];
  const bFirst = tb[0], bLast = tb[tb.length - 1];
  const lastOk =
    ta.length < 2 || tb.length < 2
      ? true
      : aLast === bLast || similarityRatio(aLast, bLast) >= 0.9;
  return lastOk && firstNamesMatch(aFirst, bFirst);
}
