const DROP = /[ـً-ٰٟ​-‏‪-‮⁦-⁩﻿]/;
const SPACE = /[\s ]/;
const MAP: Record<string, string> = {
  'أ': 'ا', 'إ': 'ا', 'آ': 'ا', 'ٱ': 'ا',
  'ى': 'ي', 'ی': 'ي', 'ة': 'ه', 'ک': 'ك',
  '–': '-', '—': '-', '−': '-',
};

/**
 * Normalise Arabic text for matching (brief §4): drop tatweel/diacritics, unify
 * alef / ya / ta-marbuta, collapse whitespace. `map[i]` is the index in the
 * original string of normalised character `i`, so matches can be traced back
 * to the original runs (underline).
 */
export function normalizeWithMap(s: string): { text: string; map: number[] } {
  let text = '';
  const map: number[] = [];
  let pendingSpace = -1;
  for (let i = 0; i < s.length; i++) {
    let ch = s[i];
    if (DROP.test(ch)) continue;
    if (SPACE.test(ch)) {
      if (text.length) pendingSpace = i;
      continue;
    }
    const code = ch.charCodeAt(0);
    if (code >= 0x0660 && code <= 0x0669) ch = String(code - 0x0660);
    else ch = MAP[ch] ?? ch;
    if (pendingSpace >= 0) {
      text += ' ';
      map.push(pendingSpace);
      pendingSpace = -1;
    }
    text += ch;
    map.push(i);
  }
  return { text, map };
}

export function normalize(s: string): string {
  return normalizeWithMap(s).text;
}

/** Letters only — the key used to compare names (`عبد الرحيم` = `عبدالرحيم`). */
export function nameKey(s: string): string {
  return normalize(s).replace(/[^\p{L}]/gu, '');
}

/** Text as written but tidy: no tatweel, single spaces. */
export function tidy(s: string): string {
  return s.replace(/[ـ​-‏‪-‮﻿]/g, '').replace(/[\s ]+/g, ' ').trim();
}
