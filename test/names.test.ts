import { describe, it, expect } from 'vitest';
import { levenshtein, similarityRatio, namesMatchFuzzy } from '../src/engine/names';

describe('levenshtein', () => {
  it('counts single edits', () => {
    expect(levenshtein('kitten', 'sitting')).toBe(3);
    expect(levenshtein('abc', 'abc')).toBe(0);
  });
});

describe('similarityRatio', () => {
  it('is 1 for identical strings', () => {
    expect(similarityRatio('john', 'john')).toBe(1);
  });
  it('is between 0 and 1 for near matches', () => {
    expect(similarityRatio('jon', 'john')).toBeGreaterThan(0.7);
  });
});

describe('namesMatchFuzzy', () => {
  it('matches diminutive/prefix names with same surname', () => {
    expect(namesMatchFuzzy('chris cooper', 'christopher cooper')).toBe(true);
  });
  it('matches initial to full first name', () => {
    expect(namesMatchFuzzy('j smith', 'john smith')).toBe(true);
  });
  it('matches known nicknames', () => {
    expect(namesMatchFuzzy('bob jones', 'robert jones')).toBe(true);
  });
  it('tolerates a typo', () => {
    expect(namesMatchFuzzy('jonh smith', 'john smith')).toBe(true);
  });
  it('does NOT match different surnames', () => {
    expect(namesMatchFuzzy('chris cooper', 'chris baker')).toBe(false);
  });
  it('does NOT match unrelated first names with same surname', () => {
    expect(namesMatchFuzzy('john smith', 'peter smith')).toBe(false);
  });
  it('does NOT match distinct short first names that differ by one substitution', () => {
    expect(namesMatchFuzzy('joan smith', 'john smith')).toBe(false);
    expect(namesMatchFuzzy('mark cooper', 'mary cooper')).toBe(false);
    expect(namesMatchFuzzy('anna lee', 'anne lee')).toBe(false);
  });
  it('returns false when either name is empty', () => {
    expect(namesMatchFuzzy('', 'john smith')).toBe(false);
  });
});
