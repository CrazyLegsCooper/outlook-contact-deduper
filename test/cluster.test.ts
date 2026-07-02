import { describe, it, expect } from 'vitest';
import { clusterPairs } from '../src/engine/cluster';

describe('clusterPairs', () => {
  it('merges transitive pairs into one group', () => {
    const groups = clusterPairs([
      { aId: 'a', bId: 'b' },
      { aId: 'b', bId: 'c' },
    ]);
    expect(groups).toHaveLength(1);
    expect([...groups[0]].sort()).toEqual(['a', 'b', 'c']);
  });

  it('keeps disjoint pairs as separate groups', () => {
    const groups = clusterPairs([
      { aId: 'a', bId: 'b' },
      { aId: 'x', bId: 'y' },
    ]);
    expect(groups).toHaveLength(2);
  });
});
