import { describe, it, expect } from 'vitest';
import { backupFilename } from '../src/graph/backup';

describe('backupFilename', () => {
  it('formats the date as YYYY-MM-DD', () => {
    expect(backupFilename(new Date('2026-07-01T10:20:30Z'))).toBe('contacts-backup-2026-07-01.json');
  });
});
