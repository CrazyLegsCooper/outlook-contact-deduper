import type { Contact } from '../engine/types';

export function backupFilename(d: Date): string {
  const iso = d.toISOString().slice(0, 10);
  return `contacts-backup-${iso}.json`;
}

export function downloadBackup(contacts: Contact[], d: Date = new Date()): void {
  const blob = new Blob([JSON.stringify(contacts, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = backupFilename(d);
  a.click();
  URL.revokeObjectURL(url);
}
