import { useEffect } from 'react';

import { photoRepository } from '@/data/repositories/photoRepository';

/**
 * Runs the progress-photo orphan sweep once on launch (DATABASE §3.6): deletes
 * files with no metadata row (e.g. a kill between file-write and row-insert), so
 * the filesystem and the database reconcile to a single consistent state. Renders
 * nothing. Failures are swallowed — a sweep is best-effort housekeeping.
 */
export function PhotoSweeper(): null {
  useEffect(() => {
    void photoRepository.sweepOrphans().catch(() => undefined);
  }, []);
  return null;
}
