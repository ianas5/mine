import { useEffect, useState } from 'react';

import { useTableVersion } from '@/core/db';
import type { Program } from '@/domain/models';
import { programRepository } from '@/data/repositories/programRepository';

/** All non-archived programs (each with templates), reactive to program writes. */
export function usePrograms(): Program[] | undefined {
  const version = useTableVersion('programs');
  const [programs, setPrograms] = useState<Program[] | undefined>(undefined);

  useEffect(() => {
    let live = true;
    void programRepository.listPrograms().then((rows) => {
      if (live) setPrograms(rows);
    });
    return () => {
      live = false;
    };
  }, [version]);

  return programs;
}

/** One program by id (with templates), reactive. `undefined` loading, `null` missing. */
export function useProgram(id: string): Program | null | undefined {
  const version = useTableVersion('programs');
  const [program, setProgram] = useState<Program | null | undefined>(undefined);

  useEffect(() => {
    let live = true;
    void programRepository.getProgram(id).then((row) => {
      if (live) setProgram(row);
    });
    return () => {
      live = false;
    };
  }, [id, version]);

  return program;
}
