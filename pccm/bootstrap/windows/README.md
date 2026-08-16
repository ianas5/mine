# PCCM Stage B bootstrap (Windows / Excel COM)

Production bootstrap that will take the Stage A skeleton and produce the final
`.xlsm`: importing VBA from `pccm/src/vba/`, assigning worksheet CodeNames,
creating Excel-runtime-only objects, then saving, reopening and verifying.

**Empty in Phase 1.** Nothing here is written yet.

This directory is deliberately separate from `pccm/readiness/windows/`, which
holds the disposable Excel COM smoke test. That readiness gate is closed and
passed on the target machine; its script is a throwaway diagnostic, not
production build code, and the two must not be mixed.
