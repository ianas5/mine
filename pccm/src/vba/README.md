# PCCM VBA source

Authoritative source for every VBA module, class module and document module in
the production workbook. These are plain text files under version control; the
`.xlsm` is a build artifact that Stage B assembles from them.

**Empty in Phase 1.** No production VBA has been written yet — see
`pccm/docs/phase1.md` for what Phase 1 does and does not cover.

Planned layout (later phases):

    modConstants.bas, modAppState.bas, modStructure.bas, ...
    clsRng.cls, clsRunContext.cls, ...
    doc/ThisWorkbook.txt, doc/shDashboard.txt, ...

Document-module code lives under `doc/` because it is *injected* into an
existing component rather than imported as a new one.
