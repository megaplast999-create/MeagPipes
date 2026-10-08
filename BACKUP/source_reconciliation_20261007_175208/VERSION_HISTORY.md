# Version History

| Phase | Date | Version | Changes | Tests | Known issues |
|---|---|---|---|---|---|
| Phase 0 | 2026-10-07 | DEV-0.1 | Created the initial structural workbook and project folders. | Initial workbook package inspection. | No VBA project. |
| Phase 1 | 2026-10-07 | DEV-0.1 | Added workbook architecture, required tables, VBA folders, and architecture documentation. | `phase1_architecture_check.py` passed. | No VBA project. |
| Phases 2-4 foundation | 2026-10-07 | DEV-0.2 | Seeded documented master data and settings; added reproducible workbook builder and independent calculation reference layer. | Structural validation passed; calculation tests pending runner. | VBA embedding and Excel runtime validation are blocked/pending. |
| Header and quote-line source stage | 2026-10-07 | DEV-0.3 | Audited the manually embedded workbook, documented current state, and added source modules for master-driven header validation, customer lookup, local DEV sequencing, and validated quote-line insertion. | Python workbook/runtime audit and calculation tests. | New source modules require import/reconciliation; Excel runtime validation pending. |
