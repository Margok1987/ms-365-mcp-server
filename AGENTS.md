# Repository engineering instructions

## Mandatory Ponytail-full method (Owner decision 2026-10-11)

For every future change to code, tests, scripts, configuration, builds, dependencies, or refactoring in this repository, apply the **Ponytail full** implementation method (source: https://github.com/DietrichGebert/ponytail; do not auto-install an upstream plugin).

- First inspect the actual execution path and all affected callers, tests, fixtures, configuration, exports, and user/data risks.
- Prefer the **smallest complete, readable change**. Avoid unnecessary features; reuse existing project code, then standard-library/native facilities, then installed dependencies. Add wrappers, abstractions, packages, options, or services only for a demonstrated requirement.
- Fix shared root causes once, preserve existing conventions, and change all affected paths. Keep validation, meaningful errors, recovery, and security controls intact. Test nontrivial logic and run every existing required integration/regression check.
- Never shrink explicitly requested behavior, supported input/response combinations, safety, permissions, live evidence, verification, deployment/CI gates, or release acceptance just to reduce code. Dedicated repository/domain instructions remain authoritative; Ponytail is their compatible simplicity discipline.
- Do not refactor untouched code or introduce a runtime dependency merely to adopt this method.
