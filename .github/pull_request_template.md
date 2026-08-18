## Summary

Describe what changed, why it changed, and the user or operator impact.

## Root cause

For fixes, explain the underlying cause rather than only the visible symptom.

## Verification

- [ ] Ruff and compileall pass.
- [ ] Source smoke and JavaScript checks pass.
- [ ] Relevant live/browser/terminal flows pass.
- [ ] Dependency locks and audits are current when dependencies changed.
- [ ] Full Windows x64 build passes when packaging changed.
- [ ] Performance-sensitive changes include before/after measurements.
- [ ] No secret, access code, token, PFX, or real `.env` value is included.

## Compatibility and risk

List public-interface, settings-schema, deployment, hardware, and rollback considerations. State
anything that remains unverified.
