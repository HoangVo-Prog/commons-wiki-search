# Agent 02 — State

- **Status:** ACTIVE (reconcile with `record.py status`)
- **Agent new images collected:** not yet reconciled; `progress.json` is a cache
- **Quota:** read the `agent_quota` from config.json; never compute it from a global total
- **Current keyword:** Osaka busy street
- **Search page:** API offset=0
- **Next action:** Run python project.py manifest --agent-id 02 then python commons_api.py collect --agent-id 02 --max-pages 8 --max-accept 25

**Recovery:** Preserve paired images.xlsx + collection_state.json; run `python record.py status --agent-id 02`. Never count historical workbook rows toward your quota.
