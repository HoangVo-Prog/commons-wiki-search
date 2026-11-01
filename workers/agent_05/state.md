# Agent 05 — State

- **Status:** ACTIVE (reconcile with `record.py status`)
- **Agent new images collected:** not yet reconciled; `progress.json` is a cache
- **Quota:** read the `agent_quota` from config.json; never compute it from a global total
- **Current keyword:** Taipei Taiwan busy street
- **Search page:** API offset=0
- **Next action:** Run python project.py manifest --agent-id 05 then python commons_api.py collect --agent-id 05 --max-pages 8 --max-accept 25

**Recovery:** Preserve paired images.xlsx + collection_state.json; run `python record.py status --agent-id 05`. Never count historical workbook rows toward your quota.
