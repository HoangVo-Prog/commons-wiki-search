# Worker 07 — independent API-only collection

AGENT_ID is **07**. Use `python project.py show` and `python project.py manifest --agent-id 07` to locate root and *your own* worker instructions, keyword pool, cursor, progress, state and recent log. Never read or modify another worker's directory.

Your `agent_quota` is controlled **only by config.json**. `agent_new_images_collected` counts only rows added after the persistent `collection_state.json` baseline; historical images remain globally deduplicated but do NOT count. Stop when **`agent_new_images_collected >= agent_quota`**, never due to any global count, workbook size or other agents' progress.

On resume, run `python record.py status --agent-id 07` to reconcile real new accepted rows from the paired workbook + baseline. Never delete/reset the baseline, replace the workbook or double-claim an ID. Use `python commons_api.py collect --agent-id 07 --max-pages 8 --max-accept 25` in batches, using your large keyword pool and persistent API cursor. Batch sizes are not quotas.

Use Wikimedia Commons API metadata only; no image downloads or fabricated URLs, no claim of visual verification. Save via locked record.py exclusively and honor shared API throttling.
