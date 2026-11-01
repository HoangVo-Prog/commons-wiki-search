# HRVS Collection — fresh independent quotas per run

Ten API-only workers collect unique high-resolution Wikimedia Commons **original image URLs** into a shared `images.xlsx`. No images are downloaded. Each agent has its own `agent_quota` in `config.json`; there is **no global target or limit**.

## Fresh run semantics
At the start of each **new collection run**, `record.py` snapshots each agent's existing image count in `images.xlsx`:

`initial_count = existing records belonging to this agent`  
`target_count = initial_count + agent_quota`  
`agent_new_images_collected = records appended by this agent since this run began`  
`current_count = initial_count + agent_new_images_collected`

**Stop only when `agent_new_images_collected >= agent_quota`.** Example: 1,000 existing images and a 500-image config quota means **1000/1500** at initialization, then 1001/1500 after one accepted new URL, and completion at 1500/1500. Every run gives a fresh quota, regardless of historical rows from the same agent.

## Initialization, resume and creating another run

1. **Never replace your actual `images.xlsx`.** Place your populated workbook at the path configured by `paths.workbook`. Use its matching `collection_state.json` if one already exists; these are a pair and must be preserved across backups/deployments.
2. `python record.py status --all` creates the initial run snapshot if none exists, or **resumes the existing run** without resetting counts. On a pristine project, `python record.py start-run` also initializes run 1. The snapshot is serialized under the workbook lock.
3. `python record.py status --agent-id 01` reports both total progress (`initial_count/target_count` at start) and new unique images collected in the current run. Progress is reconstructed from workbook rows, NOT from cached `progress.json`.
4. Once all agents finish, explicitly run `python record.py start-run` to start the next fresh run. It snapshots the *entire existing workbook* and saves a new per-agent initial count/target in `collection_state.json`, while keeping all prior snapshots. To intentionally abandon an incomplete run, use `python record.py start-run --force`.
5. **Do not use `start-run` to resume.** Just run status and continue with the same agent and cursor. Only change the individual `agent_quota` settings in `config.json` between runs. A mid-run quota change is refused rather than silently altering a persisted target.

Example configuration (the values are editable, not hardcoded script quotas):

```json
"agent_quota": {
  "01": 500, "02": 500, "03": 500, "04": 500, "05": 500,
  "06": 500, "07": 500, "08": 500, "09": 500, "10": 500
}
```

## Commands

```powershell
python project.py show
python record.py status --all
python record.py status --agent-id 01
python record.py audit
python record.py start-run        # Only to open a NEW collection run
python commons_api.py test
python commons_api.py collect --agent-id 01 --max-pages 8 --max-accept 25
```

For independent Codex workers, use `PROMPT_FOR_CODEX.txt` or PowerShell launchers (`run_one_agent.ps1`, `run_all_agents.ps1`). Agent claims are unique and persistent. `project.py manifest --agent-id 01` resolves all worker and instruction paths from config. The worker search cursors/keyword pools persist for resume and subsequent runs; expand an exhausted agent's keywords instead of resetting its cursor.

## How consistency is maintained

- `collection_state.json` version 2 stores a **history of runs**, active run ID, per-agent `initial_count`, `agent_quota` and `target_count`, and a hash of the exact workbook prefix at each initialization. Version-1 collection state is migrated without reclassifying past collected images.
- Existing workbook URLs remain in the global deduplication set forever. No pre-existing rows are deleted, reordered or rewritten by initialization. Only accepting a qualified new unique record appends to the XLSX.
- OS lock serializes initialization, status, append and run transitions. Atomic XLSX saves avoid partial writes. State is written before a new run can collect; after an interrupted append, recovery recomputes new counts from rows beyond the saved run baseline.
- Each batch pins a run ID: a collector left running across a new-run boundary cannot append to the new run by mistake.
- There is no project-wide image cap; `--max-accept` and `--max-pages` control only individual API batches.

## Notes

Set a real `wikimedia_contact` (or `HRVS_WIKIMEDIA_CONTACT`) before API calls. Collect metadata candidates only, never image bytes. Complex visual content, authenticity and each Commons license require independent verification. Never reset `collection_state.json` to reuse a quota: start a new run explicitly. Run offline tests with `python -m unittest discover -s tests -v`.
