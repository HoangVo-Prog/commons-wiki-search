# HRVS Wikimedia Commons — fresh per-run quotas (API only)

## Quota semantics (mandatory)
Each configured agent has its own `agent_quota` in `config.json`. There is NO global limit or derived per-agent cap. Every **new collection run** must snapshot existing `images.xlsx` rows by agent ID under the shared writer lock:

- `initial_count` = existing workbook image rows for this agent at **run initialization** (including images it collected in older runs).
- `agent_quota` = a **fresh number of additional unique images** taken from `config.json` and persisted for this run.
- `target_count = initial_count + agent_quota` (persisted per agent in `collection_state.json`).
- `current_count = initial_count + agent_new_images_collected`.
- The **only quota stopping condition** is `agent_new_images_collected >= agent_quota` for this run.

Example: if agent 01 has 1,000 existing images and `agent_quota` is 500, its **initial progress is 1000/1500**, and it must add 500 unique images before stopping. Pre-existing rows never use up the new quota. Other agents' counts and total workbook size never affect its quota.

## Run lifecycle (do not confuse a new run with resume)
1. Preserve your existing `images.xlsx` and `collection_state.json` together. Never overwrite a populated workbook with the template. Never hand-edit/remove persistent run state. Before first collection, run `python record.py status --all` to atomically create a run snapshot. Alternatively, on a pristine project, `python record.py start-run` creates the first run.
2. **Resume** the existing run with `python record.py status --agent-id YOUR_ID` or `--all`. This restores the *same* run's `initial_count`, `target_count`, and workbook-reconstructed new count; it NEVER resets a quota.
3. **Start a genuinely new run** only by explicitly calling `python record.py start-run`. Once every agent finished, it creates a new run and snapshots ALL currently existing rows as fresh initial counts. If the old run was unfinished, `python record.py start-run --force` intentionally abandons its unfinished quotas and starts another run. Do not invoke this during an ordinary resume.
4. `collection_state.json` retains each run's ID, initialized timestamp, baseline row count/hash, and per-agent `initial_count`, `agent_quota` and `target_count`. Preserve it with the workbook for migration/backup. Old version-1 baseline state is migrated preserving its original baseline and new rows.
5. Quotas are snapshotted from config at run initialization. Do not edit config quotas mid-run; such a change is rejected to prevent silent target drift. Edit config when planning the **next** run. Run status again to reconcile `progress.json` and `state.md` from authoritative rows.

## Worker discovery and isolation
- Run `python project.py show` to discover configured project-level docs. If your Codex conversation has no AGENT_ID, call `python claim_agent.py claim` **once**; never re-claim after context compaction.
- Run `python project.py manifest --agent-id YOUR_ID` to find only your worker's instructions, identity, keywords, cursor, progress, state and logs. Do not hardcode paths or modify another agent's files.
- Run `python record.py status --agent-id YOUR_ID` to see the authoritative `run_id`, `initial_count`, `target_count`, `current_count`, and new images this run. The old agents' images remain globally deduplicated.

## Collection procedure
- First set your real Wikimedia contact in config or `HRVS_WIKIMEDIA_CONTACT`, and run `python commons_api.py test`.
- Use assigned keywords and repeat `python commons_api.py collect --agent-id YOUR_ID --max-pages 8 --max-accept 25`. These flags are **batch sizes**, not quotas. `commons_api.py` pins `run_id` per batch; an old process cannot silently write into a new run.
- Only record qualified original Wikimedia Commons URL **metadata candidates** via `record.py`. Resolution >=8MP. Seek realistic, crowded urban imagery with many small objects and distractors. Metadata alone does NOT visually verify suitability.
- `record.py` serializes every write, URL/filename dedupe and quota check under an OS lock and atomically replaces the workbook. After a crash between the append and progress.json update, reconcile from workbook rows after the run baseline. A duplicate never increments count.
- Only stop due to the **current run's** `agent_new_images_collected >= agent_quota`, or a real blocker recorded using `record.py note`. Do not impose global caps.

## Useful commands
```powershell
python project.py show
python project.py manifest --agent-id YOUR_ID
python record.py status --all                # First-ever init OR resume; NOT a new run
python record.py status --agent-id YOUR_ID
python record.py start-run                   # Explicit NEW run, only if old is complete
python record.py start-run --force           # Explicit NEW run, abandon incomplete old quotas
python record.py audit
python commons_api.py test
python commons_api.py collect --agent-id YOUR_ID --max-pages 8 --max-accept 25
```

## Safety
Use Wikimedia Commons Action API only: no Browser, Computer Use, Playwright, MCP browsing or image downloading. Do not fabricate URLs, silently reinitialize a run, reset state, overwrite the existing XLSX, hand-edit workbook rows, claim visual QA, or bypass API throttling. Inspect licenses/visual quality separately before publication.
