# HRVS Collection — per-agent quota / metadata-only Wikimedia API

Collect high-resolution Wikimedia Commons image URL **metadata candidates** (no downloaded images) in a shared `images.xlsx`. Ten configured workers have independent quotas and keyword pools; workbook size never caps collection. Neither a global image target nor a derived division of a total is used.

## Configure independent quotas

Edit `config.json` only. `agent_quota` maps each configured two-digit agent ID to its **new image quota**:

```json
{
  "agent_count": 10,
  "agent_quota": {
    "01": 500,
    "02": 500,
    "03": 500,
    "04": 500,
    "05": 500,
    "06": 500,
    "07": 500,
    "08": 500,
    "09": 500,
    "10": 500
  },
  "wikimedia_contact": "YOUR_EMAIL_OR_WIKIMEDIA_USER_PAGE"
}
```

These are editable **configuration examples**, not constants in any script. Each worker stops when **`agent_new_images_collected >= agent_quota`**, independent of other workers and workbook size. Each key must exist for its agent; all quotas must be nonnegative integers. The included project initializes ten worker directories; changing agent count requires corresponding worker directories/IDs.

## Critical first-use migration for a populated workbook

1. **Keep** your existing `images.xlsx` in the path configured as `paths.workbook`; do **not** replace it with an empty template. It can already contain rows (including rows with current agent IDs), or it can be empty.
2. **Before running any collection process**, run `python record.py status --all` once to create `collection_state.json` beside the config at `paths.collection_state`. The baseline records the number and hash of all *existing* data rows; none of these rows count toward any agent's new quota. They remain included in global URL/filename deduplication.
3. **Never delete, regenerate or reset `collection_state.json` after collection begins**. Always backup, move, and restore it **together with its matching `images.xlsx`**. Removing the state would incorrectly reclassify new rows as pre-existing. This file is initialized automatically under the shared lock if missing; proactively running status before agents is best practice.

Each accepted append happens under the shared OS lock, rechecks the agent's newly-collected count and global URL deduplication, atomically replaces the XLSX, and updates the per-agent `progress.json` and `state.md`. After crashes, actual counts are rebuilt from workbook rows *after* the saved baseline, so stale cached progress cannot cause double-counting. No records are removed or rewritten. Out-of-band modification/reordering of original rows is detected through the baseline hash and requires recovery, not resetting the baseline.

## Run independent Codex agents

Paste `PROMPT_FOR_CODEX.txt` into each independent Codex conversation at project root, or use PowerShell launchers:

```powershell
python project.py show
python record.py status --all
python claim_agent.py status
python commons_api.py test
.\run_one_agent.ps1 -AgentID 1
.\run_all_agents.ps1
# Resume only when earlier session for this ID is stopped:
.\run_one_agent.ps1 -AgentID 1 -Resume
```

`claim_agent.py` atomically assigns unused IDs. Do not claim again in the same conversation. The worker manifest (`python project.py manifest --agent-id 01`) resolves all instruction and worker paths dynamically and displays that agent's `agent_quota`.

```powershell
python commons_api.py collect --agent-id 01 --max-pages 8 --max-accept 25
python record.py status --agent-id 01
python record.py status --all
python record.py audit
```

`--max-accept` and `--max-pages` are **batch limits only**. API calls are globally throttled. Add keywords only to your own worker file. Configure a genuine Wikimedia contact in `config.json` or `HRVS_WIKIMEDIA_CONTACT` before making API calls. Do not invent a contact.

## Project files

- `config.json`: `agent_count`, independently set `agent_quota` per ID, API settings, configured paths.
- `project.py`: validates config and provides quota/path manifests.
- `collection_state.json` (created on first status/add, not initially bundled): persistent historical workbook baseline; preserve forever with workbook.
- `images.xlsx`: existing historical + newly appended candidate records, retaining the original four-column format.
- `record.py`: locked atomic append, deduplication against all workbook records, baseline reconciliation, per-agent progress and audit.
- `commons_api.py`: Wikimedia API search and per-agent collection loop.
- `claim_agent.py`: atomic ID claim/reservation.
- `workers/`: independent worker keywords, cursor, progress, state and logs.
- `AGENTS.md`, `TASK_PROMPT.txt`, `PROMPT_FOR_CODEX.txt`: instructions and prompts.

## Dataset limitations

Search results are metadata candidates, not visually verified HRVS images. Metadata cannot establish genuine visual complexity, small-object scale or authenticity. Carry out visual QA and per-file Wikimedia Commons license review before distributing image datasets. The shared Excel file has its intrinsic Excel format limitations; **no additional project-wide image quota or cap is imposed**.
