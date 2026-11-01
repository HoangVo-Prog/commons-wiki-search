# HRVS Wikimedia Commons — independent quota / API-only Codex instructions

## Goal and configuration
Collect unique original Wikimedia Commons high-resolution image **URL metadata candidates** to the shared workbook. No image bytes are downloaded. **There is no total/global image target or cap.** Each agent has its **own** `agent_quota` entry in `config.json`, keyed by its two-digit agent ID. The ONLY stopping rule is:

`agent_new_images_collected >= agent_quota`

Neither existing workbook size nor other agents' progress changes an agent's quota. Do not calculate an agent quota from the number of agents, workbook rows, or another global count. Only `config.json` defines quotas and locations.

## Startup and worker isolation
1. Before starting collectors, the operator must put the **existing** `images.xlsx` in its configured location and run `python record.py status --all` **once**. This initializes `collection_state.json`, a durable snapshot of all pre-existing workbook rows. If the file already exists, **never delete/reset it** to change counts. Preserve and migrate it together with the workbook on every restart, backup, or move. Never overwrite a populated workbook with the empty template.
2. Run `python project.py show`, then read the root instructions and task instructions returned by its manifest. If your Codex conversation does not already have an assigned ID, run `python claim_agent.py claim` exactly once. After assigning ID, run `python project.py manifest --agent-id YOUR_ID` to locate **only your worker's** instructions, identity, keywords, cursor, progress, state, and recent events log. Do not read or modify other workers' directories.
3. Once assigned, keep the same AGENT_ID across context compactions and restarts. Do not claim a second ID. The claim registry uses the shared lock to avoid collisions.
4. Run `python record.py status --agent-id YOUR_ID`. The reported `agent_new_images_collected` counts exclusively new rows *after* the saved historical baseline, not historical rows even if those carry your ID. The workbook and `collection_state.json` are authoritative; `progress.json` is a persistently updated recovery cache.

## Search and collection
- Use the original-image `imageinfo.url` from the Wikimedia Commons Action API through `commons_api.py`. Collect only metadata candidates for realistically crowded urban scenes with small objects/distractors. Prefer photographs of markets, busy streets, pedestrian zones, festivals, shopping districts, public transit, etc. Avoid aerial imagery, staged portraits, illustrations, panoramas, and empty streets. Minimum resolution is 8 megapixels.
- Image suitability is **not visually verified**. The record explanation must explicitly say metadata-only and unverified; the final dataset requires separate visual/license review.
- Configure an operator-supplied genuine `wikimedia_contact` or `HRVS_WIKIMEDIA_CONTACT` environment variable; never invent one. Run `python commons_api.py test` before collecting.
- Use your own extensive keyword pool and persist the assigned cursor. Repeatedly execute `python commons_api.py collect --agent-id YOUR_ID --max-pages 8 --max-accept 25`; these flags control **batch processing**, not your quota. Add more queries ONLY to your own keyword file if needed.
- Stop **only** when `agent_new_images_collected >= agent_quota`, or when a real technical blocker is logged. Other agents completing, an arbitrary workbook size, or a hypothetical global total is never a stopping reason.
- Use `record.py add` exclusively for writes. It serializes quota checks, global dedup against **every historical and new URL**, Excel append, and progress updates under one shared file lock. The XLSX append uses atomic replacement. If a process crashes after appending but before progress.json updates, status reconstructs the count from the workbook, so it cannot double-count on resume.

## Monitoring and recovery
```powershell
python project.py show
python project.py manifest --agent-id YOUR_ID
python record.py status --all
python record.py status --agent-id YOUR_ID
python record.py audit
python commons_api.py test
python commons_api.py collect --agent-id YOUR_ID --max-pages 8 --max-accept 25
```
`YOUR_ID` means the claimed or reserved ID, not a literal argument. Before resuming, preserve the paired `images.xlsx` + `collection_state.json`, then run status. Do not edit either outside the collector; never remove the baseline to reset counts. When quotas change in `config.json`, status immediately uses the new per-agent quotas without erasing previous accepted rows.

## API and data safety
**No Browser, Computer Use, Playwright, MCP browsing, scraping or image downloads.** Do not create side workbooks, synthesize original URLs, alter other agent files, or bypass shared API throttling/backoff. Do not use `--relaxed-title` without operator approval. If blocked, log a specific reason with `record.py note`, report actual newly collected counts, and stop safely.
