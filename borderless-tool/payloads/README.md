# payloads

One file per week, named `YYYY-MM-weekN.json`, holding exactly what was sent
to `borderless_update.py` that week.

`input.json` in the parent folder is the scratch working copy and is
gitignored; these are the kept record, so a run can be reproduced or audited
later without retyping the text.

Run one with:

```bash
python3 ../borderless_update.py payloads/2026-09-week4.json --dry-run
```
