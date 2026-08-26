# PR vào NousResearch/hermes-plugin-index

> Copy nội dung dưới đây khi bạn mở PR vào repo `NousResearch/hermes-plugin-index`
> (thêm 1 entry vào file `index.json`, mảng `"plugins"`).

---

## PR Title (tiêu đề)

```
Add excel_line — Excel-backed long-term memory provider
```

## PR Body (mô tả)

```markdown
## Summary

Add `excel_line` to the community plugin index. It is a Hermes **memory
provider** plugin that keeps durable, human-auditable knowledge in Excel
workbooks alongside the built-in `MEMORY.md` / `USER.md`.

- Mirrors the built-in memory write path (`on_memory_write`) so agent-authored
  and auto-extracted memories persist.
- `prefetch()` injects indexed memories into every session (auto-retrieve, like
  built-in memory).
- Exposes `search` / `read` tools for on-demand retrieval.
- **Direct-store mode**: when the agent passes explicit `zone` + `brief` /
  `content`, the record is written straight to the Excel store with **no LLM
  dependency** — this fixes the original bug where memory was logged but never
  retrievable ("saved but unusable") because the classifier silently dropped
  logs on LLM failure.
- `on_session_end` auto-extract falls back to a raw-text backup, so memory is
  never silently lost.

## Entry added to index.json

{
  "name": "excel_line",
  "description": "Excel-backed long-term memory provider for Hermes. Keeps durable, human-auditable knowledge in Excel workbooks alongside built-in MEMORY.md/USER.md; mirrors the built-in write path, injects indexed memories via prefetch() every session, and exposes search/read tools. Direct-store mode writes straight to the store with no LLM dependency; auto-extract falls back to a raw-text backup so memory is never silently lost.",
  "author": "kalice-vi",
  "tags": ["memory", "excel", "knowledge", "long-term-memory", "provider"],
  "repo": "kalice-vi/hermes-excel-line",
  "ref": "f68ef7fd85af344eb554c94f6998317a5895bd69",
  "homepage": "https://github.com/kalice-vi/hermes-excel-line",
  "capabilities": ["memory"],
  "api_version": 1,
  "added_at": "2026-08-26"
}

## Verification

- `hermes plugins doctor .` → OK: runtime discovery, manifest parsing, import,
  and registration passed.
- 33-test QA suite (`tests/test_excel_line.py`) all green: store CRUD, Unicode /
  Vietnamese search, search-by-title, concurrent writes, cross-process lock with
  stale-lock pid recovery, worker fallback, lazy index, on_session_end data
  survival, exception safety.
- No user data shipped (`.gitignore` excludes `*.xlsx`, `logs/`).

## Install

hermes plugins install kalice-vi/hermes-excel-line
```

---

## Các bước bạn làm (thủ công trên GitHub)

1. Mở https://github.com/NousResearch/hermes-plugin-index (fork nếu cần).
2. Edit file `index.json`.
3. Thêm entry JSON ở trên vào mảng `"plugins"` (sau dấu phẩy của entry cuối).
4. Commit → mở Pull Request với Title + Body ở trên.
5. Đợi maintainer merge → `hermes plugins search excel_line` sẽ thấy.

> Lưu ý: index là "discovery metadata ONLY. Indexed ≠ audited" — PR thêm entry
> thường được merge nhanh nếu metadata đúng định dạng.
