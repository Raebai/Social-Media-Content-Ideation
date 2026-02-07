---
status: complete
phase: 02-query-api-integration
source: [02-01-SUMMARY.md, 02-02-SUMMARY.md]
started: 2026-02-07T17:35:00Z
updated: 2026-02-07T17:35:00Z
---

## Current Test

[testing complete]

## Tests

### 1. Query generation produces 3-6 queries per row
expected: Run `python enrich_calendar.py` and observe that each of the 30 rows shows 3-6 queries. Queries should be short (max 5 words each) and relevant to the row's topic.
result: pass

### 2. Format-specific queries match content type
expected: For each row, one query should clearly match the content type mapping (e.g., Story rows include "founder story", BTS rows include "day in the life", Teach rows include "startup lesson"). Check a few rows across different types.
result: pass

### 3. APIFY_TOKEN missing produces clear error
expected: Run `python -c "from enrich_calendar import fetch_tiktok_results; fetch_tiktok_results(['test'])"` without APIFY_TOKEN set. Should get a clear ValueError saying "APIFY_TOKEN not found in environment" (not a crash or obscure error).
result: pass

### 4. Query output looks useful for TikTok search
expected: Looking at the generated queries, they should feel like something you'd actually type into TikTok search. Not too generic ("content"), not too specific/long. Should capture the essence of each content idea in search-friendly terms.
result: pass

## Summary

total: 4
passed: 4
issues: 0
pending: 0
skipped: 0

## Gaps

[none yet]
