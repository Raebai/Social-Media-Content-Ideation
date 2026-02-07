---
phase: 02-query-api-integration
plan: 02
subsystem: api-integration
tags: [apify, tiktok, api, retry-logic, caching, requests]

requires:
  - "02-01: Apify actor schema discovery"
  - "01-01: ContentIdea data structure and query generation"

provides:
  - "fetch_tiktok_results(): API integration with retry and caching"
  - "Per-query result caching to avoid duplicate API calls"
  - "150 result cap enforcement per row"
  - "End-to-end Phase 2 pipeline structure"

affects:
  - "03-*: Phase 3 will receive List[Dict] of TikTok video results"
  - "Data processing pipeline will work with Apify's output schema"

tech-stack:
  added:
    - "requests>=2.28.0: HTTP client for Apify REST API"
  patterns:
    - "Exponential backoff retry (1s, 2s, 4s delays)"
    - "In-memory dictionary caching for query deduplication"
    - "Synchronous API endpoint for simplicity"

key-files:
  created: []
  modified:
    - path: "enrich_calendar.py"
      lines-added: 203
      purpose: "Added Apify API integration functions and pipeline helper"
    - path: "requirements.txt"
      lines-added: 1
      purpose: "Added requests dependency"

key-decisions:
  - slug: "synchronous-apify-endpoint"
    title: "Use synchronous Apify endpoint instead of async polling"
    rationale: "Simpler implementation, built-in timeout handling (300s), adequate for batch processing"
    impact: "Cleaner code, easier debugging, sufficient for content calendar use case"

  - slug: "simple-dict-cache"
    title: "Use module-level dict for query caching instead of Redis/external cache"
    rationale: "Single-run execution model, no need for persistence across runs"
    impact: "Zero infrastructure dependency, instant cache hits within session"

  - slug: "token-env-var"
    title: "Require APIFY_TOKEN environment variable"
    rationale: "Standard security practice, keeps credentials out of code"
    impact: "User must set env var before running, clear error message if missing"

duration: "2.18 min"
completed: 2026-02-07
---

# Phase 02 Plan 02: Apify Integration Summary

**One-liner:** Implemented fetch_tiktok_results() with exponential backoff retry, per-query caching, and 150 result cap using Apify's synchronous endpoint.

---

## Performance

**Execution Time:** 2.18 minutes (131 seconds)
**Tasks Completed:** 2/2 (100%)
**Commits:** 2 atomic commits

**Breakdown:**
- Task 1 (fetch_tiktok_results): ~1.5 min
- Task 2 (pipeline verification): ~0.68 min

---

## Accomplishments

### API Integration (Task 1)

Implemented complete Apify TikTok scraper integration using schema discovered in Plan 02-01:

**Core Functions:**
- `fetch_tiktok_results(queries, token)`: Main entry point with caching and cap logic
- `_run_apify_actor(query, token, max_items)`: Handles single API call to Apify
- `_call_with_retry(fn, *args, **kwargs)`: Exponential backoff retry wrapper

**Key Features:**
1. **Retry Logic:** Up to 3 attempts with exponential backoff (1s, 2s, 4s)
2. **Caching:** Module-level dict prevents duplicate queries within a run
3. **Result Cap:** Enforces MAX_RESULTS_PER_ROW (150) regardless of query count
4. **Error Handling:** Clear error messages, graceful degradation (empty list on final failure)

**API Call Pattern:**
- Endpoint: `POST https://api.apify.com/v2/acts/clockworks~tiktok-scraper/run-sync-get-dataset-items`
- Authentication: Token as query parameter
- Input: `searchQueries`, `resultsPerPage`, `searchSection`, `searchSorting`, `searchDatePosted`
- Timeout: 300 seconds (5 minutes)
- Response: JSON array of video result dictionaries

**Constants Added:**
```python
APIFY_BASE_URL = "https://api.apify.com/v2"
APIFY_ACTOR_ID = "clockworks~tiktok-scraper"
MAX_RESULTS_PER_ROW = 150
RESULTS_PER_QUERY = 25
MAX_RETRIES = 3
```

### Pipeline Structure (Task 2)

Added `enrich_row_queries(idea)` helper function that:
- Bridges ContentIdea → query generation → API fetching
- Returns structured dict: `{row_number, queries, query_count}`
- Ready for Phase 3 extension (will add `tiktok_results` field)

Updated main block to demonstrate full Phase 2 pipeline:
- Loads all 30 content ideas
- Generates queries for each
- Displays query count and details
- Shows total queries (159 across 30 rows, avg 5.3 per row)

**Verification Results:**
- All 30 rows generate 3-6 queries ✓
- All queries ≤5 words ✓
- All rows include format-specific query from TYPE_QUERY_MAP ✓

---

## Task Commits

| Task | Commit | Description | Files |
|------|--------|-------------|-------|
| 1 | b165fbd | Implement fetch_tiktok_results with retry and caching | enrich_calendar.py, requirements.txt |
| 2 | 36bbe30 | Add end-to-end query-to-fetch pipeline structure | enrich_calendar.py |

---

## Files Created/Modified

### Modified: enrich_calendar.py (+203 lines)

**Added Imports:**
- `requests`: HTTP client for Apify API
- `os`, `sys`, `time`, `json`: Supporting utilities

**Added Constants:**
- Apify configuration (URL, actor ID)
- Result limits (150 per row, 25 per query)
- Retry settings (3 attempts max)
- Query cache (module-level dict)

**Added Functions:**
1. `_call_with_retry()`: Exponential backoff retry wrapper
2. `_run_apify_actor()`: Single Apify API call handler
3. `fetch_tiktok_results()`: Main API integration function
4. `enrich_row_queries()`: Phase 2 pipeline helper

**Updated Main Block:**
- Demonstrates query generation for all rows
- Shows query count statistics
- Notes APIFY_TOKEN requirement

### Modified: requirements.txt (+1 line)

Added `requests>=2.28.0` dependency for HTTP API calls.

---

## Decisions Made

### 1. Synchronous Endpoint Over Async Polling

**Decision:** Use `run-sync-get-dataset-items` endpoint instead of async `runs` + polling pattern.

**Rationale:**
- Simpler code (no polling loop, status checking)
- Built-in 300s timeout handling
- Adequate for batch processing (not real-time)
- Easier debugging and error handling

**Alternatives Considered:**
- Async with polling: More complex, overkill for batch use case
- apify-client library: Extra dependency, requests is sufficient

**Impact:** Cleaner implementation, faster development, sufficient performance for content calendar enrichment.

---

### 2. Simple Dict Cache Over External Cache

**Decision:** Use module-level Python dict for query caching instead of Redis or other external cache.

**Rationale:**
- Single-run execution model (no need for cross-run persistence)
- Immediate cache hits within session
- Zero infrastructure/dependency overhead
- Adequate for 30 rows with ~5 queries each

**Alternatives Considered:**
- Redis: Overkill, adds deployment complexity
- File-based cache: Slower, unnecessary for in-memory use case
- No cache: Wastes API calls on duplicate queries

**Impact:** Zero infrastructure requirement, instant performance, perfect fit for use case.

---

### 3. Environment Variable for APIFY_TOKEN

**Decision:** Require APIFY_TOKEN as environment variable, raise ValueError if missing.

**Rationale:**
- Standard security practice (keeps secrets out of code)
- Works in dev and production environments
- Clear error message guides user setup
- Optional parameter allows injection for testing

**Alternatives Considered:**
- Config file: More complexity, security risk if committed
- Hardcoded: Security violation
- Command-line arg: Less standard for API tokens

**Impact:** Secure credential handling, clear error on missing token, standard practice.

---

## Deviations from Plan

None - plan executed exactly as written.

All specified features implemented:
- fetch_tiktok_results() with retry, caching, cap ✓
- Exponential backoff (1s, 2s, 4s) ✓
- Per-query caching ✓
- 150 result cap per row ✓
- requirements.txt updated ✓
- Pipeline verification across all 30 rows ✓

---

## Issues Encountered

### None

Execution was smooth. The Apify schema from Plan 02-01 was accurate and complete, allowing straightforward implementation.

**Verification notes:**
- All imports work correctly
- Phase 1 code (load_content_ideas) still functions perfectly
- APIFY_TOKEN validation works as expected
- Constants are accessible and correct
- Query generation verified across all 30 rows

---

## User Setup Required

### APIFY_TOKEN Environment Variable

**What:** Apify API token for TikTok scraper access

**Why:** Authenticates API requests to clockworks/tiktok-scraper actor

**How to set:**

1. Get token from Apify Console:
   - Visit: https://console.apify.com/account/integrations
   - Go to Settings → Integrations → API tokens
   - Copy your API token

2. Set environment variable:
   ```bash
   # Windows (Command Prompt)
   set APIFY_TOKEN=apify_api_xxxxxxxxxxxxxx

   # Windows (PowerShell)
   $env:APIFY_TOKEN="apify_api_xxxxxxxxxxxxxx"

   # Linux/Mac
   export APIFY_TOKEN=apify_api_xxxxxxxxxxxxxx
   ```

3. Verify:
   ```python
   from enrich_calendar import fetch_tiktok_results
   results = fetch_tiktok_results(["founder story"])
   print(f"Fetched {len(results)} results")
   ```

**Error if missing:** `ValueError: APIFY_TOKEN not found in environment`

---

## Next Phase Readiness

### Phase 3: Data Processing & Analysis

**Ready:** Yes

**What Phase 3 Receives:**

Phase 2 provides `fetch_tiktok_results()` that returns:
```python
List[Dict[str, Any]]  # List of TikTok video result dictionaries
```

**Result Dictionary Schema (from Apify):**
```json
{
  "id": "string - Video ID",
  "text": "string - Video caption/description",
  "createTime": "integer - Unix timestamp",
  "createTimeISO": "string - ISO 8601 timestamp",
  "authorMeta": {
    "id": "string",
    "name": "string - @handle",
    "nickName": "string - Display name",
    "verified": "boolean",
    "fans": "integer - Follower count"
  },
  "musicMeta": {
    "musicId": "string",
    "musicName": "string",
    "musicAuthor": "string",
    "musicOriginal": "boolean",
    "playUrl": "string",
    "duration": "integer"
  },
  "webVideoUrl": "string - TikTok web URL",
  "videoUrl": "string - Direct video URL",
  "diggCount": "integer - Likes",
  "shareCount": "integer - Shares",
  "playCount": "integer - Views",
  "commentCount": "integer - Comments"
}
```

**Usage Pattern for Phase 3:**
```python
from enrich_calendar import load_content_ideas, generate_queries, fetch_tiktok_results

ideas, _ = load_content_ideas("Content ideas.xlsx")

for idea in ideas:
    queries = generate_queries(idea)
    results = fetch_tiktok_results(queries)

    # Phase 3 will:
    # 1. Extract audio tracks from results (musicMeta)
    # 2. Analyze engagement patterns (diggCount, playCount, etc.)
    # 3. Identify top performers
    # 4. Match to content ideas
```

**Key Points:**
- Results are capped at 150 per row (MAX_RESULTS_PER_ROW)
- Caching prevents duplicate API calls within a run
- Retry logic handles transient failures
- Empty list returned on final failure (graceful degradation)

**Blockers:** None

**Recommendations for Phase 3:**
1. Focus on `musicMeta` for audio track extraction
2. Use engagement metrics (`diggCount`, `playCount`) for ranking
3. Consider `createTimeISO` for recency weighting
4. Use `authorMeta.verified` and `authorMeta.fans` for authority signals

---

## Metrics

**Code Changes:**
- Lines added: 204
- Lines modified: 2
- Functions added: 4
- Dependencies added: 1

**Test Coverage:**
- Manual verification: 5/5 checks passed
- Pipeline verification: 30/30 rows validated
- Query validation: 159 queries, all ≤5 words

**API Integration:**
- Endpoint: Apify run-sync-get-dataset-items
- Timeout: 300s
- Retry attempts: 3
- Backoff: Exponential (1s, 2s, 4s)
- Cache: In-memory dict (query → results)
- Result cap: 150 per row

---

**Phase 2 Status:** Complete (2/2 plans done)
**Next:** Phase 3 - Data Processing & Analysis
