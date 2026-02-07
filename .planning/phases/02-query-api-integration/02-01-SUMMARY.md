---
phase: 02-query-api-integration
plan: 01
subsystem: query-generation
tags: [tiktok, search, apify, api-discovery, query-generation]

requires:
  - phase: 01-input-reading
    provides: ContentIdea dataclass with topic, description, content_type

provides:
  - generate_queries() function returning 3-6 TikTok search queries per content idea
  - TYPE_QUERY_MAP constant mapping content types to format-specific queries
  - Complete Apify TikTok Scraper actor schema documentation

affects:
  - 02-02 (API integration implementation - depends on schema knowledge)
  - 03-XX (Filtering phases - will use query results)

tech-stack:
  added: []
  patterns: [keyword-extraction, query-generation, format-based-search]

key-files:
  created: []
  modified: [enrich_calendar.py]

key-decisions:
  - "Use format-specific queries from TYPE_QUERY_MAP to ensure at least one content-type-aligned query per idea"
  - "Query word limit set to 5 words for TikTok search effectiveness"
  - "Keyword extraction filters stop words and requires >3 character length"
  - "Use searchQueries array input for Apify actor (not hashtags or profiles)"

duration: 2.40min
completed: 2026-02-07
---

# Phase 2 Plan 1: Query Generation and Schema Discovery Summary

**Implemented query generation with format-specific queries and documented complete Apify TikTok Scraper schema for Plan 02-02**

## Performance
- **Duration:** 2.40 min
- **Started:** 2026-02-07T17:12:43Z
- **Completed:** 2026-02-07T17:15:05Z
- **Tasks:** 2
- **Files modified:** 1

## Accomplishments

1. **Query Generation Implementation**
   - Created TYPE_QUERY_MAP with 7 content type format queries
   - Implemented generate_queries() returning 3-6 queries per ContentIdea
   - All queries constrained to <=5 words for TikTok search optimization
   - At least one format-specific query per content type guaranteed
   - Keyword extraction with stop word filtering and length requirements
   - Query deduplication and fallback handling for minimum 3 queries

2. **Apify Schema Discovery**
   - Confirmed actor ID: clockworks/tiktok-scraper (GdWCkxBtKWOsKjdch)
   - Documented complete input schema with all available fields
   - Documented expected output schema with audio/music metadata
   - Documented API call pattern for running actors and retrieving results
   - Identified optimal input fields for Plan 02-02 implementation

## Task Commits

1. **Task 1: Implement generate_queries() with Type-based format queries** - `8c953f0` (feat)

**Plan metadata:** (pending - created with SUMMARY commit)

## Files Created/Modified

- `enrich_calendar.py` - Added TYPE_QUERY_MAP constant and generate_queries() function for TikTok search query generation

## Decisions Made

1. **TYPE_QUERY_MAP Design**: Mapped each of the 7 content types to a format-specific search query that aligns with TikTok content patterns (e.g., "Story" -> "founder story", "BTS" -> "day in the life")

2. **5-Word Query Limit**: Set hard limit of 5 words per query based on TikTok search effectiveness patterns. Queries exceeding this are truncated.

3. **Keyword Extraction Strategy**: Extract keywords >3 characters from topic and description, filtering common stop words. Prioritize topic keywords over description keywords.

4. **Query Diversity**: Generate 3-6 queries per idea with multiple strategies:
   - Format-specific query (from TYPE_QUERY_MAP)
   - Topic as query (if <=5 words)
   - Keyword combination queries (2-3 keywords)
   - Niche query (content type + topic keyword)
   - Cross-pollination (description keyword + format keyword)

5. **Apify Input Format**: Use `searchQueries` array for actor input (not `hashtags` or `profiles`) with `/video` search section for video-focused results.

## Deviations from Plan

None - plan executed exactly as written.

## Issues Encountered

None

## User Setup Required

**Service:** Apify
**Why:** TikTok data fetching via clockworks/tiktok-scraper actor
**Environment Variable:**
- `APIFY_TOKEN` - Get from Apify Console -> Settings -> Integrations -> API tokens

## Apify Schema Discovery

### Actor Information

**Actor ID:** clockworks/tiktok-scraper
**Full ID:** GdWCkxBtKWOsKjdch
**URL Format:** clockworks~tiktok-scraper
**Status:** Active and well-maintained (65M+ total runs, 4.65/5.0 rating)

### Input Schema - Key Fields for Plan 02-02

The actor accepts a JSON input with the following relevant fields:

```json
{
  "searchQueries": ["array", "of", "search", "terms"],
  "resultsPerPage": 10,
  "searchSection": "/video",
  "searchSorting": "0",
  "searchDatePosted": "0"
}
```

**Field Details:**

| Field | Type | Description | Default |
|-------|------|-------------|---------|
| `searchQueries` | array[string] | Search keywords - applies to videos and profiles | required |
| `resultsPerPage` | integer | Number of videos per query (1-1000000) | 1 |
| `searchSection` | string | "" (Top), "/video" (Videos only), "/user" (Profiles) | "" |
| `searchSorting` | string | "0" (Most relevant), "1" (Most liked), "3" (Latest) | "0" |
| `searchDatePosted` | string | "0" (All), "1" (24h), "2" (Week), "3" (Month), "4" (3mo), "5" (6mo) | "0" |

**Other Available Fields** (optional, some charged):
- `hashtags`: Array of hashtags to scrape
- `profiles`: Array of usernames to scrape
- `postURLs`: Array of direct video URLs
- `oldestPostDateUnified`: Date filter for profile videos ($)
- `leastDiggs` / `mostDiggs`: Popularity filters ($)
- `shouldDownloadVideos`: Download videos to storage ($)
- `commentsPerPost`: Scrape comments ($)
- `proxyCountryCode`: Country-specific scraping with residential proxies ($)

### Output Schema - Per Video Result

Each result in the dataset contains:

```json
{
  "id": "string - Video ID",
  "text": "string - Video caption/description",
  "createTime": "integer - Unix timestamp",
  "createTimeISO": "string - ISO 8601 timestamp",

  "authorMeta": {
    "id": "string - Author user ID",
    "name": "string - Author username (@handle)",
    "nickName": "string - Author display name",
    "verified": "boolean - Verified account",
    "signature": "string - Author bio",
    "avatar": "string - Profile picture URL",
    "following": "integer - Accounts following",
    "fans": "integer - Follower count",
    "heart": "integer - Total likes received",
    "video": "integer - Total videos posted"
  },

  "musicMeta": {
    "musicId": "string - Audio track ID",
    "musicName": "string - Audio track name",
    "musicAuthor": "string - Audio creator",
    "musicOriginal": "boolean - Is original audio",
    "musicAlbum": "string - Album name if applicable",
    "playUrl": "string - Audio file URL",
    "coverThumb": "string - Audio cover thumbnail URL",
    "coverMedium": "string - Audio cover medium URL",
    "coverLarge": "string - Audio cover large URL",
    "duration": "integer - Audio duration in seconds"
  },

  "covers": {
    "default": "string - Default cover image URL",
    "origin": "string - Original cover URL",
    "dynamic": "string - Animated cover URL"
  },

  "webVideoUrl": "string - TikTok web URL",
  "videoUrl": "string - Direct video URL",
  "diggCount": "integer - Likes/hearts",
  "shareCount": "integer - Shares",
  "playCount": "integer - Views",
  "commentCount": "integer - Comments",
  "downloaded": "boolean - Video downloaded flag",
  "mentions": "array - @mentioned users",
  "hashtags": "array - Hashtags used",
  "effectStickers": "array - Effects/stickers used"
}
```

**Critical Fields for This Project:**
- `musicMeta.*` - Complete audio metadata for soundtrack recommendations
- `webVideoUrl` - Link to example videos
- `diggCount`, `playCount`, `commentCount` - Performance metrics for filtering
- `authorMeta.name`, `authorMeta.fans` - Creator credibility
- `text` - Video caption for context matching
- `hashtags` - Additional trend signals

### API Call Pattern

**1. Run the Actor:**

```bash
POST https://api.apify.com/v2/acts/clockworks~tiktok-scraper/runs
Authorization: Bearer {APIFY_TOKEN}
Content-Type: application/json

{
  "searchQueries": ["query1", "query2"],
  "resultsPerPage": 10,
  "searchSection": "/video"
}
```

**2. Monitor Run Status:**

Response contains:
```json
{
  "data": {
    "id": "run_id",
    "status": "RUNNING" | "SUCCEEDED" | "FAILED",
    "defaultDatasetId": "dataset_id"
  }
}
```

Poll status:
```bash
GET https://api.apify.com/v2/acts/clockworks~tiktok-scraper/runs/{run_id}
```

**3. Retrieve Results:**

When status is "SUCCEEDED":
```bash
GET https://api.apify.com/v2/datasets/{dataset_id}/items
```

Returns array of video results matching output schema.

### Python Implementation Options

**Option A: requests library (lightweight)**
```python
import requests
import os

APIFY_TOKEN = os.getenv("APIFY_TOKEN")
headers = {"Authorization": f"Bearer {APIFY_TOKEN}"}

# Start run
response = requests.post(
    "https://api.apify.com/v2/acts/clockworks~tiktok-scraper/runs",
    json={"searchQueries": ["founder story"], "resultsPerPage": 10},
    headers=headers
)
run_data = response.json()["data"]
run_id = run_data["id"]
dataset_id = run_data["defaultDatasetId"]

# Wait for completion (polling)
# ... polling logic ...

# Get results
results = requests.get(
    f"https://api.apify.com/v2/datasets/{dataset_id}/items",
    headers=headers
).json()
```

**Option B: apify-client library (official, recommended)**
```python
from apify_client import ApifyClient

client = ApifyClient(os.getenv("APIFY_TOKEN"))

# Run actor and wait for completion
run = client.actor("clockworks/tiktok-scraper").call(
    run_input={
        "searchQueries": ["founder story"],
        "resultsPerPage": 10,
        "searchSection": "/video"
    }
)

# Get results
results = client.dataset(run["defaultDatasetId"]).list_items().items
```

**Recommendation for Plan 02-02:** Use `apify-client` library for cleaner API and automatic polling/error handling.

### Pricing Notes

- Actor uses "Pay Per Event" pricing model
- Base charge: $0.005 per actor start
- Per result: $0.003 per video scraped
- Additional filters (date, popularity): $0.001 per item
- For 100 results per query across 50 ideas (5 queries each): ~$0.005 + (100 * 5 * 50 * $0.003) = ~$75 per full run
- Free tier: 5760 trial minutes available for testing

### Key Insights for Plan 02-02

1. **Use searchQueries input**: Best fit for our generated queries from generate_queries()
2. **searchSection: "/video"**: Focus on video results, not profiles
3. **resultsPerPage: 10-20**: Good balance for finding relevant examples without excessive API costs
4. **musicMeta is comprehensive**: All audio fields needed for Phase 4 are present
5. **No authentication issues expected**: Actor is public and well-maintained
6. **Error handling needed**: Actor can fail or time out - need retry logic
7. **Rate limiting**: Consider batching queries to avoid overwhelming API

## Next Phase Readiness

**Phase 2 Plan 2 is ready to proceed** with the following inputs:
- generate_queries() function available in enrich_calendar.py
- Complete Apify input/output schema documented
- API call pattern and authentication requirements known
- Python implementation options evaluated (recommend apify-client)

**No blockers identified.**

**Recommended next steps for Plan 02-02:**
1. Install apify-client: `pip install apify-client`
2. Implement fetch_tiktok_examples() using documented schema
3. Add rate limiting and error handling for API stability
4. Store results with proper attribution (video URLs, music metadata)
