# Requirements: Logara Content Calendar Enricher

**Defined:** 2026-02-06
**Core Value:** One command transforms a bare content calendar into an actionable production sheet with audio picks, proven examples, and remix ideas.

## v1 Requirements

### Input/Output

- [ ] **IO-01**: Read `Content ideas.xlsx` Sheet1 with columns Day, Date, Type, Topic, Description
- [ ] **IO-02**: Construct `idea_text = f"{Type} | {Topic} | {Description}"` per row
- [ ] **IO-03**: Write enriched Excel with all output columns to specified path
- [ ] **IO-04**: Normalize dates in output Excel to consistent format
- [ ] **IO-05**: CLI with `--input` and `--output` flags

### Query Generation

- [ ] **QRY-01**: Generate 3-6 short queries (<=5 words) per row from idea_text
- [ ] **QRY-02**: Include at least one format query based on Type mapping (Story->"founder story", BTS->"day in the life", Teach->"startup lesson", Trend->"new chapter", Breakdown->"mindset shift", Reflection->"founder reflection", Depth->"longform reflection")
- [ ] **QRY-03**: Write topic_keywords and search_queries columns to output

### Apify Integration

- [ ] **API-01**: Use Apify actor `clockworks/tiktok-scraper` with APIFY_TOKEN from env
- [ ] **API-02**: Fetch 20-30 results per query
- [ ] **API-03**: Hard cap total results per row at 150
- [ ] **API-04**: Inspect actor schema and adapt input format as needed
- [ ] **API-05**: Retry logic: 3 attempts with exponential backoff
- [ ] **API-06**: Cache query results within a run to avoid duplicate calls

### Data Processing

- [ ] **DAT-01**: Normalize items to schema: video_id, url, caption, author_username, create_time, views, likes, comments, shares, audio (audio_id, title, author, url)
- [ ] **DAT-02**: Deduplicate results by video_id
- [ ] **DAT-03**: Filter videos older than 120 days when create_time exists
- [ ] **DAT-04**: Score: base = log10(views+1)*0.65 + log10(likes+1)*0.25 + log10(comments+1)*0.10
- [ ] **DAT-05**: Recency boost: +0.2 if <=14 days, +0.1 if <=30 days
- [ ] **DAT-06**: Relevance boost: +0.05 per keyword overlap in caption (cap +0.2)
- [ ] **DAT-07**: Final score = base + recency_boost + relevance_boost

### Example Selection

- [ ] **EX-01**: Select top 3 videos by final score as examples
- [ ] **EX-02**: Write per example: url, views, likes, comments, shares, author, date, caption columns (ex1_, ex2_, ex3_ prefixed)

### Audio Selection

- [ ] **AUD-01**: Pick most common audio among top 20 results by final score
- [ ] **AUD-02**: Fallback to ex1 audio if no repeats found
- [ ] **AUD-03**: Audio confidence: high (>=3 occurrences), medium (2), low (1/unknown)
- [ ] **AUD-04**: Write audio_title, audio_author, audio_id, audio_url, audio_confidence columns

### LLM Text Generation

- [ ] **LLM-01**: Use OpenAI GPT-4o-mini via OPENAI_API_KEY from env
- [ ] **LLM-02**: Generate audio_fit_reason (1 sentence) per row
- [ ] **LLM-03**: Generate hook_summary (1 sentence) per example (ex1_hook_summary, ex2_hook_summary, ex3_hook_summary)
- [ ] **LLM-04**: Generate remix_ideas (2-3 bullets) per row in hopecore + narrative stakes style, matching row Type
- [ ] **LLM-05**: Write ex1_audio_title, ex2_audio_title, ex3_audio_title columns

### Status & Logging

- [ ] **LOG-01**: Enrich status per row: ok (audio + >=2 examples), partial (examples but missing audio or only 1 example), skipped (no results), error (exception)
- [ ] **LOG-02**: Write enrich_status and enrich_reason columns
- [ ] **LOG-03**: Produce run_log.json with per-row: queries, total_results, chosen_audio, example_urls

### Packaging

- [ ] **PKG-01**: Single file `enrich_calendar.py`
- [ ] **PKG-02**: `requirements.txt` with pandas, openpyxl, requests, python-dateutil, openai

## v2 Requirements

### Enhancements

- **V2-01**: Support multiple sheets or date ranges
- **V2-02**: HTML report with embedded video previews
- **V2-03**: Configurable scoring weights via CLI flags
- **V2-04**: Progress bar with ETA during enrichment

## Out of Scope

| Feature | Reason |
|---------|--------|
| Direct TikTok API/scraping | Apify only -- avoids rate limits and ToS issues |
| Web dashboard / UI | CLI tool only |
| Auto-posting / scheduling | Research/planning tool only |
| Real-time monitoring | One-shot enrichment per run |
| Instagram Reels data | TikTok only for v1 |

## Traceability

| Requirement | Phase | Status |
|-------------|-------|--------|
| IO-01 | Phase 1 | Complete |
| IO-02 | Phase 1 | Complete |
| IO-03 | Phase 5 | Pending |
| IO-04 | Phase 5 | Pending |
| IO-05 | Phase 5 | Pending |
| QRY-01 | Phase 2 | Complete |
| QRY-02 | Phase 2 | Complete |
| QRY-03 | Phase 5 | Pending |
| API-01 | Phase 2 | Complete |
| API-02 | Phase 2 | Complete |
| API-03 | Phase 2 | Complete |
| API-04 | Phase 2 | Complete |
| API-05 | Phase 2 | Complete |
| API-06 | Phase 2 | Complete |
| DAT-01 | Phase 3 | Pending |
| DAT-02 | Phase 3 | Pending |
| DAT-03 | Phase 3 | Pending |
| DAT-04 | Phase 3 | Pending |
| DAT-05 | Phase 3 | Pending |
| DAT-06 | Phase 3 | Pending |
| DAT-07 | Phase 3 | Pending |
| EX-01 | Phase 3 | Pending |
| EX-02 | Phase 3 | Pending |
| AUD-01 | Phase 3 | Pending |
| AUD-02 | Phase 3 | Pending |
| AUD-03 | Phase 3 | Pending |
| AUD-04 | Phase 3 | Pending |
| LLM-01 | Phase 4 | Pending |
| LLM-02 | Phase 4 | Pending |
| LLM-03 | Phase 4 | Pending |
| LLM-04 | Phase 4 | Pending |
| LLM-05 | Phase 4 | Pending |
| LOG-01 | Phase 5 | Pending |
| LOG-02 | Phase 5 | Pending |
| LOG-03 | Phase 5 | Pending |
| PKG-01 | Phase 5 | Pending |
| PKG-02 | Phase 5 | Pending |

**Coverage:**
- v1 requirements: 35 total
- Mapped to phases: 35
- Unmapped: 0

---
*Requirements defined: 2026-02-06*
*Last updated: 2026-02-07 after Phase 2 completion*
