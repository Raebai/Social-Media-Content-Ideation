# Logara Content Calendar Enricher

## What This Is

A Python CLI tool that reads a content calendar Excel file (`Content ideas.xlsx`) and enriches each row with trending TikTok audio recommendations, top-performing example videos in the niche, and creative remix ideas. It outputs `Content ideas - enriched.xlsx` using Apify's TikTok scraper for data and OpenAI for creative text generation.

## Core Value

One command transforms a bare content calendar into an actionable production sheet with audio picks, proven examples to reference, and remix ideas ready to film.

## Requirements

### Validated

(None yet -- ship to validate)

### Active

- [ ] Read `Content ideas.xlsx` (Sheet1: Day, Date, Type, Topic, Description)
- [ ] Generate 3-6 short search queries per row from `idea_text = f"{Type}. {Topic}. {Description}"`, including format-specific queries mapped to Type
- [ ] Fetch TikTok results via Apify actor `clockworks/tiktok-scraper`, 20-30 results per query, capped at 150 per row
- [ ] Normalize results into standard schema (video_id, url, caption, author, create_time, views, likes, comments, shares, audio metadata)
- [ ] Deduplicate and filter videos older than 120 days
- [ ] Score videos: `base_score = log10(views+1)*0.65 + log10(likes+1)*0.25 + log10(comments+1)*0.10` + recency_boost + relevance_boost
- [ ] Select top 3 examples per row with full metadata (url, views, likes, comments, shares, author, date, caption)
- [ ] Select suggested audio: most common audio among top 20 results, fallback to ex1 audio
- [ ] Audio confidence: high (>=3 occurrences), medium (2), low (1/unknown)
- [ ] Generate LLM text via OpenAI: audio_fit_reason, hook_summary per example, remix_ideas (2-3 bullets in hopecore + narrative stakes style)
- [ ] Write enriched Excel with all output columns (topic_keywords, search_queries, audio fields, ex1-ex3 fields, remix_ideas, enrich_status, enrich_reason)
- [ ] Normalize dates in output Excel
- [ ] Enrich status: ok / partial / skipped / error with reason
- [ ] Retry logic: 3 attempts with exponential backoff for Apify calls
- [ ] Query result caching within a run (avoid duplicate API calls)
- [ ] Produce `run_log.json` with per-row diagnostics
- [ ] CLI: `python enrich_calendar.py --input "./Content ideas.xlsx" --output "./Content ideas - enriched.xlsx"`

### Out of Scope

- Direct TikTok scraping (Apify only) -- avoids rate limits and ToS issues
- Web dashboard or UI -- CLI only
- Scheduling or auto-posting -- this is research/planning only
- Real-time monitoring -- one-shot enrichment per run

## Context

- Logara is a startup; content calendar covers 30 days of founder journey content
- Content types: Story, BTS, Teach, Trend, Breakdown, Reflection, Depth
- Brand tone: hopecore, hopeful, narrative-driven, cinematic, authentic vulnerability
- The Excel has 30 rows covering Feb 6 - Mar 7, 2026
- Apify token available as `APIFY_TOKEN` env var
- OpenAI key available as `OPENAI_API_KEY` env var

## Constraints

- **Data source**: Apify only (`clockworks/tiktok-scraper`) -- no direct TikTok API
- **Cost control**: Hard cap of 150 results per row
- **Language**: Python with pandas, openpyxl, requests, python-dateutil
- **LLM**: OpenAI (GPT-4o-mini) for text generation
- **Output**: Single enriched `.xlsx` file + `run_log.json`

## Key Decisions

| Decision | Rationale | Outcome |
|----------|-----------|---------|
| Apify over direct scraping | Reliability, ToS compliance, managed infrastructure | -- Pending |
| OpenAI for text generation | User has API key, GPT-4o-mini is cheap and fast | -- Pending |
| 150 result cap per row | Cost control for Apify usage | -- Pending |
| Scoring formula with log scale | Prevents mega-viral outliers from dominating, rewards engagement ratio | -- Pending |

---
*Last updated: 2026-02-06 after initialization*
