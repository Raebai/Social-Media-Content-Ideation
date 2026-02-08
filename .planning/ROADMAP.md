# Roadmap: Logara Content Calendar Enricher

## Overview

This roadmap transforms a bare content calendar into an actionable production sheet through five phases: reading Excel input, generating queries and fetching TikTok data via Apify, processing and selecting top examples and audio, enhancing with OpenAI-generated creative text, and packaging as a complete CLI tool with enriched Excel output.

## Phases

**Phase Numbering:**
- Integer phases (1, 2, 3): Planned milestone work
- Decimal phases (2.1, 2.2): Urgent insertions (marked with INSERTED)

Decimal phases appear between their surrounding integers in numeric order.

- [x] **Phase 1: Input Reading** - Parse Excel content calendar into idea structures
- [x] **Phase 2: Query & API Integration** - Generate search queries and fetch TikTok data via Apify
- [x] **Phase 3: Data Processing & Selection** - Score, filter, and select top examples and audio
- [x] **Phase 4: LLM Enhancement** - Generate creative text with OpenAI
- [x] **Phase 5: Output & Packaging** - Write enriched Excel and package as CLI tool

## Phase Details

### Phase 1: Input Reading
**Goal**: Parse content calendar Excel into structured idea data ready for enrichment
**Depends on**: Nothing (first phase)
**Requirements**: IO-01, IO-02
**Success Criteria** (what must be TRUE):
  1. Tool reads `Content ideas.xlsx` Sheet1 with columns Day, Date, Type, Topic, Description
  2. Tool constructs idea_text combining Type, Topic, and Description per row
  3. All rows from input file are loaded into memory as structured data
**Plans**: 1 plan

Plans:
- [x] 01-01-PLAN.md -- Build Excel parser with ContentIdea dataclass, validation, Unicode normalization, and idea_text construction

### Phase 2: Query & API Integration
**Goal**: Generate targeted search queries and fetch TikTok results via Apify scraper
**Depends on**: Phase 1
**Requirements**: QRY-01, QRY-02, API-01, API-02, API-03, API-04, API-05, API-06
**Success Criteria** (what must be TRUE):
  1. Tool generates 3-6 short search queries per row based on idea_text
  2. Tool includes format-specific query based on Type mapping (Story, BTS, Teach, etc.)
  3. Tool fetches 20-30 TikTok results per query via Apify clockworks/tiktok-scraper
  4. Tool enforces hard cap of 150 total results per row
  5. Tool retries failed API calls with exponential backoff and caches successful results
**Plans**: 2 plans

Plans:
- [x] 02-01-PLAN.md -- Query generation (generate_queries + TYPE_QUERY_MAP) and Apify actor schema discovery
- [x] 02-02-PLAN.md -- Apify integration (fetch_tiktok_results with retry, caching, cap) and pipeline wiring

### Phase 3: Data Processing & Selection
**Goal**: Transform raw TikTok data into scored, filtered results with top examples and audio
**Depends on**: Phase 2
**Requirements**: DAT-01, DAT-02, DAT-03, DAT-04, DAT-05, DAT-06, DAT-07, EX-01, EX-02, AUD-01, AUD-02, AUD-03, AUD-04
**Success Criteria** (what must be TRUE):
  1. Tool normalizes TikTok results to consistent schema with video_id, url, caption, author, metrics, audio
  2. Tool deduplicates by video_id and filters videos older than 120 days
  3. Tool scores each video using engagement metrics (views, likes, comments) with recency and relevance boosts
  4. Tool selects top 3 videos by final score as examples per row
  5. Tool identifies most common audio among top 20 results with confidence level (high/medium/low)
**Plans**: 2 plans

Plans:
- [x] 03-01-PLAN.md -- Normalize, deduplicate, filter, and score TikTok results (process_results pipeline)
- [x] 03-02-PLAN.md -- Select top 3 examples and audio with confidence level (enrich_row orchestrator)

### Phase 4: LLM Enhancement
**Goal**: Generate creative text for audio fit reasoning, hook summaries, and remix ideas
**Depends on**: Phase 3
**Requirements**: LLM-01, LLM-02, LLM-03, LLM-04, LLM-05
**Success Criteria** (what must be TRUE):
  1. Tool uses OpenAI GPT-4o-mini to generate audio_fit_reason (1 sentence) per row
  2. Tool generates hook_summary (1 sentence) for each example (ex1, ex2, ex3)
  3. Tool generates remix_ideas (2-3 bullets) in hopecore + narrative stakes style matching row Type
  4. Tool extracts and writes audio_title for each example
**Plans**: 1 plan

Plans:
- [x] 04-01-PLAN.md -- OpenAI GPT-4o-mini integration with generate_llm_content function and enrich_row pipeline wiring

### Phase 5: Output & Packaging
**Goal**: Write enriched Excel with all columns, implement CLI interface, and package as complete tool
**Depends on**: Phase 4
**Requirements**: IO-03, IO-04, IO-05, QRY-03, LOG-01, LOG-02, LOG-03, PKG-01, PKG-02
**Success Criteria** (what must be TRUE):
  1. Tool writes enriched Excel with all output columns (queries, examples, audio, LLM text, status)
  2. Tool normalizes dates to consistent format in output Excel
  3. Tool accepts --input and --output CLI flags for file paths
  4. Tool tracks enrich_status per row (ok, partial, skipped, error) with reason
  5. Tool produces run_log.json with per-row queries, results count, chosen audio, example URLs
  6. Single enrich_calendar.py file runs end-to-end with requirements.txt dependencies
**Plans**: 2 plans

Plans:
- [x] 05-01-PLAN.md -- Excel output writer (write_enriched_excel) and run log builder (build_run_log)
- [x] 05-02-PLAN.md -- CLI interface with argparse, main loop, and end-to-end wiring

## Progress

**Execution Order:**
Phases execute in numeric order: 1 -> 2 -> 3 -> 4 -> 5

| Phase | Plans Complete | Status | Completed |
|-------|----------------|--------|-----------|
| 1. Input Reading | 1/1 | Complete | 2026-02-07 |
| 2. Query & API Integration | 2/2 | Complete | 2026-02-07 |
| 3. Data Processing & Selection | 2/2 | Complete | 2026-02-07 |
| 4. LLM Enhancement | 1/1 | Complete | 2026-02-08 |
| 5. Output & Packaging | 2/2 | Complete | 2026-02-08 |
