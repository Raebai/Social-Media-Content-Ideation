# Project State

## Project Reference

See: .planning/PROJECT.md (updated 2026-02-06)

**Core value:** One command transforms a bare content calendar into an actionable production sheet with audio picks, proven examples, and remix ideas.
**Current focus:** Phase 3 - Data Processing & Selection

## Current Position

Phase: 3 of 5 (Data Processing & Selection)
Plan: 2 of 2 in current phase
Status: Phase complete
Last activity: 2026-02-07 — Completed 03-02-PLAN.md

Progress: [█████░░░░░] 50%

## Performance Metrics

**Velocity:**
- Total plans completed: 5
- Average duration: 2.25 min
- Total execution time: 11.25 min

**By Phase:**

| Phase | Plans | Total | Avg/Plan |
|-------|-------|-------|----------|
| 1. Input Reading | 1 | 2.25 min | 2.25 min |
| 2. Query & API Integration | 2 | 4.58 min | 2.29 min |
| 3. Data Processing & Selection | 2 | 4.42 min | 2.21 min |

**Recent Trend:**
- Last 5 plans: 02-01 (2.40 min), 02-02 (2.18 min), 03-01 (2.05 min), 03-02 (2.37 min)
- Trend: Consistent velocity ~2.2-2.4 min/plan, stable performance

*Updated after each plan completion*

## Accumulated Context

### Decisions

Decisions are logged in PROJECT.md Key Decisions table.
Recent decisions affecting current work:

| Phase | Decision | Rationale | Impact |
|-------|----------|-----------|--------|
| 01-01 | Used openpyxl instead of pandas | Lighter dependency footprint for Excel reading | Faster install, smaller deployment |
| 01-01 | Case-insensitive Type validation with casing preservation | Handles acronyms like BTS correctly | Robust validation without breaking on edge cases |
| 01-01 | Preserved original row numbers (1-indexed) | Better debugging and error reporting | Error messages can reference Excel row numbers |
| 02-01 | Use format-specific queries from TYPE_QUERY_MAP | Ensures content-type alignment for TikTok search | Each idea has at least one query matching its format |
| 02-01 | 5-word query limit | TikTok search optimization | Queries stay focused and effective |
| 02-01 | Use searchQueries (not hashtags/profiles) for Apify | Best fit for generated queries | Clean API integration in Plan 02-02 |
| 02-02 | Use synchronous Apify endpoint | Simpler than async polling for batch processing | Cleaner code, adequate performance |
| 02-02 | Module-level dict for query caching | No need for external cache in single-run model | Zero infrastructure dependency |
| 02-02 | APIFY_TOKEN environment variable | Standard security practice for credentials | Secure, clear error on missing token |
| 03-01 | Log10 scale for engagement metrics | Prevents mega-viral outliers from dominating scores | Balanced scoring across view ranges |
| 03-01 | Keep videos with missing create_time | Avoids false negatives from incomplete Apify data | More robust results, no good examples lost |
| 03-01 | Relevance boost capped at +0.2 | Prevents keyword stuffing from inflating scores | Meaningful but not dominant relevance factor |
| 03-01 | Recency boost tiers (14d/30d) | TikTok algorithm favors recent content | Balances recency with quality metrics |
| 03-02 | Top 3 examples from highest-scored videos | Proven examples of what resonates with audience | Content creators have reference material for creation |
| 03-02 | Audio selection from top 20 using Counter | Find most common audio across broader result set | Better pattern detection than top 3 alone |
| 03-02 | Three-tier confidence: high/medium/low | Transparent recommendation strength signaling | Creators know when to trust audio recommendation |
| 03-02 | Fallback to example 1 audio when no repeats | Best single video likely has good audio choice | Graceful degradation when no clear pattern |

### Pending Todos

[From .planning/todos/pending/ — ideas captured during sessions]

None yet.

### Blockers/Concerns

[Issues that affect future work]

None yet.

## Session Continuity

Last session: 2026-02-07T18:17:45Z
Stopped at: Completed 03-02-PLAN.md (Phase 3 complete)
Resume file: None
