# Phase 5: Output & Packaging - Context

**Gathered:** 2026-02-08
**Status:** Ready for planning

<domain>
## Phase Boundary

Write enriched Excel with all output columns, implement CLI interface with --input and --output flags, track enrich_status per row, produce run_log.json, and package as a single enrich_calendar.py file with requirements.txt. Configurable search parameters and content type suggestions are future capabilities.

</domain>

<decisions>
## Implementation Decisions

### Excel output layout
- Columns grouped by concern: input fields, then queries, then examples (ex1/ex2/ex3 blocks), then audio, then LLM text, then status
- Styled headers: bold, auto-fit column widths, frozen header row
- Default output filename: `Enriched Content ideas_2026-02-08.xlsx` (enriched name + timestamp to support multiple runs without overwrite)
- Date format in output: normalized to consistent format

### Cell formatting
- Claude's Discretion: how to handle multi-line content (remix_ideas bullets) in Excel cells -- pick what works best for readability

### CLI experience
- Use argparse for CLI parsing (auto-generated --help, proper flag handling)
- Default --input to `Content ideas.xlsx` in current directory (zero-config for typical workflow)
- --output defaults to timestamped enriched name in same directory
- Row-by-row progress printed to console: `Row 3/30: Teach | Building MVP... ok`
- Support --dry-run flag: validate input and show what would be processed, skip Apify and OpenAI calls

### End-to-end run flow
- Fail fast on startup if APIFY_TOKEN or OPENAI_API_KEY is missing -- check both before processing any rows
- On per-row failure (Apify timeout, etc.): continue processing remaining rows, mark failed row with enrich_status='error'
- Sequential processing: one row at a time (simple, respects API rate limits naturally)
- Write Excel all at once at the end (collect results in memory, write final file)

### Run log structure
- Detailed per-row logging: queries used, total results count, scored count, chosen audio, all 3 example URLs, LLM status, timing
- Top-level run summary: timestamp, input file, total rows, ok/partial/error/skipped counts, total duration
- Saved in same directory as output Excel
- Timestamped filename: `run_log_2026-02-08.json` (supports multiple runs)

### Claude's Discretion
- Exact column header names and ordering within each concern group
- Excel cell width sizing and text wrapping approach
- Exact console output format and color/emoji usage
- How to handle --dry-run output display
- run_log.json internal structure and field naming

</decisions>

<specifics>
## Specific Ideas

- Output filename pattern: "Enriched [original name]_[YYYY-MM-DD].xlsx"
- Run log filename pattern: "run_log_[YYYY-MM-DD].json"
- Progress should show content type and topic so user knows what's being processed
- --dry-run is for testing without burning API credits

</specifics>

<deferred>
## Deferred Ideas

- Configurable Apify search parameters (hashtags, search sections, etc.) -- future phase for customizable search
- Tool suggests what to specify per content type (recommended search parameters based on Type) -- future phase for intelligent defaults
- Customizable prompt templates for LLM generation -- captured in Phase 4 deferred ideas
- Brand info upload / dynamic brand switching -- captured in Phase 4 deferred ideas

</deferred>

---

*Phase: 05-output-packaging*
*Context gathered: 2026-02-08*
