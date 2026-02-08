---
phase: 05-output-packaging
plan: 02
subsystem: cli-interface
tags: [cli, argparse, main-loop, error-handling]

dependencies:
  requires:
    - "05-01: Data flow fixes and output functions"
    - "04-01: LLM content generation"
    - "03-02: Example and audio selection"
    - "02-02: Apify API integration"
    - "01-01: Excel input parsing"
  provides:
    - "Complete CLI tool with --input, --output, --dry-run flags"
    - "Production-ready main() entry point"
    - "Row-by-row processing with progress output"
    - "Per-row error handling with continue-on-error"
  affects:
    - "Users can now run `python enrich_calendar.py --input \"Content ideas.xlsx\"` to enrich calendars"
    - "Future: CLI could be extended with additional flags (--verbose, --max-results, etc.)"

tech-stack:
  added: []
  patterns:
    - "argparse for CLI argument parsing"
    - "Continue-on-error pattern for per-row processing"
    - "Automatic output path generation with timestamps"

key-files:
  created: []
  modified:
    - path: "enrich_calendar.py"
      changes:
        - "Added argparse import and CLI argument parsing"
        - "Implemented main() function with full pipeline orchestration"
        - "Replaced demo __main__ code with production CLI"
        - "Added environment variable validation"
        - "Removed temporary test code from Plan 05-01"
        - "Deleted test artifacts (test_output.xlsx, test_run_log.json)"

decisions:
  - what: "Continue-on-error for per-row processing"
    why: "One bad row shouldn't stop entire enrichment run"
    impact: "Users get partial results even with some failures; errors logged to status"
    phase: "05-02"

  - what: "Automatic output path with timestamp"
    why: "Prevents accidental overwrites; clear chronological ordering"
    impact: "Output files named like 'Enriched Content ideas_2026-02-08.xlsx'"
    phase: "05-02"

  - what: "Fail-fast on missing environment variables"
    why: "Better to error immediately than fail mid-processing"
    impact: "Clear error messages guide users to set APIFY_TOKEN and OPENAI_API_KEY"
    phase: "05-02"

  - what: "--dry-run skips environment variable check"
    why: "Users can validate input without API credentials"
    impact: "Easier onboarding and input testing"
    phase: "05-02"

metrics:
  duration: "2.9 min"
  completed: "2026-02-08"

status: complete
---

# Phase 5 Plan 02: CLI Interface & Main Loop Summary

**One-liner:** Production CLI with argparse, row-by-row processing, error handling, and automatic timestamped output

## What Was Built

This plan completed the final integration of the entire enrichment pipeline into a production-ready CLI tool. All functions from Plans 01-01 through 05-01 are now wired together in a single `main()` entry point.

### Key Components

**1. CLI Interface (argparse)**
- `--input`: Path to input Excel file (default: "Content ideas.xlsx")
- `--output`: Path to output Excel file (default: auto-generated with timestamp)
- `--dry-run`: Validate input and show rows without making API calls

**2. Main Loop Orchestration**
The `main()` function implements the full pipeline:
1. Parse command-line arguments
2. Validate input file exists
3. Check environment variables (APIFY_TOKEN, OPENAI_API_KEY) unless --dry-run
4. Load content ideas from Excel
5. Process each row sequentially:
   - Generate queries
   - Fetch TikTok results
   - Enrich row (process + select + LLM)
   - Print progress
   - Handle errors gracefully
6. Write enriched Excel output
7. Save run log JSON
8. Print summary statistics

**3. Error Handling**
- **Per-row errors:** Caught, logged as `enrich_status="error"`, processing continues
- **Missing env vars:** Fail-fast with clear error messages
- **Missing input:** Fail-fast with "file not found" error
- **Errors don't cascade:** One bad row doesn't affect others

**4. Output Path Generation**
When `--output` not specified:
- Derives from input filename
- Format: `Enriched {base_name}_{YYYY-MM-DD}.xlsx`
- Example: `Content ideas.xlsx` → `Enriched Content ideas_2026-02-08.xlsx`
- Run log saved alongside: `run_log_2026-02-08.json`

**5. Progress Output**
Console shows row-by-row progress:
```
Row 1/30: Story | Why I Left Singapore... ok
Row 2/30: BTS | Airport / First London Walk... partial
Row 3/30: Teach | Big Decision Framework... ok
```

Final summary:
```
Done! 30 rows in 45.2s
  ok: 25 | partial: 3 | skipped: 1 | error: 1
```

## Implementation Details

### Module Docstring
Updated to reflect production usage:
```python
"""
Content calendar enrichment tool.

Reads a content calendar Excel file, enriches each row with TikTok trends,
audio recommendations, and AI-generated creative text. Outputs enriched Excel
and diagnostic run log.

Usage: python enrich_calendar.py --input "Content ideas.xlsx"
"""
```

### Imports
All required imports present and ordered:
- `argparse` (new)
- `datetime`, `json`, `math`, `os`, `sys`, `time`
- `collections.Counter`, `dataclasses.dataclass`, `typing.*`
- `openpyxl` (load_workbook, Workbook, Font, Alignment, get_column_letter)
- `dateutil.parser.parse`, `requests`, `openai`

### Error Fallback Structure
The exception handler creates a minimal enrichment dict with:
- All required fields (row_number, content_type, topic, topic_keywords, etc.)
- Empty values for examples, audio, LLM content
- `enrich_status="error"` with error message in `enrich_reason`
- Structure matches enrich_row() output for uniform handling by write_enriched_excel()

### OpenAI Client Reuse
OpenAI client created once before the loop, passed to each enrich_row() call:
```python
openai_client = openai.OpenAI()
for i, idea in enumerate(ideas, 1):
    enriched = enrich_row(idea, raw_results, openai_client=openai_client)
```
This avoids unnecessary client instantiation overhead.

## Testing Results

All verification checks passed:

1. `python enrich_calendar.py --help` → Shows all 3 flags with descriptions
2. `python enrich_calendar.py --dry-run` → Lists all 30 rows, exits 0
3. `python enrich_calendar.py --dry-run --input nonexistent.xlsx` → Error, exits 1
4. `python -c "import enrich_calendar"` → No import errors
5. enrich_status values → Only ok/partial/skipped/error (matches LOG-01)
6. enrich_row return dict → Contains topic_keywords
7. Error fallback dict → Contains topic_keywords

## Cleanup Performed

**Removed:**
- 230+ lines of demo/test code from Plan 05-01
- Synthetic data generation for examples
- Phase 3 pipeline demonstration
- Test file generation (test_output.xlsx, test_run_log.json)

**Result:** Clean production code with single `main()` call in `if __name__ == "__main__"` block.

## Deviations from Plan

None - plan executed exactly as written.

## File Changes

### Modified: `enrich_calendar.py`

**Changes:**
1. Added `import argparse` at top
2. Updated module docstring with usage example
3. Implemented `main()` function (133 lines) with:
   - argparse setup
   - Input validation
   - Environment variable checks
   - Row-by-row processing loop
   - Error handling with continue-on-error
   - Output path generation
   - File writing and summary output
4. Replaced entire `__main__` block with `main()` call
5. Removed 230+ lines of demo/test code

**Final state:** Single-file CLI tool ready for production use

## Next Phase Readiness

**Phase 5 Complete!** All objectives achieved:
- ✅ Data flow fixes (Plan 05-01)
- ✅ Output functions (Plan 05-01)
- ✅ CLI interface (Plan 05-02)
- ✅ Main loop orchestration (Plan 05-02)

**Project Complete:** `enrich_calendar.py` is now a fully functional tool that transforms content calendars into enriched production sheets with:
- TikTok trend data
- Top 3 proven examples
- Audio recommendations with confidence levels
- AI-generated creative text (hook summaries, remix ideas, audio fit reasoning)
- Diagnostic run logs

**Usage:**
```bash
# Set environment variables
export APIFY_TOKEN="your-token"
export OPENAI_API_KEY="your-key"

# Run enrichment
python enrich_calendar.py --input "Content ideas.xlsx"

# Outputs:
# - Enriched Content ideas_2026-02-08.xlsx
# - run_log_2026-02-08.json
```

## Commits

1. **859474e** - `feat(05-02): implement CLI with argparse and main() function`
   - Add argparse with --input, --output, --dry-run flags
   - Implement main() function with full pipeline orchestration
   - Replace demo __main__ code with production CLI
   - Environment variable validation (APIFY_TOKEN, OPENAI_API_KEY)
   - Row-by-row processing with progress output
   - Per-row error handling with continue-on-error logic
   - Automatic output path generation with timestamp
   - Summary output with status counts

2. **4c39cc6** - `chore(05-02): clean up imports and remove test artifacts`
   - Module docstring updated with usage example
   - All imports present and ordered (argparse, openpyxl.Workbook, Font, Alignment, get_column_letter)
   - __main__ block contains only main() call
   - Removed test_output.xlsx and test_run_log.json artifacts

3. **96e0ee9** - `test(05-02): verify end-to-end CLI functionality`
   - All verification checks pass
   - --help shows all 3 flags with descriptions
   - --dry-run lists all 30 rows and exits 0
   - Missing file error returns exit code 1
   - No import errors
   - enrich_status values: ok/partial/skipped/error only
   - topic_keywords in enrich_row return dict and error fallback dict
   - Ready for real API runs

## Lessons Learned

**1. Continue-on-error is essential for batch processing**
When processing 30 rows, one API failure shouldn't stop the entire run. Per-row error handling with graceful degradation gives users partial results.

**2. Timestamped output prevents overwrites**
Auto-generating output paths with dates prevents accidental overwrites and provides clear chronological ordering.

**3. --dry-run enables early validation**
Users can validate input files without needing API credentials, making onboarding easier.

**4. Single client reuse matters**
Creating the OpenAI client once and reusing it across all rows avoids unnecessary instantiation overhead.

**5. Consistent error dict structure is critical**
The error fallback dict in main() must match enrich_row() output structure exactly, including topic_keywords, so write_enriched_excel() handles it uniformly.
