# Phase 1: Input Reading - Context

**Gathered:** 2026-02-06
**Status:** Ready for planning

<domain>
## Phase Boundary

Parse `Content ideas.xlsx` Sheet1 into structured idea data ready for enrichment. The file has 30 rows with columns Day, Date, Type, Topic, Description. Each row becomes an idea with a constructed `idea_text`. No API calls, no output files — just clean, validated data in memory.

</domain>

<decisions>
## Implementation Decisions

### Dirty data handling
- Normalize all Unicode: replace non-breaking hyphens (U+2011) with regular hyphens, non-breaking spaces (U+00A0) with regular spaces
- Parse date strings into proper date objects (dates are stored as strings like "2026-02-06" with non-breaking hyphens, not Excel date objects)
- Strip leading/trailing whitespace AND collapse multiple internal spaces into one for all text fields
- Validate Type values against the known set: Story, BTS, Teach, Trend, Breakdown, Reflection, Depth — error if unknown Type found

### Validation & errors
- All 5 columns (Day, Date, Type, Topic, Description) are required — fail fast with error if any column is missing from the header
- Skip rows where any of the key fields (Type, Topic, Description) are empty
- Log each skipped row with its row number and which field was missing
- Error messages should be detailed and helpful: show expected vs found, suggest fixes (e.g., "Expected column 'Description', found columns: Day, Date, Type, Topic, Desc")

### Data carried forward
- Preserve original Excel row number with each idea (for debugging and mapping back to source)
- Don't carry the Day column forward — Date is sufficient for ordering
- Store cleaned text only — no need to preserve pre-normalization originals
- Data held in memory only as Python objects — no intermediate JSON files

### idea_text construction
- Use pipe separator: `idea_text = f"{Type} | {Topic} | {Description}"`
- Normalize case: title-case the Type, sentence-case Topic and Description
- Print a load summary after parsing: count of ideas loaded, unique types found, date range
- Print details for any skipped rows (row number + missing field)

### Claude's Discretion
- Sheet name fallback strategy (try Sheet1, fall back to first sheet if missing)
- Exact Python data structure (dict, dataclass, namedtuple)
- Date parsing library choice
- Column matching strategy (exact match vs case-insensitive)

</decisions>

<specifics>
## Specific Ideas

- The actual Excel file has 30 data rows, 7 content types, all cells currently populated
- Dates in the file use non-breaking hyphens (U+2011) — normalization must handle this before date parsing
- Some text cells contain non-breaking spaces (U+00A0) — e.g., "Week\u00a01 Vlog"
- The requirement formula was `f"{Type}. {Topic}. {Description}"` but user prefers pipe separator instead

</specifics>

<deferred>
## Deferred Ideas

None — discussion stayed within phase scope

</deferred>

---

*Phase: 01-input-reading*
*Context gathered: 2026-02-06*
