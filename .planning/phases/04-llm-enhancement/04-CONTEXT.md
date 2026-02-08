# Phase 4: LLM Enhancement - Context

**Gathered:** 2026-02-08
**Status:** Ready for planning

<domain>
## Phase Boundary

Generate creative text via OpenAI GPT-4o-mini for three output fields: audio_fit_reason (1 sentence), hook_summary per example (2 sentences each), and remix_ideas (2-3 bullets). Integrate LLM calls into the existing enrich_row() pipeline. Brand info upload and dynamic brand switching are future capabilities.

</domain>

<decisions>
## Implementation Decisions

### Prompt tone & style
- Voice: Hype coach — energetic, direct, motivational ("Open with your face, drop the hook in 2 seconds, hit them with the transformation")
- Hopecore flavor: Light — keep it optimistic and authentic, but don't force hopecore language into every idea
- Hook summaries: Context-aware, referencing what the example video actually did (using caption/metadata)
- Audio fit reason: General vibe description, not tied to the specific content idea ("Trending motivational track with 1.5M uses, high energy, good for visual transitions")

### Output structure
- remix_ideas: 2-3 one-liner bullets, punchy and scannable
- hook_summary: Two sentences per example — what they did + why it works/what to learn
- audio_fit_reason: One sentence max describing the audio's vibe and relevance
- Excel cell format for remix ideas: Newline-separated bullets with bullet prefix

### Content personalization
- Type-specific remix ideas: Yes — Teach gets educational angles, Story gets narrative arcs, BTS gets behind-the-scenes framing, etc.
- Brand context: Logara-aware — prompt includes startup founder context, building journey, founder content narrative
- Idea context passed to LLM: Type + Topic (skip Description to keep prompts shorter)
- Audio confidence affects tone: High confidence = assertive ("Use this track"), Low confidence = suggestive ("Consider this audio, but feel free to pick your own")

### Error & fallback behavior
- LLM failure: Mark enrich_status as 'partial', leave LLM fields empty (row retains example/audio data)
- Retries: 2 retries (3 total attempts) with exponential backoff, matching Apify retry pattern
- API call strategy: One call per row generating all fields (audio_fit_reason + hooks + remix_ideas)
- Bad output: Accept and best-effort parse — extract what we can, leave the rest empty, don't burn another API call

### Claude's Discretion
- Exact prompt wording and structure
- JSON vs structured text output format from OpenAI
- Temperature and model parameters
- How to extract/parse individual fields from single API response

</decisions>

<specifics>
## Specific Ideas

- User wants brand info upload capability eventually, but for now hardcode Logara context (startup building journey, founder documenting the process)
- Hype coach voice example: "Open with your face, drop the hook in 2 seconds, hit them with the transformation"
- Remix ideas should be filmable — concrete actions, not abstract advice

</specifics>

<deferred>
## Deferred Ideas

- Brand info upload / dynamic brand switching — future capability beyond Phase 4 scope
- Customizable prompt templates — future enhancement

</deferred>

---

*Phase: 04-llm-enhancement*
*Context gathered: 2026-02-08*
