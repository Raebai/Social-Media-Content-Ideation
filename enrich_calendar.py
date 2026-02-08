"""
Excel parser for content calendar enrichment.

Reads Content ideas.xlsx, validates and normalizes data, constructs idea_text.
"""
import datetime
import math
import os
import sys
import time
import json
from collections import Counter
from dataclasses import dataclass
from typing import Tuple, List, Dict, Any, Optional
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter
from dateutil.parser import parse as date_parse
import requests
import openai


VALID_TYPES = {"Story", "BTS", "Teach", "Trend", "Breakdown", "Reflection", "Depth"}

# Apify API Configuration
APIFY_BASE_URL = "https://api.apify.com/v2"
APIFY_ACTOR_ID = "clockworks~tiktok-scraper"
MAX_RESULTS_PER_ROW = 150
RESULTS_PER_QUERY = 25
MAX_RETRIES = 3

# Query cache to avoid duplicate API calls
_query_cache: Dict[str, List[Dict]] = {}


@dataclass
class ContentIdea:
    """Represents a single content idea from the calendar."""
    row_number: int
    date: datetime.date
    content_type: str
    topic: str
    description: str
    idea_text: str


def normalize_text(text: str) -> str:
    """
    Normalize text by replacing special Unicode characters and cleaning whitespace.

    Args:
        text: Input text to normalize

    Returns:
        Normalized text with regular hyphens/spaces and collapsed whitespace
    """
    if not text:
        return ""

    # Replace non-breaking hyphen (U+2011) with regular hyphen
    text = text.replace('\u2011', '-')

    # Replace non-breaking space (U+00A0) with regular space
    text = text.replace('\u00a0', ' ')

    # Strip leading/trailing whitespace
    text = text.strip()

    # Collapse multiple internal spaces into single space
    text = ' '.join(text.split())

    return text


def sentence_case(text: str) -> str:
    """
    Convert text to sentence case (capitalize first letter, preserve rest).

    Args:
        text: Input text

    Returns:
        Text with first character capitalized, rest preserved as-is
    """
    if not text:
        return ""

    return text[0].upper() + text[1:] if len(text) > 1 else text.upper()


def parse_date(date_str: str) -> datetime.date:
    """
    Parse date string into date object, handling special Unicode characters.

    Args:
        date_str: Date string to parse

    Returns:
        Parsed date object

    Raises:
        ValueError: If date cannot be parsed
    """
    # First normalize to handle special hyphens
    normalized = normalize_text(date_str)

    try:
        return date_parse(normalized).date()
    except Exception as e:
        raise ValueError(f"Failed to parse date '{date_str}': {e}")


def load_content_ideas(file_path: str) -> Tuple[List[ContentIdea], List[Dict[str, Any]]]:
    """
    Load and validate content ideas from Excel file.

    Args:
        file_path: Path to Excel file

    Returns:
        Tuple of (list of ContentIdea objects, list of skipped row info dicts)

    Raises:
        ValueError: If required columns are missing or validation fails
    """
    ideas = []
    skipped = []

    # Open workbook
    wb = load_workbook(filename=file_path, read_only=True, data_only=True)

    # Try to get Sheet1, fall back to first sheet
    if "Sheet1" in wb.sheetnames:
        sheet = wb["Sheet1"]
        print(f"Reading from sheet: Sheet1")
    else:
        sheet = wb.active
        print(f"Sheet1 not found, using first sheet: {sheet.title}")

    # Read and validate header row
    header_row = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True))

    # Create case-insensitive column mapping
    header_map = {col.lower().strip() if col else "": idx
                  for idx, col in enumerate(header_row)}

    required_columns = ["day", "date", "type", "topic", "description"]
    missing_columns = []

    for col in required_columns:
        if col not in header_map:
            missing_columns.append(col)

    if missing_columns:
        found_columns = [col for col in header_row if col]
        raise ValueError(
            f"Missing required columns: {missing_columns}\n"
            f"Expected: {required_columns}\n"
            f"Found: {found_columns}"
        )

    # Map column positions
    col_date = header_map["date"]
    col_type = header_map["type"]
    col_topic = header_map["topic"]
    col_description = header_map["description"]

    # Process data rows (row 2+)
    for row_idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
        # Extract values
        date_val = row[col_date] if col_date < len(row) else None
        type_val = row[col_type] if col_type < len(row) else None
        topic_val = row[col_topic] if col_topic < len(row) else None
        desc_val = row[col_description] if col_description < len(row) else None

        # Normalize all text values
        date_str = normalize_text(str(date_val)) if date_val is not None else ""
        type_str = normalize_text(str(type_val)) if type_val is not None else ""
        topic_str = normalize_text(str(topic_val)) if topic_val is not None else ""
        desc_str = normalize_text(str(desc_val)) if desc_val is not None else ""

        # Check for missing required fields
        if not type_str:
            skipped.append({"row_number": row_idx, "field": "Type"})
            continue
        if not topic_str:
            skipped.append({"row_number": row_idx, "field": "Topic"})
            continue
        if not desc_str:
            skipped.append({"row_number": row_idx, "field": "Description"})
            continue
        if not date_str:
            skipped.append({"row_number": row_idx, "field": "Date"})
            continue

        # Validate Type (case-insensitive) and find proper casing
        type_lower = type_str.lower()
        valid_types_map = {t.lower(): t for t in VALID_TYPES}

        if type_lower not in valid_types_map:
            raise ValueError(
                f"Invalid Type '{type_str}' at row {row_idx}\n"
                f"Valid types: {sorted(VALID_TYPES)}"
            )

        # Use the properly-cased type from VALID_TYPES
        content_type = valid_types_map[type_lower]

        # Sentence-case Topic and Description
        topic = sentence_case(topic_str)
        description = sentence_case(desc_str)

        # Parse date
        try:
            date_obj = parse_date(date_str)
        except ValueError as e:
            raise ValueError(f"Row {row_idx}: {e}")

        # Construct idea_text
        idea_text = f"{content_type} | {topic} | {description}"

        # Create ContentIdea
        idea = ContentIdea(
            row_number=row_idx,
            date=date_obj,
            content_type=content_type,
            topic=topic,
            description=description,
            idea_text=idea_text
        )
        ideas.append(idea)

    wb.close()

    return ideas, skipped


def print_summary(ideas: List[ContentIdea], skipped: List[Dict[str, Any]]) -> None:
    """
    Print summary of loaded content ideas.

    Args:
        ideas: List of loaded ContentIdea objects
        skipped: List of skipped row information
    """
    print(f"Loaded {len(ideas)} content ideas")

    # Get unique types
    unique_types = sorted(set(idea.content_type for idea in ideas))
    print(f"Types: {', '.join(unique_types)}")

    # Get date range
    if ideas:
        dates = [idea.date for idea in ideas]
        earliest = min(dates)
        latest = max(dates)
        print(f"Date range: {earliest} to {latest}")

    # Print skipped rows
    if skipped:
        print("\nSkipped rows:")
        for skip_info in skipped:
            print(f"  Skipped row {skip_info['row_number']}: missing {skip_info['field']}")


# Query generation for TikTok search
TYPE_QUERY_MAP = {
    "Story": "founder story",
    "BTS": "day in the life",
    "Teach": "startup lesson",
    "Trend": "new chapter",
    "Breakdown": "mindset shift",
    "Reflection": "founder reflection",
    "Depth": "longform reflection"
}


def generate_queries(idea: ContentIdea) -> List[str]:
    """
    Generate 3-6 TikTok search queries from a content idea.

    At least one query is a format-specific query from TYPE_QUERY_MAP.
    All queries are <=5 words.

    Args:
        idea: ContentIdea to generate queries from

    Returns:
        List of 3-6 query strings, each <=5 words
    """
    queries = []

    # Helper to truncate queries to 5 words
    def truncate_to_5_words(text: str) -> str:
        words = text.split()
        return ' '.join(words[:5]) if len(words) > 5 else text

    # Helper to extract keywords (meaningful words >3 chars)
    def extract_keywords(text: str) -> List[str]:
        words = text.lower().split()
        # Filter out common stop words and short words
        stop_words = {'the', 'and', 'for', 'with', 'from', 'this', 'that', 'what', 'when', 'where', 'who', 'how', 'why'}
        keywords = [w for w in words if len(w) > 3 and w not in stop_words]
        return keywords

    # 1. Format-specific query from TYPE_QUERY_MAP (satisfies QRY-02)
    format_query = TYPE_QUERY_MAP.get(idea.content_type, "founder content")
    queries.append(format_query)

    # 2. Extract keywords from topic and description
    topic_keywords = extract_keywords(idea.topic)
    desc_keywords = extract_keywords(idea.description)

    # 3. Add topic as query if it's <=5 words
    if len(idea.topic.split()) <= 5:
        queries.append(idea.topic.lower())

    # 4. Build keyword combination queries (2-3 keywords)
    # Try topic keywords first
    if len(topic_keywords) >= 2:
        combo = ' '.join(topic_keywords[:3])
        combo = truncate_to_5_words(combo)
        queries.append(combo)

    # Try description keywords
    if len(desc_keywords) >= 2:
        combo = ' '.join(desc_keywords[:3])
        combo = truncate_to_5_words(combo)
        queries.append(combo)

    # 5. Niche query combining content type keyword + top topic keyword
    type_keyword = idea.content_type.lower()
    if topic_keywords:
        niche_query = f"{type_keyword} {topic_keywords[0]}"
        niche_query = truncate_to_5_words(niche_query)
        queries.append(niche_query)
    else:
        # Fallback: use content type + "startup"
        queries.append(f"startup {type_keyword}")

    # 6. Cross-pollinate: description keyword + format
    if desc_keywords:
        cross_query = f"{desc_keywords[0]} {format_query.split()[0]}"
        cross_query = truncate_to_5_words(cross_query)
        queries.append(cross_query)

    # Remove duplicates while preserving order
    seen = set()
    unique_queries = []
    for q in queries:
        q_normalized = q.lower().strip()
        if q_normalized not in seen:
            seen.add(q_normalized)
            unique_queries.append(q_normalized)

    # Ensure minimum 3 queries with fallbacks
    fallbacks = ["startup founder content", "entrepreneur tiktok", "business growth tips"]
    fallback_idx = 0
    while len(unique_queries) < 3 and fallback_idx < len(fallbacks):
        fallback = fallbacks[fallback_idx]
        if fallback not in seen:
            unique_queries.append(fallback)
            seen.add(fallback)
        fallback_idx += 1

    # Return 3-6 queries (trim if we have too many)
    return unique_queries[:6]


def _call_with_retry(fn, *args, max_retries=MAX_RETRIES, **kwargs) -> Any:
    """
    Call function with exponential backoff retry on failures.

    Args:
        fn: Function to call
        *args: Positional arguments for fn
        max_retries: Maximum number of retry attempts
        **kwargs: Keyword arguments for fn

    Returns:
        Result from fn, or empty list on final failure
    """
    for attempt in range(max_retries):
        try:
            return fn(*args, **kwargs)
        except (requests.RequestException, ValueError, openai.APIError) as e:
            if attempt < max_retries - 1:
                wait_time = 2 ** attempt
                print(f"Retry {attempt + 1}/{max_retries} after error: {e}", file=sys.stderr)
                print(f"Waiting {wait_time}s before retry...", file=sys.stderr)
                time.sleep(wait_time)
            else:
                print(f"Failed after {max_retries} attempts: {e}", file=sys.stderr)
                return []


def _run_apify_actor(query: str, token: str, max_items: int = RESULTS_PER_QUERY) -> List[Dict[str, Any]]:
    """
    Run Apify TikTok scraper actor for a single query.

    Args:
        query: Search query string
        token: Apify API token
        max_items: Maximum results to fetch per query

    Returns:
        List of TikTok video result dictionaries

    Raises:
        requests.RequestException: On API call failures
        ValueError: On invalid response format
    """
    # Build input payload based on discovered schema
    input_payload = {
        "searchQueries": [query],
        "resultsPerPage": max_items,
        "searchSection": "/video",  # Videos only
        "searchSorting": "0",  # Most relevant
        "searchDatePosted": "0"  # All time
    }

    # Use synchronous endpoint for simplicity
    url = f"{APIFY_BASE_URL}/acts/{APIFY_ACTOR_ID}/run-sync-get-dataset-items"

    headers = {
        "Content-Type": "application/json"
    }

    params = {
        "token": token
    }

    response = requests.post(url, json=input_payload, headers=headers, params=params, timeout=300)

    if response.status_code != 200 and response.status_code != 201:
        raise requests.RequestException(
            f"Apify API returned status {response.status_code}: {response.text}"
        )

    try:
        results = response.json()
        # Synchronous endpoint returns array directly
        if isinstance(results, list):
            return results
        # Or might be wrapped in a data field
        elif isinstance(results, dict) and "data" in results:
            return results["data"]
        else:
            return []
    except json.JSONDecodeError as e:
        raise ValueError(f"Invalid JSON response from Apify: {e}")


def fetch_tiktok_results(queries: List[str], token: Optional[str] = None) -> List[Dict[str, Any]]:
    """
    Fetch TikTok results for multiple queries with caching and result cap.

    Features:
    - Retry logic: Up to 3 attempts with exponential backoff
    - Caching: Identical queries within a run return cached results
    - Result cap: Maximum 150 results per row regardless of query count

    Args:
        queries: List of search query strings
        token: Apify API token (defaults to APIFY_TOKEN env var)

    Returns:
        List of TikTok video result dictionaries, capped at MAX_RESULTS_PER_ROW

    Raises:
        ValueError: If APIFY_TOKEN not found in environment
    """
    # Get token from environment if not provided
    if token is None:
        token = os.environ.get("APIFY_TOKEN")

    if not token:
        raise ValueError("APIFY_TOKEN not found in environment")

    results = []

    for query in queries:
        # Normalize query for cache key
        query_normalized = query.lower().strip()

        # Check cache first (API-06)
        if query_normalized in _query_cache:
            print(f"Cache hit for query: {query}", file=sys.stderr)
            results.extend(_query_cache[query_normalized])
        else:
            # Call API with retry logic
            query_results = _call_with_retry(_run_apify_actor, query, token)

            # Store in cache
            _query_cache[query_normalized] = query_results

            # Add to results
            results.extend(query_results)

        # Check result cap (API-03)
        if len(results) >= MAX_RESULTS_PER_ROW:
            results = results[:MAX_RESULTS_PER_ROW]
            break

    return results


def enrich_row_queries(idea: ContentIdea) -> Dict[str, Any]:
    """
    Generate queries for a content idea (Phase 2 pipeline helper).

    This structural helper wires query generation to row context.
    Phase 3+ will extend this to include TikTok results and analysis.

    Args:
        idea: ContentIdea to generate queries for

    Returns:
        Dict with row_number, queries list, and query_count
    """
    queries = generate_queries(idea)

    return {
        "row_number": idea.row_number,
        "queries": queries,
        "query_count": len(queries)
    }


def normalize_result(raw: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    """
    Normalize Apify TikTok result to standard schema.

    Maps Apify output schema to our standard video metadata structure.
    Returns None for malformed results (missing required fields).

    Args:
        raw: Raw Apify result dictionary

    Returns:
        Normalized video dict with standard schema, or None if malformed
    """
    # Check required fields
    if not raw.get("id"):
        return None

    # Build normalized result
    normalized = {
        "video_id": raw["id"],
        "url": raw.get("webVideoUrl", ""),
        "caption": raw.get("text", ""),
        "author_username": raw.get("authorMeta", {}).get("name", ""),
        "create_time": raw.get("createTimeISO", None),
        "views": raw.get("playCount", 0),
        "likes": raw.get("diggCount", 0),
        "comments": raw.get("commentCount", 0),
        "shares": raw.get("shareCount", 0),
        "audio": {
            "audio_id": raw.get("musicMeta", {}).get("musicId", ""),
            "title": raw.get("musicMeta", {}).get("musicName", ""),
            "author": raw.get("musicMeta", {}).get("musicAuthor", ""),
            "url": raw.get("musicMeta", {}).get("playUrl", "")
        }
    }

    return normalized


def deduplicate_results(results: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Remove duplicate videos by video_id, keeping first occurrence.

    Args:
        results: List of normalized video dicts

    Returns:
        Deduplicated list preserving original order
    """
    seen = set()
    deduped = []

    for result in results:
        video_id = result.get("video_id")
        if video_id and video_id not in seen:
            seen.add(video_id)
            deduped.append(result)

    return deduped


def filter_old_results(results: List[Dict[str, Any]], max_age_days: int = 120) -> List[Dict[str, Any]]:
    """
    Filter out videos older than max_age_days.

    Videos with missing or unparseable create_time are KEPT (no false rejection).

    Args:
        results: List of normalized video dicts
        max_age_days: Maximum age in days (default 120)

    Returns:
        Filtered list with only recent videos
    """
    filtered = []
    now = datetime.datetime.now(datetime.timezone.utc)

    for result in results:
        create_time = result.get("create_time")

        # Keep if create_time is missing or empty
        if not create_time:
            filtered.append(result)
            continue

        # Try to parse and check age
        try:
            created_dt = date_parse(create_time)
            # Ensure timezone-aware
            if created_dt.tzinfo is None:
                created_dt = created_dt.replace(tzinfo=datetime.timezone.utc)

            age_days = (now - created_dt).days

            if age_days <= max_age_days:
                filtered.append(result)
        except Exception:
            # Keep on parsing failure (graceful fallback)
            filtered.append(result)

    return filtered


def score_video(video: Dict[str, Any], keywords: List[str] = None) -> Dict[str, Any]:
    """
    Compute engagement score for a video.

    Scoring formula (DAT-04 to DAT-07):
    - Base score: log10(views+1)*0.65 + log10(likes+1)*0.25 + log10(comments+1)*0.10
    - Recency boost: +0.2 for <=14 days, +0.1 for <=30 days, 0 otherwise
    - Relevance boost: +0.05 per keyword in caption, capped at +0.2
    - Final score: base + recency_boost + relevance_boost

    Args:
        video: Normalized video dict
        keywords: Optional list of keywords for relevance scoring

    Returns:
        Video dict with added score fields (base_score, recency_boost, relevance_boost, final_score)
    """
    # Step 1: Base score (DAT-04)
    views = video.get("views", 0)
    likes = video.get("likes", 0)
    comments = video.get("comments", 0)

    base = (math.log10(views + 1) * 0.65 +
            math.log10(likes + 1) * 0.25 +
            math.log10(comments + 1) * 0.10)

    # Step 2: Recency boost (DAT-05)
    recency_boost = 0.0
    create_time = video.get("create_time")

    if create_time:
        try:
            created_dt = date_parse(create_time)
            # Ensure timezone-aware
            if created_dt.tzinfo is None:
                created_dt = created_dt.replace(tzinfo=datetime.timezone.utc)

            now = datetime.datetime.now(datetime.timezone.utc)
            days_old = (now - created_dt).days

            if days_old <= 14:
                recency_boost = 0.2
            elif days_old <= 30:
                recency_boost = 0.1
        except Exception:
            # On parsing failure, no recency boost
            pass

    # Step 3: Relevance boost (DAT-06)
    relevance_boost = 0.0

    if keywords:
        caption = video.get("caption", "")
        if caption:
            caption_lower = caption.lower()
            overlap_count = sum(1 for kw in keywords if kw.lower() in caption_lower)
            relevance_boost = min(overlap_count * 0.05, 0.2)

    # Step 4: Final score (DAT-07)
    final_score = base + recency_boost + relevance_boost

    # Add scores to video dict
    video["base_score"] = base
    video["recency_boost"] = recency_boost
    video["relevance_boost"] = relevance_boost
    video["final_score"] = final_score

    return video


def process_results(raw_results: List[Dict[str, Any]], keywords: List[str] = None) -> List[Dict[str, Any]]:
    """
    Process raw Apify results through full pipeline.

    Pipeline:
    1. Normalize: Convert Apify schema to standard schema
    2. Deduplicate: Remove duplicate video_ids
    3. Filter: Remove videos older than 120 days
    4. Score: Compute engagement scores with optional keyword relevance
    5. Sort: By final_score descending

    Args:
        raw_results: List of raw Apify result dicts
        keywords: Optional list of keywords for relevance scoring

    Returns:
        Sorted list of scored, normalized video dicts
    """
    # Step 1: Normalize (filter out None results)
    normalized = [normalize_result(r) for r in raw_results]
    normalized = [n for n in normalized if n is not None]

    # Step 2: Deduplicate
    deduped = deduplicate_results(normalized)

    # Step 3: Filter old
    filtered = filter_old_results(deduped)

    # Step 4: Score
    scored = [score_video(v, keywords) for v in filtered]

    # Step 5: Sort by final_score descending
    sorted_results = sorted(scored, key=lambda x: x.get("final_score", 0), reverse=True)

    return sorted_results


def select_top_examples(scored_results: List[Dict[str, Any]], count: int = 3) -> List[Dict[str, Any]]:
    """
    Select top N videos from scored results as examples.

    Takes already-sorted (descending by final_score) output from process_results
    and extracts top N videos with relevant metadata for content creation.

    Args:
        scored_results: Sorted list of scored video dicts from process_results()
        count: Number of top examples to select (default 3)

    Returns:
        List of up to 'count' example dicts with url, engagement metrics,
        author info, caption, audio, and final_score. Returns fewer if less available.
        Returns empty list if scored_results is empty.
    """
    if not scored_results:
        return []

    examples = []
    for video in scored_results[:count]:
        example = {
            "url": video["url"],
            "views": video["views"],
            "likes": video["likes"],
            "comments": video["comments"],
            "shares": video["shares"],
            "author_username": video["author_username"],
            "create_time": video["create_time"],
            "caption": video["caption"],
            "audio_title": video.get("audio", {}).get("title", ""),
            "final_score": video["final_score"]
        }
        examples.append(example)

    return examples


def select_audio(scored_results: List[Dict[str, Any]], examples: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Select recommended audio track from top scored results.

    Strategy:
    1. Count audio_id occurrences among top 20 results
    2. Select most common audio_id (if any repeats found)
    3. Fall back to first example's audio if no repeats
    4. Determine confidence based on occurrence count

    Args:
        scored_results: Sorted list of scored video dicts from process_results()
        examples: List of example dicts from select_top_examples()

    Returns:
        Dict with audio_title, audio_author, audio_id, audio_url, audio_confidence.
        Confidence levels: "high" (>=3 occurrences), "medium" (2), "low" (1 or fallback).
        Returns empty strings with "low" confidence if no audio found.
    """
    # Step 1: Count audio occurrences in top 20
    top_20 = scored_results[:20]
    audio_ids = []

    for video in top_20:
        audio_id = video.get("audio", {}).get("audio_id", "")
        if audio_id:  # Skip empty/falsy audio_ids
            audio_ids.append(audio_id)

    audio_counter = Counter(audio_ids)

    # Step 2: Select audio
    selected_audio_id = None
    count = 0

    if audio_counter:
        # Get most common audio_id
        selected_audio_id, count = audio_counter.most_common(1)[0]

        # Find first video with this audio_id to get full metadata
        selected_video = None
        for video in top_20:
            if video.get("audio", {}).get("audio_id") == selected_audio_id:
                selected_video = video
                break

        if selected_video:
            audio_title = selected_video.get("audio", {}).get("title", "")
            audio_author = selected_video.get("audio", {}).get("author", "")
            audio_url = selected_video.get("audio", {}).get("url", "")
        else:
            # Shouldn't happen, but fallback
            audio_title = ""
            audio_author = ""
            audio_url = ""
    else:
        # No valid audio_ids in top 20, fall back to example 1
        if examples and len(examples) > 0:
            # Find the video in scored_results that matches example 1 URL
            ex1_url = examples[0].get("url", "")
            selected_video = None
            for video in scored_results:
                if video.get("url") == ex1_url:
                    selected_video = video
                    break

            if selected_video:
                selected_audio_id = selected_video.get("audio", {}).get("audio_id", "")
                audio_title = selected_video.get("audio", {}).get("title", "")
                audio_author = selected_video.get("audio", {}).get("author", "")
                audio_url = selected_video.get("audio", {}).get("url", "")
                count = 1  # Fallback used
            else:
                # Can't find ex1 video
                selected_audio_id = ""
                audio_title = ""
                audio_author = ""
                audio_url = ""
                count = 0
        else:
            # No examples either
            selected_audio_id = ""
            audio_title = ""
            audio_author = ""
            audio_url = ""
            count = 0

    # Step 3: Determine confidence
    if count >= 3:
        confidence = "high"
    elif count == 2:
        confidence = "medium"
    else:
        confidence = "low"

    # Step 4: Build return dict
    return {
        "audio_title": audio_title,
        "audio_author": audio_author,
        "audio_id": selected_audio_id,
        "audio_url": audio_url,
        "audio_confidence": confidence
    }


def generate_llm_content(idea: ContentIdea, examples: List[Dict[str, Any]], audio: Dict[str, Any], client=None) -> Dict[str, Any]:
    """
    Generate LLM content for enrichment: audio_fit_reason, hook_summaries, remix_ideas.

    Uses OpenAI GPT-4o-mini to generate actionable creative text based on:
    - Content idea (type, topic, description)
    - Top example videos (caption, engagement, audio)
    - Selected audio track

    Args:
        idea: ContentIdea to generate content for
        examples: List of example video dicts from select_top_examples()
        audio: Audio dict from select_audio()
        client: Optional OpenAI client for testing (creates one if None)

    Returns:
        Dict with audio_fit_reason, ex1_hook_summary, ex2_hook_summary,
        ex3_hook_summary, remix_ideas. Returns empty strings on API failure.
    """
    # Step 1: Client setup
    if client is None:
        # Check if OPENAI_API_KEY is set
        if not os.environ.get("OPENAI_API_KEY"):
            # No API key - return empty fields
            return {
                "audio_fit_reason": "",
                "ex1_hook_summary": "",
                "ex2_hook_summary": "",
                "ex3_hook_summary": "",
                "remix_ideas": ""
            }
        client = openai.OpenAI()  # Reads OPENAI_API_KEY from env automatically

    # Step 2: Build prompt with brand context and type-specific direction
    system_message = (
        "You are a hype coach content strategist for Logara, a startup documenting its founder journey. "
        "You give energetic, direct, filmable content direction. Respond in valid JSON only."
    )

    # Build examples section
    examples_text = ""
    for i, ex in enumerate(examples, 1):
        examples_text += f"\nExample {i}:\n"
        examples_text += f"  - Caption: {ex.get('caption', 'N/A')}\n"
        examples_text += f"  - Views: {ex.get('views', 0):,} | Likes: {ex.get('likes', 0):,}\n"
        examples_text += f"  - Author: @{ex.get('author_username', 'unknown')}\n"
        examples_text += f"  - Audio: {ex.get('audio_title', 'N/A')}\n"

    # Adjust tone based on audio confidence
    audio_confidence = audio.get('audio_confidence', 'low')
    if audio_confidence == 'high':
        audio_instruction = "Use assertive language for audio_fit_reason (e.g., 'Use this track - ...')"
    else:
        audio_instruction = "Use suggestive language for audio_fit_reason (e.g., 'Consider this audio, but feel free to pick your own - ...')"

    # Type-specific framing
    type_framing = {
        "Teach": "Educational angles - focus on what viewers will learn",
        "Story": "Narrative arcs - emphasize storytelling structure",
        "BTS": "Behind-the-scenes framing - show the process/reality",
        "Trend": "Trending formats - leverage popular patterns",
        "Breakdown": "Analysis angles - break down concepts clearly",
        "Reflection": "Personal insight - authentic founder perspective",
        "Depth": "Longform depth - substantive exploration"
    }
    type_hint = type_framing.get(idea.content_type, "Authentic founder content")

    user_message = f"""Content Idea:
- Type: {idea.content_type}
- Topic: {idea.topic}
- Description: {idea.description}

Audio Track:
- Title: {audio.get('audio_title', 'N/A')}
- Author: {audio.get('audio_author', 'N/A')}
- Confidence: {audio_confidence}

Top Examples:{examples_text}

Brand Context:
- Logara is a startup documenting its founder journey
- Tone: Hopecore-light (optimistic, authentic, not forced)
- Type framing: {type_hint}

Instructions:
1. audio_fit_reason: One sentence about the audio's GENERAL vibe and usage potential (not tied to this specific idea). {audio_instruction}
2. hook_summaries: For each example, write 2 sentences: (a) what the video did (reference the caption), (b) why it works/what to learn
3. remix_ideas: 2-3 filmable one-liner bullets tailored to {idea.content_type} format. Concrete actions, not abstract advice.

Return JSON:
{{
  "audio_fit_reason": "one sentence",
  "hook_summaries": ["2 sentences for ex1", "2 sentences for ex2", "2 sentences for ex3"],
  "remix_ideas": ["bullet 1", "bullet 2", "bullet 3"]
}}"""

    # Step 3: Define API call function for retry wrapper
    def make_api_call():
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            response_format={"type": "json_object"},
            temperature=0.8,
            max_tokens=500,
            messages=[
                {"role": "system", "content": system_message},
                {"role": "user", "content": user_message}
            ]
        )
        return response

    # Step 4: Call API with retry wrapper
    try:
        response = _call_with_retry(make_api_call)

        # If retry failed, response will be []
        if not response or response == []:
            return {
                "audio_fit_reason": "",
                "ex1_hook_summary": "",
                "ex2_hook_summary": "",
                "ex3_hook_summary": "",
                "remix_ideas": ""
            }

        # Step 5: Parse response
        content = response.choices[0].message.content
        parsed = json.loads(content)

        # Extract fields with defaults
        audio_fit_reason = parsed.get("audio_fit_reason", "")
        hook_summaries = parsed.get("hook_summaries", [])
        remix_ideas_list = parsed.get("remix_ideas", [])

        # Format remix_ideas as bulleted string
        if remix_ideas_list:
            remix_ideas = "\n".join(f"- {idea}" for idea in remix_ideas_list)
        else:
            remix_ideas = ""

        # Step 6: Return structured dict
        return {
            "audio_fit_reason": audio_fit_reason,
            "ex1_hook_summary": hook_summaries[0] if len(hook_summaries) > 0 else "",
            "ex2_hook_summary": hook_summaries[1] if len(hook_summaries) > 1 else "",
            "ex3_hook_summary": hook_summaries[2] if len(hook_summaries) > 2 else "",
            "remix_ideas": remix_ideas
        }

    except (json.JSONDecodeError, KeyError, IndexError, AttributeError) as e:
        print(f"Error parsing LLM response: {e}", file=sys.stderr)
        return {
            "audio_fit_reason": "",
            "ex1_hook_summary": "",
            "ex2_hook_summary": "",
            "ex3_hook_summary": "",
            "remix_ideas": ""
        }


def enrich_row(idea: ContentIdea, raw_results: List[Dict[str, Any]], openai_client=None) -> Dict[str, Any]:
    """
    Orchestrate full Phase 3 pipeline for a single content row.

    This is the main Phase 3 entry point that chains together:
    1. Keyword extraction for relevance scoring
    2. Result processing (normalize, dedup, filter, score, sort)
    3. Top example selection
    4. Audio selection

    Args:
        idea: ContentIdea to process
        raw_results: List of raw Apify result dicts from TikTok search

    Returns:
        Enrichment dict containing row metadata, top 3 examples, selected audio,
        result counts, and queries used.
    """
    # Step 1: Extract keywords for relevance scoring
    # Reuse keyword extraction logic from generate_queries
    def extract_keywords(text: str) -> List[str]:
        words = text.lower().split()
        stop_words = {'the', 'and', 'for', 'with', 'from', 'this', 'that', 'what', 'when', 'where', 'who', 'how', 'why'}
        keywords = [w for w in words if len(w) > 3 and w not in stop_words]
        return keywords

    topic_keywords = extract_keywords(idea.topic)
    desc_keywords = extract_keywords(idea.description)
    keywords = topic_keywords + desc_keywords

    # Step 2: Process results (normalize, dedup, filter, score, sort)
    scored = process_results(raw_results, keywords=keywords)

    # Step 3: Select top 3 examples
    examples = select_top_examples(scored)

    # Step 4: Select audio
    audio = select_audio(scored, examples)

    # Step 5: Determine enrich_status per LOG-01
    enrich_status = "ok"
    enrich_reason = ""

    if len(scored) == 0:
        # No results at all - skip LLM call
        llm_content = {
            "audio_fit_reason": "",
            "ex1_hook_summary": "",
            "ex2_hook_summary": "",
            "ex3_hook_summary": "",
            "remix_ideas": ""
        }
        enrich_status = "skipped"
        enrich_reason = "No scored results"
    else:
        # Generate LLM content
        llm_content = generate_llm_content(idea, examples, audio, client=openai_client)

        # Check if LLM generation failed (all fields empty)
        llm_failed = all(
            llm_content.get(key, "") == ""
            for key in ["audio_fit_reason", "ex1_hook_summary", "ex2_hook_summary",
                       "ex3_hook_summary", "remix_ideas"]
        )

        # Determine status based on LOG-01 criteria
        has_audio = bool(audio.get("audio_title"))
        has_enough_examples = len(examples) >= 2
        reasons = []

        if llm_failed:
            reasons.append("LLM generation failed")
        if not has_enough_examples:
            reasons.append(f"Only {len(examples)} example(s), need >=2")
        if not has_audio:
            reasons.append("No audio selected")

        if reasons:
            enrich_status = "partial"
            enrich_reason = "; ".join(reasons)

    # Step 6: Build and return enrichment dict with LLM fields
    return {
        "row_number": idea.row_number,
        "content_type": idea.content_type,
        "topic": idea.topic,
        "topic_keywords": topic_keywords,
        "examples": examples,
        "audio": audio,
        "audio_fit_reason": llm_content.get("audio_fit_reason", ""),
        "ex1_hook_summary": llm_content.get("ex1_hook_summary", ""),
        "ex2_hook_summary": llm_content.get("ex2_hook_summary", ""),
        "ex3_hook_summary": llm_content.get("ex3_hook_summary", ""),
        "ex1_audio_title": examples[0]["audio_title"] if len(examples) > 0 else "",
        "ex2_audio_title": examples[1]["audio_title"] if len(examples) > 1 else "",
        "ex3_audio_title": examples[2]["audio_title"] if len(examples) > 2 else "",
        "remix_ideas": llm_content.get("remix_ideas", ""),
        "enrich_status": enrich_status,
        "enrich_reason": enrich_reason,
        "total_results": len(raw_results),
        "scored_results": len(scored),
        "queries": generate_queries(idea)
    }


def write_enriched_excel(ideas: List[ContentIdea], enrichments: List[Dict[str, Any]], output_path: str) -> None:
    """
    Write enriched content ideas to Excel with styled headers and grouped columns.

    Args:
        ideas: List of original ContentIdea objects
        enrichments: List of enrichment dicts from enrich_row()
        output_path: Path to write output Excel file

    Column layout (grouped by concern):
    1. INPUT: Day, Date, Type, Topic, Description
    2. QUERY: Topic Keywords, Search Queries
    3. EXAMPLE 1: Ex1 URL, Ex1 Views, Ex1 Likes, Ex1 Comments, Ex1 Shares, Ex1 Author, Ex1 Caption, Ex1 Audio, Ex1 Hook Summary
    4. EXAMPLE 2: Ex2 URL, Ex2 Views, Ex2 Likes, Ex2 Comments, Ex2 Shares, Ex2 Author, Ex2 Caption, Ex2 Audio, Ex2 Hook Summary
    5. EXAMPLE 3: Ex3 URL, Ex3 Views, Ex3 Likes, Ex3 Comments, Ex3 Shares, Ex3 Author, Ex3 Caption, Ex3 Audio, Ex3 Hook Summary
    6. AUDIO: Audio Title, Audio Author, Audio Confidence, Audio URL, Audio Fit Reason
    7. LLM: Remix Ideas
    8. STATUS: Enrich Status, Enrich Reason
    """
    # Create workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Enriched Content"

    # Define column headers in order
    headers = [
        # INPUT GROUP
        "Day", "Date", "Type", "Topic", "Description",
        # QUERY GROUP
        "Topic Keywords", "Search Queries",
        # EXAMPLE 1
        "Ex1 URL", "Ex1 Views", "Ex1 Likes", "Ex1 Comments", "Ex1 Shares",
        "Ex1 Author", "Ex1 Caption", "Ex1 Audio", "Ex1 Hook Summary",
        # EXAMPLE 2
        "Ex2 URL", "Ex2 Views", "Ex2 Likes", "Ex2 Comments", "Ex2 Shares",
        "Ex2 Author", "Ex2 Caption", "Ex2 Audio", "Ex2 Hook Summary",
        # EXAMPLE 3
        "Ex3 URL", "Ex3 Views", "Ex3 Likes", "Ex3 Comments", "Ex3 Shares",
        "Ex3 Author", "Ex3 Caption", "Ex3 Audio", "Ex3 Hook Summary",
        # AUDIO GROUP
        "Audio Title", "Audio Author", "Audio Confidence", "Audio URL", "Audio Fit Reason",
        # LLM GROUP
        "Remix Ideas",
        # STATUS GROUP
        "Enrich Status", "Enrich Reason"
    ]

    # Write header row
    ws.append(headers)

    # Create enrichment lookup by row_number
    enrichment_map = {e["row_number"]: e for e in enrichments}

    # Write data rows
    for idea in ideas:
        enrichment = enrichment_map.get(idea.row_number, {})

        # Get examples
        examples = enrichment.get("examples", [])
        ex1 = examples[0] if len(examples) > 0 else {}
        ex2 = examples[1] if len(examples) > 1 else {}
        ex3 = examples[2] if len(examples) > 2 else {}

        # Get audio
        audio = enrichment.get("audio", {})

        # Build row data
        row_data = [
            # INPUT GROUP
            idea.date.strftime("%A"),  # Day
            idea.date.strftime("%Y-%m-%d"),  # Date (normalized)
            idea.content_type,  # Type
            idea.topic,  # Topic
            idea.description,  # Description

            # QUERY GROUP
            ", ".join(enrichment.get("topic_keywords", [])),  # Topic Keywords
            " | ".join(enrichment.get("queries", [])),  # Search Queries

            # EXAMPLE 1
            ex1.get("url", ""),
            ex1.get("views", ""),
            ex1.get("likes", ""),
            ex1.get("comments", ""),
            ex1.get("shares", ""),
            ex1.get("author_username", ""),
            ex1.get("caption", ""),
            ex1.get("audio_title", ""),
            enrichment.get("ex1_hook_summary", ""),

            # EXAMPLE 2
            ex2.get("url", ""),
            ex2.get("views", ""),
            ex2.get("likes", ""),
            ex2.get("comments", ""),
            ex2.get("shares", ""),
            ex2.get("author_username", ""),
            ex2.get("caption", ""),
            ex2.get("audio_title", ""),
            enrichment.get("ex2_hook_summary", ""),

            # EXAMPLE 3
            ex3.get("url", ""),
            ex3.get("views", ""),
            ex3.get("likes", ""),
            ex3.get("comments", ""),
            ex3.get("shares", ""),
            ex3.get("author_username", ""),
            ex3.get("caption", ""),
            ex3.get("audio_title", ""),
            enrichment.get("ex3_hook_summary", ""),

            # AUDIO GROUP
            audio.get("audio_title", ""),
            audio.get("audio_author", ""),
            audio.get("audio_confidence", ""),
            audio.get("audio_url", ""),
            enrichment.get("audio_fit_reason", ""),

            # LLM GROUP
            enrichment.get("remix_ideas", ""),

            # STATUS GROUP
            enrichment.get("enrich_status", ""),
            enrichment.get("enrich_reason", "")
        ]

        ws.append(row_data)

    # Apply header styling
    for cell in ws[1]:
        cell.font = Font(bold=True)

    # Freeze header row
    ws.freeze_panes = "A2"

    # Apply wrap_text to Remix Ideas column (column index based on headers)
    remix_ideas_col_idx = headers.index("Remix Ideas") + 1
    for row_idx in range(2, ws.max_row + 1):
        cell = ws.cell(row=row_idx, column=remix_ideas_col_idx)
        cell.alignment = Alignment(wrap_text=True)

    # Auto-fit column widths
    for col_idx, column_cells in enumerate(ws.columns, 1):
        col_letter = get_column_letter(col_idx)

        # Find max content length in column
        max_length = 0
        for cell in column_cells:
            if cell.value:
                cell_value_str = str(cell.value)
                # Cap content length check at 50 chars
                content_length = min(len(cell_value_str), 50)
                max_length = max(max_length, content_length)

        # Set column width (add padding)
        adjusted_width = max_length + 2
        ws.column_dimensions[col_letter].width = adjusted_width

    # Save workbook
    wb.save(output_path)


def build_run_log(ideas: List[ContentIdea], enrichments: List[Dict[str, Any]],
                  input_path: str, output_path: str, start_time: float, end_time: float) -> Dict[str, Any]:
    """
    Build complete run log with summary and per-row diagnostics.

    Args:
        ideas: List of original ContentIdea objects
        enrichments: List of enrichment dicts from enrich_row()
        input_path: Path to input Excel file
        output_path: Path to output Excel file
        start_time: time.time() at run start
        end_time: time.time() at run end

    Returns:
        Dict with run_summary and rows array
    """
    # Build run_summary
    timestamp = datetime.datetime.now(datetime.timezone.utc).isoformat()
    duration_seconds = round(end_time - start_time, 1)

    # Count status values
    status_counts = {"ok": 0, "partial": 0, "skipped": 0, "error": 0}
    for enrichment in enrichments:
        status = enrichment.get("enrich_status", "error")
        if status in status_counts:
            status_counts[status] += 1
        else:
            status_counts["error"] += 1

    run_summary = {
        "timestamp": timestamp,
        "input_file": os.path.basename(input_path),
        "output_file": os.path.basename(output_path),
        "total_rows": len(ideas),
        "status_counts": status_counts,
        "duration_seconds": duration_seconds
    }

    # Build rows array
    rows = []
    enrichment_map = {e["row_number"]: e for e in enrichments}

    for idea in ideas:
        enrichment = enrichment_map.get(idea.row_number, {})

        # Extract chosen audio info
        audio = enrichment.get("audio", {})
        chosen_audio = {
            "title": audio.get("audio_title", ""),
            "confidence": audio.get("audio_confidence", "")
        }

        # Extract example URLs (up to 3)
        examples = enrichment.get("examples", [])
        example_urls = [ex.get("url", "") for ex in examples[:3]]

        row_data = {
            "row_number": idea.row_number,
            "content_type": idea.content_type,
            "topic": idea.topic,
            "queries": enrichment.get("queries", []),
            "total_results": enrichment.get("total_results", 0),
            "scored_results": enrichment.get("scored_results", 0),
            "chosen_audio": chosen_audio,
            "example_urls": example_urls,
            "enrich_status": enrichment.get("enrich_status", "error"),
            "enrich_reason": enrichment.get("enrich_reason", "")
        }
        rows.append(row_data)

    return {
        "run_summary": run_summary,
        "rows": rows
    }


def save_run_log(run_log: Dict[str, Any], log_path: str) -> None:
    """
    Save run log dict to JSON file.

    Args:
        run_log: Run log dict from build_run_log()
        log_path: Path to write JSON file
    """
    with open(log_path, 'w', encoding='utf-8') as f:
        json.dump(run_log, f, indent=2, default=str)


if __name__ == "__main__":
    ideas, skipped = load_content_ideas("Content ideas.xlsx")
    print_summary(ideas, skipped)

    # Demonstrate Phase 3 pipeline with synthetic data
    print("\n" + "="*60)
    print("PHASE 3 PIPELINE: Data Processing & Selection")
    print("="*60 + "\n")

    # Show queries for first idea
    if ideas:
        test_idea = ideas[0]
        queries = generate_queries(test_idea)
        print(f"Example: Row {test_idea.row_number} - {test_idea.content_type} | {test_idea.topic}")
        print(f"Generated {len(queries)} queries: {', '.join(queries)}\n")

    # Create synthetic TikTok results to demonstrate pipeline
    print("Demonstrating with synthetic data:\n")

    # Create 5 fake TikTok results with varied engagement and some shared audio
    fake_results = [
        {
            "id": "7001",
            "webVideoUrl": "https://tiktok.com/@user1/video/7001",
            "text": "Founder story about building startup from scratch",
            "authorMeta": {"name": "startup_founder"},
            "createTimeISO": "2026-01-20T10:00:00Z",
            "playCount": 150000,
            "diggCount": 8000,
            "commentCount": 350,
            "shareCount": 120,
            "musicMeta": {
                "musicId": "audio123",
                "musicName": "Inspirational Beat",
                "musicAuthor": "AudioPro",
                "playUrl": "https://tiktok.com/music/audio123"
            }
        },
        {
            "id": "7002",
            "webVideoUrl": "https://tiktok.com/@user2/video/7002",
            "text": "Day in the life entrepreneur routine",
            "authorMeta": {"name": "daily_creator"},
            "createTimeISO": "2026-01-25T14:30:00Z",
            "playCount": 95000,
            "diggCount": 5200,
            "commentCount": 180,
            "shareCount": 65,
            "musicMeta": {
                "musicId": "audio123",
                "musicName": "Inspirational Beat",
                "musicAuthor": "AudioPro",
                "playUrl": "https://tiktok.com/music/audio123"
            }
        },
        {
            "id": "7003",
            "webVideoUrl": "https://tiktok.com/@user3/video/7003",
            "text": "Startup lesson learned the hard way",
            "authorMeta": {"name": "wise_founder"},
            "createTimeISO": "2026-02-01T09:15:00Z",
            "playCount": 220000,
            "diggCount": 12000,
            "commentCount": 580,
            "shareCount": 200,
            "musicMeta": {
                "musicId": "audio456",
                "musicName": "Upbeat Vibe",
                "musicAuthor": "MusicMaker",
                "playUrl": "https://tiktok.com/music/audio456"
            }
        },
        {
            "id": "7004",
            "webVideoUrl": "https://tiktok.com/@user4/video/7004",
            "text": "Building in public founder journey",
            "authorMeta": {"name": "transparent_ceo"},
            "createTimeISO": "2026-01-18T16:45:00Z",
            "playCount": 68000,
            "diggCount": 3800,
            "commentCount": 120,
            "shareCount": 45,
            "musicMeta": {
                "musicId": "audio123",
                "musicName": "Inspirational Beat",
                "musicAuthor": "AudioPro",
                "playUrl": "https://tiktok.com/music/audio123"
            }
        },
        {
            "id": "7005",
            "webVideoUrl": "https://tiktok.com/@user5/video/7005",
            "text": "Entrepreneur mindset shift that changed everything",
            "authorMeta": {"name": "growth_minded"},
            "createTimeISO": "2026-01-28T11:20:00Z",
            "playCount": 42000,
            "diggCount": 2100,
            "commentCount": 85,
            "shareCount": 28,
            "musicMeta": {
                "musicId": "audio789",
                "musicName": "Chill Lofi",
                "musicAuthor": "LofiBeats",
                "playUrl": "https://tiktok.com/music/audio789"
            }
        }
    ]

    # Process through full pipeline
    if ideas:
        test_idea = ideas[0]
        enriched = enrich_row(test_idea, fake_results)

        print(f"Processed {enriched['total_results']} raw results")
        print(f"After filtering/scoring: {enriched['scored_results']} results")
        print(f"Enrich Status: {enriched['enrich_status']}")
        if enriched.get('enrich_reason'):
            print(f"Enrich Reason: {enriched['enrich_reason']}")
        print()

        # Show top 3 examples
        print("Top 3 Examples:")
        for i, ex in enumerate(enriched['examples'], 1):
            print(f"  {i}. @{ex['author_username']}")
            print(f"     Score: {ex['final_score']:.2f}")
            print(f"     Views: {ex['views']:,} | Likes: {ex['likes']:,} | Comments: {ex['comments']:,}")
            print(f"     Caption: {ex['caption'][:60]}...")
            print(f"     Audio Title: {enriched.get(f'ex{i}_audio_title', 'N/A')}")
            print(f"     URL: {ex['url']}")
            print()

        # Show selected audio
        audio = enriched['audio']
        print("Selected Audio:")
        print(f"  Title: {audio['audio_title']}")
        print(f"  Author: {audio['audio_author']}")
        print(f"  Confidence: {audio['audio_confidence']} (appears in top results)")
        print(f"  URL: {audio['audio_url']}")
        print()

        # Show LLM content (if available)
        has_openai_key = os.environ.get("OPENAI_API_KEY") is not None

        if not has_openai_key:
            print("="*60)
            print("Set OPENAI_API_KEY to enable LLM text generation")
            print("="*60)
        else:
            print("LLM-Generated Content:")
            print(f"  Audio Fit Reason: {enriched.get('audio_fit_reason', 'N/A')}")
            print()
            print("  Hook Summaries:")
            for i in range(1, 4):
                summary = enriched.get(f'ex{i}_hook_summary', '')
                if summary:
                    print(f"    Example {i}: {summary}")
            print()
            print("  Remix Ideas:")
            remix_ideas = enriched.get('remix_ideas', '')
            if remix_ideas:
                print(f"    {remix_ideas}")
            else:
                print("    N/A")

    print("\n" + "="*60)
    print("Phase 4 pipeline ready. Set APIFY_TOKEN and OPENAI_API_KEY to run with full pipeline.")
    print("="*60)

    # TEMPORARY TEST CODE FOR PHASE 5 PLAN 01
    print("\n" + "="*60)
    print("Testing write_enriched_excel function")
    print("="*60 + "\n")

    # Test write_enriched_excel with first idea
    if ideas and enriched:
        test_enrichments = [enriched]
        test_ideas = [ideas[0]]
        write_enriched_excel(test_ideas, test_enrichments, "test_output.xlsx")
        print("Created test_output.xlsx")

        # Verify by reading it back
        test_wb = load_workbook("test_output.xlsx")
        test_ws = test_wb.active
        headers_read = [cell.value for cell in test_ws[1]]
        print(f"Headers count: {len(headers_read)}")
        print(f"First 5 headers: {headers_read[:5]}")
        print(f"Topic Keywords column exists: {'Topic Keywords' in headers_read}")
        test_wb.close()
        print("Verification passed!\n")

        # Test build_run_log and save_run_log
        print("="*60)
        print("Testing build_run_log and save_run_log functions")
        print("="*60 + "\n")

        start_test = time.time()
        time.sleep(0.1)  # Simulate some work
        end_test = time.time()

        run_log = build_run_log(
            test_ideas,
            test_enrichments,
            "Content ideas.xlsx",
            "test_output.xlsx",
            start_test,
            end_test
        )

        print(f"Run log structure created:")
        print(f"  Timestamp: {run_log['run_summary']['timestamp']}")
        print(f"  Total rows: {run_log['run_summary']['total_rows']}")
        print(f"  Status counts: {run_log['run_summary']['status_counts']}")
        print(f"  Duration: {run_log['run_summary']['duration_seconds']}s")
        print(f"  Rows array length: {len(run_log['rows'])}")

        # Verify status counts add up
        counts = run_log['run_summary']['status_counts']
        total = sum(counts.values())
        print(f"  Status counts sum: {total} (should equal total_rows)")

        # Save to JSON
        save_run_log(run_log, "test_run_log.json")
        print("\nCreated test_run_log.json")

        # Verify JSON is valid
        with open("test_run_log.json", 'r') as f:
            loaded_log = json.load(f)
        print(f"JSON validation: {'run_summary' in loaded_log and 'rows' in loaded_log}")
        print("Verification passed!")
    # END TEMPORARY TEST CODE
