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
from openpyxl import load_workbook
from dateutil.parser import parse as date_parse
import requests


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
        except (requests.RequestException, ValueError) as e:
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


if __name__ == "__main__":
    ideas, skipped = load_content_ideas("Content ideas.xlsx")
    print_summary(ideas, skipped)

    # Demonstrate Phase 2 pipeline: query generation
    print("\n" + "="*60)
    print("PHASE 2 PIPELINE: Query Generation")
    print("="*60 + "\n")

    total_queries = 0

    for idea in ideas:
        enriched = enrich_row_queries(idea)
        row_num = enriched["row_number"]
        queries = enriched["queries"]
        query_count = enriched["query_count"]

        print(f"Row {row_num}: {idea.content_type} | {idea.topic}")
        print(f"  Generated {query_count} queries:")
        for i, query in enumerate(queries, 1):
            print(f"    {i}. \"{query}\"")
        print()

        total_queries += query_count

    print("="*60)
    print(f"Total queries generated: {total_queries}")
    print("Average queries per row: {:.1f}".format(total_queries / len(ideas) if ideas else 0))
    print("\nNote: Set APIFY_TOKEN environment variable to fetch TikTok results")
    print("="*60)
