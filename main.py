import streamlit as st
import hashlib
import json
import os
import re
import sys
from collections import Counter
from importlib.metadata import PackageNotFoundError, version
from zipfile import BadZipFile, ZipFile
from azure.identity import DefaultAzureCredential, get_bearer_token_provider
from openai import APIConnectionError, APITimeoutError, InternalServerError, OpenAI, RateLimitError
from packaging.version import Version
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from io import BytesIO
from dotenv import load_dotenv
from tenacity import retry, retry_if_exception_type, stop_after_attempt, wait_random_exponential

MIN_OPENAI_VERSION = "1.109.1"
# These batch limits are tuned for the current model's ability to handle multiple translations at once while still following instructions and not hitting token-per-minute limits. Adjust as needed when using a different model or if the model
MAX_BATCH_SIZE = 20
# Keep each request small enough to reduce latency and lower the chance of model output drift.
MAX_BATCH_CHARACTERS = 2000
# Translation here is mostly constrained rewriting, so keep reasoning light to
# reduce latency and make batch behavior more predictable.
DEFAULT_REASONING_EFFORT = "minimal"
# A font with good Traditional Chinese support is needed to prevent missing glyphs after translation.
TRANSLATED_CHINESE_FONT_NAME = "微軟正黑體"
AZURE_OPENAI_SCOPE = "https://cognitiveservices.azure.com/.default"
# This app keeps its own cache key in st.session_state rather than using
# Streamlit's built-in st.cache_* decorators.
# Use a neutral version label because this covers any output-affecting change,
# including translation logic and font-writing behavior.
# Bump this when output behavior changes so the same upload is reprocessed
# instead of reusing a cached result from the current session.
TRANSLATION_CACHE_VERSION = "output-v3"
GLOSSARY_ENTRY_SEPARATOR_PATTERN = re.compile(r"\s*(?:=>|->|→)\s*")
TERM_CANDIDATE_MIN_OCCURRENCES = 2
TERM_CANDIDATE_MAX_RESULTS = 40
TERM_CONNECTOR_WORDS = {
    "a",
    "an",
    "and",
    "as",
    "at",
    "by",
    "for",
    "from",
    "in",
    "of",
    "on",
    "or",
    "the",
    "to",
    "via",
    "with",
}
TERM_STOPWORDS = TERM_CONNECTOR_WORDS | {
    "are",
    "be",
    "been",
    "being",
    "can",
    "current",
    "is",
    "it",
    "its",
    "new",
    "our",
    "their",
    "these",
    "this",
    "those",
    "use",
    "used",
    "using",
    "will",
    "your",
}
TERM_HEADWORDS = {
    "agent",
    "agents",
    "api",
    "apis",
    "assistant",
    "assistants",
    "automation",
    "capability",
    "capabilities",
    "cloud",
    "context",
    "copilot",
    "dashboard",
    "data",
    "database",
    "engine",
    "feature",
    "features",
    "foundation",
    "framework",
    "generation",
    "index",
    "insight",
    "insights",
    "integration",
    "interface",
    "language",
    "model",
    "models",
    "platform",
    "process",
    "processes",
    "product",
    "products",
    "protocol",
    "report",
    "reports",
    "sdk",
    "service",
    "services",
    "solution",
    "solutions",
    "stack",
    "storage",
    "studio",
    "system",
    "systems",
    "tool",
    "tools",
    "workflow",
    "workflows",
}
TERM_SUFFIXES = (
    "tion",
    "sion",
    "ment",
    "ness",
    "ity",
    "ism",
    "ics",
    "ware",
    "graphy",
    "ology",
)
CAPITALIZED_TERM_CANDIDATE_PATTERN = re.compile(
    r"\b(?:[A-Z][A-Za-z0-9+]*(?:-[A-Za-z0-9+]+)*|[A-Z]{2,})(?:\s+(?:(?:of|for|and|to|in|on|by|with|the|a|an|or|via|from)\s+)?(?:[A-Z][A-Za-z0-9+]*(?:-[A-Za-z0-9+]+)*|[A-Z]{2,})){0,3}\b"
)
LOWERCASE_TERM_CANDIDATE_PATTERN = re.compile(
    r"\b[a-z]+(?:-[a-z0-9]+)*(?:\s+[a-z]+(?:-[a-z0-9]+)*){1,3}\b"
)
UPPERCASE_TERM_TOKEN_PATTERN = re.compile(r"\b[A-Z]{2,}(?:[A-Z0-9+-]*[A-Z0-9])?\b")

# Different slide text profiles need different translation constraints.
# Statistics captions read better when kept fragment-like, while source-note rows
# need more conservative wording to preserve citations and report metadata.
TRANSLATION_SYSTEM_PROMPT = (
    "Translate the user's text into Traditional Chinese used in Taiwan. "
    "Tokens wrapped in double square brackets, such as [[TERM_0]] or [[NUM_0]], are protected placeholders and must remain unchanged exactly. "
    "Translate everything else naturally. "
    "Return only the translated text."
)
BATCH_TRANSLATION_SYSTEM_PROMPT = (
    "Translate each string in the provided JSON array into Traditional Chinese used in Taiwan. "
    "Return only a valid JSON array of translated strings in the same order as the input. "
    "Tokens wrapped in double square brackets, such as [[TERM_0]] or [[NUM_0]], are protected placeholders and must remain unchanged exactly. "
    "Translate everything else naturally."
)
STAT_CAPTION_SYSTEM_PROMPT = (
    "Translate infographic or statistics caption text into concise Traditional Chinese used in Taiwan. "
    "The input may be a sentence fragment placed next to a separate KPI number on the slide. "
    "Keep it as a natural caption fragment and do not invent a missing subject or duplicate a numeric value that is not in the input. "
    "Tokens wrapped in double square brackets, such as [[TERM_0]] or [[REF_0]], are protected placeholders and must remain unchanged exactly. "
    "Return only the translated text."
)
STAT_BATCH_TRANSLATION_SYSTEM_PROMPT = (
    "Translate each string in the provided JSON array into concise Traditional Chinese used in Taiwan for infographic or statistics captions. "
    "Each input may be a sentence fragment placed next to a separate KPI number on the slide. "
    "Keep each item as a natural caption fragment and do not invent a missing subject or duplicate a numeric value that is not in the input. "
    "Tokens wrapped in double square brackets, such as [[TERM_0]] or [[REF_0]], are protected placeholders and must remain unchanged exactly. "
    "Return only a valid JSON array of translated strings in the same order as the input."
)
SOURCE_NOTE_SYSTEM_PROMPT = (
    "Translate slide source-note and citation text into concise Traditional Chinese used in Taiwan. "
    "Preserve citation numbering, report codes, URLs, footnote markers, and organization or product names. "
    "You may keep pipe separators as separators. "
    "Translate connective phrases and descriptive words naturally, and render dates in a natural Traditional Chinese order when appropriate. "
    "If a segment looks like an official report title, keep the title compact and readable. "
    "Tokens wrapped in double square brackets are protected placeholders and must remain unchanged exactly. "
    "Return only the translated text."
)
SOURCE_NOTE_BATCH_TRANSLATION_SYSTEM_PROMPT = (
    "Translate each string in the provided JSON array into concise Traditional Chinese used in Taiwan for slide source notes and citations. "
    "Preserve citation numbering, report codes, URLs, footnote markers, and organization or product names. "
    "You may keep pipe separators as separators. "
    "Translate connective phrases and descriptive words naturally, and render dates in a natural Traditional Chinese order when appropriate. "
    "If a segment looks like an official report title, keep the title compact and readable. "
    "Tokens wrapped in double square brackets are protected placeholders and must remain unchanged exactly. "
    "Return only a valid JSON array of translated strings in the same order as the input."
)
TRANSLATION_USER_PROMPT_TEMPLATE = (
    "Translate the following text into Traditional Chinese used in Taiwan. "
    "Keep any placeholder tokens such as [[TERM_0]] or [[NUM_0]] unchanged. "
    "Return only the translated text.\n\n"
    "Text:\n{text}"
)
BATCH_TRANSLATION_USER_PROMPT_TEMPLATE = (
    "Translate each string in the following JSON array into Traditional Chinese used in Taiwan. "
    "Keep any placeholder tokens such as [[TERM_0]] or [[NUM_0]] unchanged. "
    "Return only a valid JSON array of translated strings in the same order as the input.\n\n"
    "JSON:\n{payload}"
)
STAT_CAPTION_USER_PROMPT_TEMPLATE = (
    "Translate the following infographic or statistics caption text into concise Traditional Chinese used in Taiwan. "
    "Keep any placeholder tokens such as [[TERM_0]] or [[REF_0]] unchanged. "
    "If the text is a sentence fragment, keep it as a natural caption fragment. "
    "Return only the translated text.\n\n"
    "Text:\n{text}"
)
STAT_CAPTION_BATCH_USER_PROMPT_TEMPLATE = (
    "Translate each string in the following JSON array into concise Traditional Chinese used in Taiwan for infographic or statistics captions. "
    "Keep any placeholder tokens such as [[TERM_0]] or [[REF_0]] unchanged. "
    "If an item is a sentence fragment, keep it as a natural caption fragment. "
    "Return only a valid JSON array of translated strings in the same order as the input.\n\n"
    "JSON:\n{payload}"
)
SOURCE_NOTE_USER_PROMPT_TEMPLATE = (
    "Translate the following slide source note or citation text into concise Traditional Chinese used in Taiwan. "
    "Keep citation numbering, report codes, organization names, product names, and placeholder tokens unchanged when appropriate. "
    "Return only the translated text.\n\n"
    "Text:\n{text}"
)
SOURCE_NOTE_BATCH_USER_PROMPT_TEMPLATE = (
    "Translate each string in the following JSON array into concise Traditional Chinese used in Taiwan for slide source notes and citations. "
    "Keep citation numbering, report codes, organization names, product names, and placeholder tokens unchanged when appropriate. "
    "Return only a valid JSON array of translated strings in the same order as the input.\n\n"
    "JSON:\n{payload}"
)
PLACEHOLDER_TOKEN_PATTERN = re.compile(r"\[\[[A-Z]+_\d+\]\]")
PROTECTED_TERM_PATTERN = re.compile(r"\b(?:Azure OpenAI|Microsoft|Azure|OpenAI|AI|IDC)\b")
STANDALONE_METRIC_PATTERN = re.compile(r"^\s*(?:\d+(?:\.\d+)?(?:[BMKbmk])?%?|\d+\.\s*|[¹²³⁴⁵⁶⁷⁸⁹⁰]+)\s*$")
SOURCE_TEXT_HINT_PATTERN = re.compile(r"\||\bsponsored by\b|#(?:[A-Za-z]{2,}\d{3,})|\b(?:snapshot|index|source)\b", re.IGNORECASE)
STAT_CAPTION_HINT_PATTERN = re.compile(r"\b(?:surveyed|leaders|report|reported|projected|automating|workflow|workflows|process|processes|agents?)\b", re.IGNORECASE)
CHINESE_CHARACTER_PATTERN = re.compile(r"[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]")
PROTECTED_TEXT_PATTERNS = [
    ("url", re.compile(r"https?://\S+")),
    ("email", re.compile(r"\b[\w.+-]+@[\w-]+(?:\.[\w-]+)+\b")),
    ("path", re.compile(r"\b[A-Za-z]:[\\/][^\s]+")),
    ("ref", re.compile(r"[¹²³⁴⁵⁶⁷⁸⁹⁰]+|#[A-Za-z]{2,}\d{3,}")),
    ("num", re.compile(r"(?:[$€£¥])?\d[\d,.:/-]*(?:[BMKbmk])?%?")),
]


class InvalidPowerPointFileError(Exception):
    pass

# Set up the Microsoft Entra ID token provider for authentication
entra_credential = DefaultAzureCredential(exclude_interactive_browser_credential=False)
token_provider = get_bearer_token_provider(entra_credential, AZURE_OPENAI_SCOPE)

# Set up the Streamlit app and OpenAI configuration
st.set_page_config(page_title="PowerPoint Translator")

# Keep the latest translation result in the current Streamlit session so reruns
# for the same user do not retranslate the same upload.
if "translated_file_key" not in st.session_state:
    st.session_state.translated_file_key = None

if "translated_file_bytes" not in st.session_state:
    st.session_state.translated_file_bytes = None

if "translated_file_name" not in st.session_state:
    st.session_state.translated_file_name = None

if "translated_debug_info" not in st.session_state:
    st.session_state.translated_debug_info = None

if "term_candidate_file_key" not in st.session_state:
    st.session_state.term_candidate_file_key = None

if "term_candidate_entries" not in st.session_state:
    st.session_state.term_candidate_entries = None

if "selected_term_candidates" not in st.session_state:
    st.session_state.selected_term_candidates = []

if "glossary_builder_message" not in st.session_state:
    st.session_state.glossary_builder_message = None

if "pending_glossary_lines" not in st.session_state:
    st.session_state.pending_glossary_lines = None

if "pending_selected_term_candidates_reset" not in st.session_state:
    st.session_state.pending_selected_term_candidates_reset = False

if "entra_auth_ready" not in st.session_state:
    st.session_state.entra_auth_ready = False

if "entra_auth_status" not in st.session_state:
    st.session_state.entra_auth_status = "not_checked"

if "entra_auth_message" not in st.session_state:
    st.session_state.entra_auth_message = None

# Load environment variables from a .env file
load_dotenv()

# Set up the OpenAI API configuration
def ensure_supported_runtime():
    try:
        openai_version = version("openai")
    except PackageNotFoundError:
        st.error("The openai package is not installed in the current Python environment.")
        st.stop()

    if Version(openai_version) < Version(MIN_OPENAI_VERSION):
        st.error(
            "Unsupported OpenAI SDK version detected. "
            f"Found {openai_version} in {sys.executable}. "
            f"Please upgrade to openai>={MIN_OPENAI_VERSION} and start the app with the same Python interpreter."
        )
        st.code(f'"{sys.executable}" -m pip install --upgrade "openai>={MIN_OPENAI_VERSION},<2"')
        st.code(f'"{sys.executable}" -m streamlit run main.py')
        st.stop()


def get_required_env(name):
    value = os.getenv(name)
    if value:
        return value

    st.error(f"Missing required environment variable: {name}")
    st.stop()


def build_base_url():
    openai_base_url = os.getenv("OPENAI_BASE_URL")
    if openai_base_url:
        return openai_base_url

    openai_endpoint = os.getenv("OPENAI_ENDPOINT")
    if openai_endpoint:
        return f"{openai_endpoint.rstrip('/')}/openai/v1/"

    return None


def summarize_auth_error(exc, max_length=360):
    normalized_message = " ".join(str(exc).split())
    if not normalized_message:
        return "Unable to acquire a Microsoft Entra ID token."

    if len(normalized_message) <= max_length:
        return normalized_message

    return f"{normalized_message[:max_length - 3]}..."


def verify_entra_authentication():
    try:
        entra_credential.get_token(AZURE_OPENAI_SCOPE)
    except Exception as exc:
        st.session_state.entra_auth_ready = False
        st.session_state.entra_auth_status = "failed"
        st.session_state.entra_auth_message = (
            "Microsoft Entra ID sign-in failed or was cancelled. "
            f"{summarize_auth_error(exc)}"
        )
        return False

    st.session_state.entra_auth_ready = True
    st.session_state.entra_auth_status = "success"
    st.session_state.entra_auth_message = (
        "Microsoft Entra ID sign-in verified. You can translate now. "
        "Azure OpenAI permissions are still checked when the translation request is sent."
    )
    return True


def show_entra_auth_status_panel():
    auth_status = st.session_state.entra_auth_status
    auth_ready = st.session_state.entra_auth_ready
    auth_message = st.session_state.entra_auth_message

    if auth_status == "success":
        state_value = "Ready"
        browser_value = "Not expected"
        next_step = "Click Translate PowerPoint to start translation."
        message_type = "success"
        message_text = auth_message or "Microsoft Entra ID sign-in verified."
    elif auth_status == "failed":
        state_value = "Failed"
        browser_value = "Retry may prompt"
        next_step = "Click Check Microsoft Entra sign-in again."
        message_type = "error"
        message_text = auth_message or "Microsoft Entra ID sign-in failed."
    else:
        state_value = "Not checked"
        browser_value = "Possible"
        next_step = "Click Check Microsoft Entra sign-in before translating."
        message_type = "info"
        message_text = (
            "The app has not checked Microsoft Entra ID sign-in yet. "
            "If no cached sign-in is available, the check can open your default browser."
        )

    with st.container(border=True):
        st.markdown("**Microsoft Entra ID Sign-inStatus**")

        metric_columns = st.columns(3)
        metric_columns[0].metric("Sign-in State", state_value)
        metric_columns[1].metric("Browser Sign-In", browser_value)
        metric_columns[2].metric(
            "Translate Button",
            "Enabled" if auth_ready else "Disabled",
        )

        if message_type == "success":
            st.success(message_text)
        elif message_type == "error":
            st.error(message_text)
        else:
            st.info(message_text)

        st.caption(f"Next step: {next_step}")


ensure_supported_runtime()

openai_base_url = build_base_url()
if not openai_base_url:
    st.error("Missing required environment variable: OPENAI_BASE_URL or OPENAI_ENDPOINT")
    st.stop()

client = OpenAI(
    base_url=openai_base_url,
    api_key=token_provider,
)
model = get_required_env("OPENAI_MODEL")

# Define a retry strategy for the OpenAI API call to handle the error of token-per-minute limits
@retry(
    retry=retry_if_exception_type((APIConnectionError, APITimeoutError, InternalServerError, RateLimitError)),
    wait=wait_random_exponential(min=2, max=30),
    stop=stop_after_attempt(5),
)
def completion_with_backoff(**kwargs):
    return client.chat.completions.create(**kwargs)


# Normalize chat completion parsing so single-item and batch translation can share the same fallback logic.
def get_response_content(response):
    try:
        finish_reason = response.choices[0].finish_reason
        if finish_reason == "content_filter":
            print("Content filter triggered. keep the original text.")
            return None

        return response.choices[0].message.content
    except (AttributeError, IndexError, KeyError):
        print("Error in message chat completions.")
        print(response.model_dump_json(indent=2))
        return None


def select_non_overlapping_spans(candidate_spans):
    candidate_spans.sort(key=lambda span: (span[0], -(span[1] - span[0])))

    selected_spans = []
    current_end = -1
    for span in candidate_spans:
        start = span[0]
        end = span[1]
        if start < current_end:
            continue
        selected_spans.append(span)
        current_end = end

    return selected_spans


def find_protected_spans(text, protect_numbers=True, include_term_protection=True):
    candidate_spans = []

    for label, pattern in PROTECTED_TEXT_PATTERNS:
        if label == "num" and not protect_numbers:
            continue
        for match in pattern.finditer(text):
            candidate_spans.append((match.start(), match.end(), label))

    if include_term_protection:
        for match in PROTECTED_TERM_PATTERN.finditer(text):
            candidate_spans.append((match.start(), match.end(), "term"))

    return select_non_overlapping_spans(candidate_spans)


def mask_protected_content(text, protect_numbers=True, include_term_protection=True):
    # Replace terms like AI, Microsoft, dates, and reference codes with placeholders
    # so the model can translate surrounding language without damaging those tokens.
    protected_spans = find_protected_spans(
        text,
        protect_numbers=protect_numbers,
        include_term_protection=include_term_protection,
    )
    if not protected_spans:
        return text, {}

    masked_parts = []
    replacements = {}
    last_end = 0

    for placeholder_index, (start, end, label) in enumerate(protected_spans):
        placeholder = f"[[{label.upper()}_{placeholder_index}]]"
        masked_parts.append(text[last_end:start])
        masked_parts.append(placeholder)
        replacements[placeholder] = text[start:end]
        last_end = end

    masked_parts.append(text[last_end:])
    return "".join(masked_parts), replacements


def restore_protected_content(text, replacements):
    restored_text = text
    for placeholder, original_text in replacements.items():
        restored_text = restored_text.replace(placeholder, original_text)

    return restored_text


def contains_placeholder_tokens(text):
    return bool(PLACEHOLDER_TOKEN_PATTERN.search(text))


def normalize_term_candidate(term):
    return " ".join(term.split())


def canonicalize_term_candidate(term):
    return normalize_term_candidate(term).casefold()


def trim_term_candidate_tokens(tokens):
    start_index = 0
    end_index = len(tokens)

    while start_index < end_index and tokens[start_index].casefold() in TERM_CONNECTOR_WORDS:
        start_index += 1

    while end_index > start_index and tokens[end_index - 1].casefold() in TERM_CONNECTOR_WORDS:
        end_index -= 1

    return tokens[start_index:end_index]


def looks_like_technical_headword(token):
    lower_token = token.casefold()
    if lower_token in TERM_HEADWORDS:
        return True

    return any(lower_token.endswith(suffix) for suffix in TERM_SUFFIXES)


def prepare_term_candidate(candidate):
    normalized_candidate = normalize_term_candidate(candidate)
    if not normalized_candidate:
        return None

    tokens = trim_term_candidate_tokens(normalized_candidate.split())
    if not tokens:
        return None

    normalized_candidate = " ".join(tokens)
    lower_tokens = [token.casefold() for token in tokens]

    if len(tokens) == 1:
        token = tokens[0]
        lower_token = lower_tokens[0]
        if lower_token in TERM_STOPWORDS:
            return None
        if token.isupper():
            return normalized_candidate
        if any(character.isdigit() for character in token) or "-" in token:
            return normalized_candidate
        if token[:1].isupper() and len(token) >= 4:
            return normalized_candidate
        return None

    if all(lower_token in TERM_STOPWORDS for lower_token in lower_tokens):
        return None

    has_distinctive_token = any(
        token.isupper()
        or token[:1].isupper()
        or "-" in token
        or any(character.isdigit() for character in token)
        for token in tokens
    )
    if has_distinctive_token:
        return normalized_candidate

    if looks_like_technical_headword(tokens[-1]):
        return normalized_candidate

    return None


def extract_term_candidates_from_text(text):
    normalized_text = normalize_term_candidate(text)
    if not normalized_text:
        return []

    candidates = []
    seen_candidates = set()

    for pattern in (
        CAPITALIZED_TERM_CANDIDATE_PATTERN,
        LOWERCASE_TERM_CANDIDATE_PATTERN,
        UPPERCASE_TERM_TOKEN_PATTERN,
    ):
        for match in pattern.finditer(normalized_text):
            candidate = prepare_term_candidate(match.group(0))
            if not candidate:
                continue

            candidate_key = canonicalize_term_candidate(candidate)
            if candidate_key in seen_candidates:
                continue

            seen_candidates.add(candidate_key)
            candidates.append(candidate)

    return candidates


def load_powerpoint_presentation(uploaded_bytes):
    try:
        with ZipFile(BytesIO(uploaded_bytes)) as archive:
            archive_entries = set(archive.namelist())
    except BadZipFile as exc:
        raise InvalidPowerPointFileError(
            "The uploaded file is not a valid PowerPoint file. Please upload a presentation saved in .pptx format."
        ) from exc

    required_entries = {"[Content_Types].xml", "_rels/.rels", "ppt/presentation.xml"}
    if not required_entries.issubset(archive_entries):
        raise InvalidPowerPointFileError(
            "The uploaded file is not a valid .pptx PowerPoint presentation. Please export or save it as .pptx and try again."
        )

    try:
        return Presentation(BytesIO(uploaded_bytes))
    except (BadZipFile, KeyError, ValueError) as exc:
        raise InvalidPowerPointFileError(
            "The uploaded PowerPoint file could not be opened. Please re-save it as .pptx in PowerPoint and try again."
        ) from exc


def extract_terminology_candidates(uploaded_bytes):
    presentation = load_powerpoint_presentation(uploaded_bytes)
    text_targets, _ = collect_translation_targets(presentation)
    candidate_counts = Counter()
    display_variants = {}

    for target in text_targets:
        for candidate in extract_term_candidates_from_text(target.text):
            candidate_key = canonicalize_term_candidate(candidate)
            candidate_counts[candidate_key] += 1

            variant_counter = display_variants.get(candidate_key)
            if variant_counter is None:
                variant_counter = Counter()
                display_variants[candidate_key] = variant_counter
            variant_counter[candidate] += 1

    terminology_candidates = []
    for candidate_key, occurrence_count in candidate_counts.items():
        if occurrence_count < TERM_CANDIDATE_MIN_OCCURRENCES:
            continue

        display_term = display_variants[candidate_key].most_common(1)[0][0]
        terminology_candidates.append(
            {
                "term": display_term,
                "count": occurrence_count,
                "word_count": len(display_term.split()),
            }
        )

    terminology_candidates.sort(
        key=lambda item: (-item["count"], -item["word_count"], item["term"].casefold())
    )
    return terminology_candidates[:TERM_CANDIDATE_MAX_RESULTS]


def build_upload_content_key(uploaded_bytes):
    return hashlib.sha256(uploaded_bytes).hexdigest()


def build_glossary_lines_from_terms(selected_terms, glossary_entries):
    existing_term_keys = {
        canonicalize_term_candidate(entry["source"])
        for entry in glossary_entries
    }
    glossary_lines = []
    added_terms = []

    for term in selected_terms:
        normalized_term = normalize_term_candidate(term)
        term_key = canonicalize_term_candidate(normalized_term)
        if term_key in existing_term_keys:
            continue

        glossary_lines.append(f"{normalized_term} => {normalized_term}")
        added_terms.append(normalized_term)
        existing_term_keys.add(term_key)

    return glossary_lines, added_terms


def append_glossary_lines(glossary_text, glossary_lines):
    if not glossary_lines:
        return glossary_text

    stripped_glossary_text = glossary_text.rstrip()
    if not stripped_glossary_text:
        return "\n".join(glossary_lines)

    return f"{stripped_glossary_text}\n" + "\n".join(glossary_lines)


def apply_pending_glossary_updates():
    pending_glossary_lines = st.session_state.pending_glossary_lines
    if not pending_glossary_lines:
        return

    current_glossary_text = st.session_state.get("translation_glossary_input", "")
    st.session_state.translation_glossary_input = append_glossary_lines(
        current_glossary_text,
        pending_glossary_lines,
    )
    st.session_state.pending_glossary_lines = None


def prepare_selected_term_candidates(available_terms):
    if st.session_state.pending_selected_term_candidates_reset:
        st.session_state.selected_term_candidates = []
        st.session_state.pending_selected_term_candidates_reset = False
        return

    available_term_set = set(available_terms)
    current_selection = st.session_state.get("selected_term_candidates", [])
    filtered_selection = [term for term in current_selection if term in available_term_set]
    if filtered_selection != current_selection:
        st.session_state.selected_term_candidates = filtered_selection


def parse_translation_glossary(glossary_text):
    glossary_entries = []
    seen_sources = {}

    for line_number, raw_line in enumerate(glossary_text.splitlines(), start=1):
        stripped_line = raw_line.strip()
        if not stripped_line or stripped_line.startswith("#"):
            continue

        parts = GLOSSARY_ENTRY_SEPARATOR_PATTERN.split(stripped_line, maxsplit=1)
        if len(parts) != 2:
            raise ValueError(
                f"Glossary line {line_number} must use 'source => target' format."
            )

        source_text = parts[0].strip()
        target_text = parts[1].strip()
        if not source_text or not target_text:
            raise ValueError(
                f"Glossary line {line_number} must include both a source term and a target translation."
            )

        previous_line_number = seen_sources.get(source_text)
        if previous_line_number is not None:
            raise ValueError(
                f"Glossary term '{source_text}' is defined more than once (lines {previous_line_number} and {line_number})."
            )

        seen_sources[source_text] = line_number
        glossary_entries.append({"source": source_text, "target": target_text})

    return glossary_entries


def normalize_translation_glossary(glossary_entries):
    if not glossary_entries:
        return ""

    normalized_pairs = sorted(
        (entry["source"], entry["target"]) for entry in glossary_entries
    )
    return "\n".join(
        f"{source_text} => {target_text}"
        for source_text, target_text in normalized_pairs
    )


def find_glossary_spans(text, glossary_entries):
    candidate_spans = []

    for glossary_index, entry in enumerate(glossary_entries):
        source_text = entry["source"]
        search_start = 0
        while True:
            match_start = text.find(source_text, search_start)
            if match_start == -1:
                break

            match_end = match_start + len(source_text)
            candidate_spans.append((match_start, match_end, glossary_index))
            search_start = match_end

    return select_non_overlapping_spans(candidate_spans)


def mask_builtin_term_placeholders(text):
    term_spans = [
        (match.start(), match.end(), "term")
        for match in PROTECTED_TERM_PATTERN.finditer(text)
    ]
    term_spans = select_non_overlapping_spans(term_spans)
    if not term_spans:
        return text, {}

    masked_parts = []
    replacements = {}
    last_end = 0

    for placeholder_index, (start, end, label) in enumerate(term_spans):
        placeholder = f"[[{label.upper()}_{placeholder_index}]]"
        masked_parts.append(text[last_end:start])
        masked_parts.append(placeholder)
        replacements[placeholder] = text[start:end]
        last_end = end

    masked_parts.append(text[last_end:])
    return "".join(masked_parts), replacements


def mask_glossary_terms(text, glossary_entries):
    if not glossary_entries:
        return text, {}

    glossary_spans = find_glossary_spans(text, glossary_entries)
    if not glossary_spans:
        return text, {}

    masked_parts = []
    replacements = {}
    last_end = 0

    for placeholder_index, (start, end, glossary_index) in enumerate(glossary_spans):
        placeholder = f"[[GLOSSARY_{placeholder_index}]]"
        masked_parts.append(text[last_end:start])
        masked_parts.append(placeholder)
        replacements[placeholder] = glossary_entries[glossary_index]["target"]
        last_end = end

    masked_parts.append(text[last_end:])
    return "".join(masked_parts), replacements


def mask_text_for_translation(text, glossary_entries, protect_numbers=True):
    # Protect URLs and numeric strings first, then freeze user-defined glossary
    # terms, and finally protect built-in literal product names.
    masked_text, protected_replacements = mask_protected_content(
        text,
        protect_numbers=protect_numbers,
        include_term_protection=False,
    )
    masked_text, glossary_replacements = mask_glossary_terms(masked_text, glossary_entries)
    masked_text, term_replacements = mask_builtin_term_placeholders(masked_text)

    return masked_text, protected_replacements, glossary_replacements, term_replacements


def restore_translated_text(
    translated_text,
    protected_replacements,
    glossary_replacements,
    term_replacements,
):
    restored_text = restore_protected_content(translated_text, term_replacements)
    restored_text = restore_protected_content(restored_text, glossary_replacements)
    return restore_protected_content(restored_text, protected_replacements)


def is_standalone_metric_text(text):
    return bool(STANDALONE_METRIC_PATTERN.fullmatch(text.strip()))


def is_source_note_text(text):
    return bool(SOURCE_TEXT_HINT_PATTERN.search(text))


def is_stat_caption_text(text):
    normalized_text = text.strip()
    if not normalized_text:
        return False

    starts_with_lowercase = normalized_text[:1].islower()
    has_stat_hint = bool(STAT_CAPTION_HINT_PATTERN.search(normalized_text))
    return starts_with_lowercase or has_stat_hint


def classify_text(text):
    # Route text into a translation profile before batching so captions, source rows,
    # and standalone metrics are not forced through the same prompt.
    # These profiles are inferred from the text extracted by python-pptx; they are
    # not PowerPoint metadata or a built-in pptx classification.
    normalized_text = " ".join(text.split())
    if not normalized_text:
        return "empty"
    if is_standalone_metric_text(normalized_text):
        return "metric"
    if is_source_note_text(normalized_text):
        return "source"
    if is_stat_caption_text(normalized_text):
        return "stat"
    return "general"


def build_translation_memory_key(text, text_profile=None):
    if text_profile is None:
        text_profile = classify_text(text)

    # Reuse only exact source strings so repeated sentences can share the same
    # translation without collapsing formatting-sensitive variants together.
    return text_profile, text


def normalize_translation_guidance(translation_guidance):
    return translation_guidance.replace("\r\n", "\n").strip()


def build_system_prompt(base_system_prompt, translation_guidance):
    normalized_guidance = normalize_translation_guidance(translation_guidance)
    if not normalized_guidance:
        return base_system_prompt

    return (
        "Use the following slide context and translation notes from the user to resolve ambiguity and keep terminology consistent. "
        "Treat this as context only and do not quote or translate it unless it also appears in the source text.\n\n"
        f"{normalized_guidance}\n\n"
        f"{base_system_prompt}"
    )


def translate_text_with_prompt(
    original_text,
    model,
    system_prompt,
    user_prompt_template,
    max_completion_tokens=512,
    protect_numbers=True,
    glossary_entries=None,
):
    # Single-item translation is used both directly and as a safe fallback when
    # batch output is malformed or leaks placeholders back into the result.
    if glossary_entries is None:
        glossary_entries = []

    masked_text, protected_replacements, glossary_replacements, term_replacements = mask_text_for_translation(
        original_text,
        glossary_entries,
        protect_numbers=protect_numbers,
    )

    response = completion_with_backoff(
        model=model,
        messages=[
            {"role": "system", "content": system_prompt},
            {
                "role": "user",
                "content": user_prompt_template.format(text=masked_text),
            }
        ],
        reasoning_effort=DEFAULT_REASONING_EFFORT,
        max_completion_tokens=max_completion_tokens,
    )

    chinese_text = get_response_content(response)
    if not chinese_text:
        return original_text

    restored_text = restore_translated_text(
        chinese_text,
        protected_replacements,
        glossary_replacements,
        term_replacements,
    )
    if contains_placeholder_tokens(restored_text):
        return original_text

    return restored_text


def translate_to_chinese(original_text, model, max_completion_tokens=512, translation_guidance="", glossary_entries=None):
    text_profile = classify_text(original_text)
    if text_profile in {"empty", "metric"}:
        return original_text

    if glossary_entries is None:
        glossary_entries = []

    # "source" means citation/source-note style text; "stat" means short KPI or
    # infographic caption text that should stay fragment-like and concise.
    if text_profile == "source":
        return translate_text_with_prompt(
            original_text,
            model,
            build_system_prompt(SOURCE_NOTE_SYSTEM_PROMPT, translation_guidance),
            SOURCE_NOTE_USER_PROMPT_TEMPLATE,
            max_completion_tokens=max_completion_tokens,
            protect_numbers=False,
            glossary_entries=glossary_entries,
        )

    if text_profile == "stat":
        return translate_text_with_prompt(
            original_text,
            model,
            build_system_prompt(STAT_CAPTION_SYSTEM_PROMPT, translation_guidance),
            STAT_CAPTION_USER_PROMPT_TEMPLATE,
            max_completion_tokens=max_completion_tokens,
            protect_numbers=False,
            glossary_entries=glossary_entries,
        )

    return translate_text_with_prompt(
        original_text,
        model,
        build_system_prompt(TRANSLATION_SYSTEM_PROMPT, translation_guidance),
        TRANSLATION_USER_PROMPT_TEMPLATE,
        max_completion_tokens=max_completion_tokens,
        protect_numbers=True,
        glossary_entries=glossary_entries,
    )


def parse_batch_translation_response(content, expected_count):
    # The model is instructed to return JSON only, but strip code fences in case it still wraps the payload.
    cleaned_content = content.strip()
    if cleaned_content.startswith("```"):
        lines = cleaned_content.splitlines()
        if len(lines) >= 3 and lines[0].startswith("```") and lines[-1].startswith("```"):
            cleaned_content = "\n".join(lines[1:-1]).strip()

    payload = json.loads(cleaned_content)
    if isinstance(payload, dict):
        payload = payload.get("translations")

    if not isinstance(payload, list):
        raise ValueError("The model did not return a JSON array of translations.")

    if len(payload) != expected_count:
        raise ValueError("The number of translated items does not match the input batch.")

    if not all(isinstance(item, str) for item in payload):
        raise ValueError("The model returned a non-string translation item.")

    return payload


def translate_batch(
    texts,
    model,
    max_completion_tokens=3072,
    translation_guidance="",
    translation_memory=None,
    glossary_entries=None,
):
    if not texts:
        return [], False

    if translation_memory is None:
        translation_memory = {}
    if glossary_entries is None:
        glossary_entries = []

    translation_results = [None] * len(texts)
    used_fallback = False
    grouped_texts = {
        "general": [],
        "stat": [],
        "source": [],
    }

    for index, text in enumerate(texts):
        text_profile = classify_text(text)
        if text_profile in {"empty", "metric"}:
            translation_results[index] = text
            continue

        cache_key = build_translation_memory_key(text, text_profile=text_profile)
        cached_translation = translation_memory.get(cache_key)
        if cached_translation is not None:
            translation_results[index] = cached_translation
            continue

        grouped_texts[text_profile].append((index, text, cache_key))

    for text_profile, items in grouped_texts.items():
        if not items:
            continue

        unique_items = []
        cache_key_to_item_indices = {}
        unique_item_cache_keys = {}

        for item_index, text, cache_key in items:
            existing_item_indices = cache_key_to_item_indices.get(cache_key)
            if existing_item_indices is None:
                cache_key_to_item_indices[cache_key] = [item_index]
                unique_items.append((item_index, text))
                unique_item_cache_keys[item_index] = cache_key
                continue

            existing_item_indices.append(item_index)

        # Batch only similar text profiles together so one prompt does not need to
        # simultaneously handle KPIs, citation rows, and normal body copy.
        if text_profile == "general":
            translated_group, group_used_fallback = translate_text_group(
                unique_items,
                model,
                build_system_prompt(BATCH_TRANSLATION_SYSTEM_PROMPT, translation_guidance),
                BATCH_TRANSLATION_USER_PROMPT_TEMPLATE,
                max_completion_tokens=max_completion_tokens,
                protect_numbers=True,
                translation_guidance=translation_guidance,
                glossary_entries=glossary_entries,
            )
        elif text_profile == "stat":
            translated_group, group_used_fallback = translate_text_group(
                unique_items,
                model,
                build_system_prompt(STAT_BATCH_TRANSLATION_SYSTEM_PROMPT, translation_guidance),
                STAT_CAPTION_BATCH_USER_PROMPT_TEMPLATE,
                max_completion_tokens=max_completion_tokens,
                protect_numbers=False,
                translation_guidance=translation_guidance,
                glossary_entries=glossary_entries,
            )
        else:
            translated_group, group_used_fallback = translate_text_group(
                unique_items,
                model,
                build_system_prompt(SOURCE_NOTE_BATCH_TRANSLATION_SYSTEM_PROMPT, translation_guidance),
                SOURCE_NOTE_BATCH_USER_PROMPT_TEMPLATE,
                max_completion_tokens=max_completion_tokens,
                protect_numbers=False,
                translation_guidance=translation_guidance,
                glossary_entries=glossary_entries,
            )

        used_fallback = used_fallback or group_used_fallback
        for item_index, translated_text in translated_group:
            cache_key = unique_item_cache_keys[item_index]
            translation_memory[cache_key] = translated_text

            for resolved_index in cache_key_to_item_indices[cache_key]:
                translation_results[resolved_index] = translated_text

    return translation_results, used_fallback


def translate_text_group(
    items,
    model,
    system_prompt,
    user_prompt_template,
    max_completion_tokens=3072,
    protect_numbers=True,
    translation_guidance="",
    glossary_entries=None,
):
    # Build one masked JSON payload for the model, then restore protected tokens
    # item-by-item after the batch response comes back.
    if glossary_entries is None:
        glossary_entries = []

    masked_texts = []
    protected_replacement_maps = []
    glossary_replacement_maps = []
    term_replacement_maps = []
    original_texts = []
    item_indices = []

    for item_index, text in items:
        masked_text, protected_replacements, glossary_replacements, term_replacements = mask_text_for_translation(
            text,
            glossary_entries,
            protect_numbers=protect_numbers,
        )
        item_indices.append(item_index)
        original_texts.append(text)
        masked_texts.append(masked_text)
        protected_replacement_maps.append(protected_replacements)
        glossary_replacement_maps.append(glossary_replacements)
        term_replacement_maps.append(term_replacements)

    response = completion_with_backoff(
        model=model,
        messages=[
            {"role": "system", "content": system_prompt},
            {
                "role": "user",
                "content": user_prompt_template.format(
                    payload=json.dumps(masked_texts, ensure_ascii=False)
                ),
            },
        ],
        reasoning_effort=DEFAULT_REASONING_EFFORT,
        max_completion_tokens=max_completion_tokens,
    )

    content = get_response_content(response)
    if not content:
        return [
            (
                item_index,
                translate_to_chinese(
                    original_text,
                    model,
                    translation_guidance=translation_guidance,
                    glossary_entries=glossary_entries,
                ),
            )
            for item_index, original_text in items
        ], True

    try:
        translated_texts = parse_batch_translation_response(content, len(items))
        restored_results = []

        for item_index, original_text, translated_text, protected_replacements, glossary_replacements, term_replacements in zip(
            item_indices,
            original_texts,
            translated_texts,
            protected_replacement_maps,
            glossary_replacement_maps,
            term_replacement_maps,
        ):
            restored_text = restore_translated_text(
                translated_text,
                protected_replacements,
                glossary_replacements,
                term_replacements,
            )
            if contains_placeholder_tokens(restored_text):
                restored_results.append(
                    (
                        item_index,
                        translate_to_chinese(
                            original_text,
                            model,
                            translation_guidance=translation_guidance,
                            glossary_entries=glossary_entries,
                        ),
                    )
                )
            else:
                restored_results.append((item_index, restored_text))

        return restored_results, False
    except (json.JSONDecodeError, ValueError):
        return [
            (
                item_index,
                translate_to_chinese(
                    original_text,
                    model,
                    translation_guidance=translation_guidance,
                    glossary_entries=glossary_entries,
                ),
            )
            for item_index, original_text in items
        ], True


def iter_shapes(shape_collection):
    # Walk the full shape tree so grouped and nested grouped content is processed the same as top-level shapes.
    for shape in shape_collection:
        yield shape

        try:
            child_shapes = shape.shapes
        except AttributeError:
            continue

        yield from iter_shapes(child_shapes)


def has_child_shapes(shape):
    try:
        _ = shape.shapes
    except AttributeError:
        return False

    return True


def collect_text_targets(shape, text_targets):
    # PowerPoint text is usually stored in text runs, but tables need to be collected separately.
    if shape.has_text_frame:
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                if run.text:
                    text_targets.append(run)
        return

    if shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                if cell.text:
                    text_targets.append(cell)


def build_translation_batches(text_targets):
    # Split work by item count and total characters to keep each model call predictable.
    batch = []
    batch_characters = 0

    for target in text_targets:
        target_length = max(1, len(target.text))
        batch_is_full = len(batch) >= MAX_BATCH_SIZE
        batch_exceeds_character_limit = batch and batch_characters + target_length > MAX_BATCH_CHARACTERS

        if batch_is_full or batch_exceeds_character_limit:
            yield batch
            batch = []
            batch_characters = 0

        batch.append(target)
        batch_characters += target_length

    if batch:
        yield batch


def summarize_text(text, max_length=80):
    normalized_text = " ".join(text.split())
    if len(normalized_text) <= max_length:
        return normalized_text

    return f"{normalized_text[:max_length - 3]}..."


def build_translation_cache_key(uploaded_bytes, translation_guidance, glossary_entries):
    cache_digest = hashlib.sha256()
    cache_digest.update(uploaded_bytes)
    cache_digest.update(b"\0")
    cache_digest.update(TRANSLATION_CACHE_VERSION.encode("utf-8"))
    cache_digest.update(b"\0")
    cache_digest.update(normalize_translation_guidance(translation_guidance).encode("utf-8"))
    cache_digest.update(b"\0")
    cache_digest.update(normalize_translation_glossary(glossary_entries).encode("utf-8"))
    return cache_digest.hexdigest()


def contains_chinese_characters(text):
    return bool(CHINESE_CHARACTER_PATTERN.search(text))


def set_typeface_on_run_properties(run_properties, tag_name, typeface, *successors):
    font_element = run_properties.find(qn(tag_name))
    if font_element is None:
        font_element = OxmlElement(tag_name)
        run_properties.insert_element_before(font_element, *successors)

    font_element.set("typeface", typeface)


def apply_typeface_to_run_properties(run_properties):
    # PowerPoint can keep Chinese text on the theme font unless both the generic
    # latin slot and the East Asian slot are explicitly overridden.
    set_typeface_on_run_properties(
        run_properties,
        "a:latin",
        TRANSLATED_CHINESE_FONT_NAME,
        "a:ea",
        "a:cs",
        "a:sym",
        "a:hlinkClick",
        "a:hlinkMouseOver",
        "a:rtl",
        "a:extLst",
    )
    set_typeface_on_run_properties(
        run_properties,
        "a:ea",
        TRANSLATED_CHINESE_FONT_NAME,
        "a:cs",
        "a:sym",
        "a:hlinkClick",
        "a:hlinkMouseOver",
        "a:rtl",
        "a:extLst",
    )


def apply_translated_font_to_paragraph(paragraph):
    # Some slide content still resolves from paragraph defaults after run text is
    # replaced, so mirror the same typeface onto paragraph-level defaults too.
    paragraph.font.name = TRANSLATED_CHINESE_FONT_NAME
    apply_typeface_to_run_properties(paragraph.font._rPr)
    apply_typeface_to_run_properties(paragraph._p.get_or_add_endParaRPr())


def apply_translated_font_to_run(run):
    apply_translated_font_to_paragraph(run._parent)
    run.font.name = TRANSLATED_CHINESE_FONT_NAME
    apply_typeface_to_run_properties(run.font._rPr)


def apply_translated_font(target):
    if hasattr(target, "font"):
        apply_translated_font_to_run(target)
        return

    if hasattr(target, "text_frame"):
        for paragraph in target.text_frame.paragraphs:
            apply_translated_font_to_paragraph(paragraph)
            for run in paragraph.runs:
                if run.text:
                    apply_translated_font_to_run(run)


def collect_translation_targets(presentation):
    text_targets = []
    # The debug view is populated from the same scan used for translation so the UI
    # can explain what the app actually saw without re-parsing the file differently.
    debug_info = {
        "slide_count": len(presentation.slides),
        "shape_count": 0,
        "container_shape_count": 0,
        "text_target_count": 0,
        "unique_text_target_count": 0,
        "reused_text_target_count": 0,
        "passthrough_text_target_count": 0,
        "glossary_term_count": 0,
        "batch_count": 0,
        "fallback_batch_count": 0,
        "translated_item_count": 0,
        "slides": [],
        "batches": [],
    }

    # Collect every translatable target first so progress can be based on real work remaining.
    for slide_number, slide in enumerate(presentation.slides, start=1):
        slide_shape_count = 0
        slide_text_target_count = 0

        for shape in iter_shapes(slide.shapes):
            debug_info["shape_count"] += 1
            slide_shape_count += 1

            if has_child_shapes(shape):
                debug_info["container_shape_count"] += 1

            previous_target_count = len(text_targets)
            collect_text_targets(shape, text_targets)
            slide_text_target_count += len(text_targets) - previous_target_count

        debug_info["slides"].append(
            {
                "slide_number": slide_number,
                "shape_count": slide_shape_count,
                "text_target_count": slide_text_target_count,
            }
        )

    debug_info["text_target_count"] = len(text_targets)

    unique_translation_keys = set()
    passthrough_text_target_count = 0
    for target in text_targets:
        text_profile = classify_text(target.text)
        if text_profile in {"empty", "metric"}:
            passthrough_text_target_count += 1
            continue

        unique_translation_keys.add(
            build_translation_memory_key(target.text, text_profile=text_profile)
        )

    debug_info["unique_text_target_count"] = len(unique_translation_keys)
    debug_info["passthrough_text_target_count"] = passthrough_text_target_count
    debug_info["reused_text_target_count"] = max(
        0,
        len(text_targets) - passthrough_text_target_count - len(unique_translation_keys),
    )

    for batch_number, batch in enumerate(build_translation_batches(text_targets), start=1):
        batch_texts = [target.text for target in batch]
        debug_info["batches"].append(
            {
                "batch_number": batch_number,
                "item_count": len(batch_texts),
                "character_count": sum(len(text) for text in batch_texts),
                "preview": [summarize_text(text) for text in batch_texts[:3]],
            }
        )

    debug_info["batch_count"] = len(debug_info["batches"])
    return text_targets, debug_info


def analyze_presentation(uploaded_bytes):
    presentation = load_powerpoint_presentation(uploaded_bytes)
    _, debug_info = collect_translation_targets(presentation)
    return debug_info


def translate_presentation(
    uploaded_bytes,
    model,
    progress_bar,
    status_placeholder,
    translation_guidance="",
    glossary_entries=None,
):
    presentation = load_powerpoint_presentation(uploaded_bytes)
    text_targets, debug_info = collect_translation_targets(presentation)
    translation_memory = {}
    if glossary_entries is None:
        glossary_entries = []

    debug_info["glossary_term_count"] = len(glossary_entries)

    total_text_items = len(text_targets)
    if total_text_items == 0:
        progress_bar.progress(1.0)
        status_placeholder.info("No translatable text found. The original presentation is ready to download.")
        buffer = BytesIO()
        presentation.save(buffer)
        return buffer.getvalue(), debug_info

    unique_text_count = debug_info["unique_text_target_count"]
    reused_text_count = debug_info["reused_text_target_count"]
    if reused_text_count > 0 and glossary_entries:
        status_placeholder.info(
            f"Found {total_text_items} text items. Applying {len(glossary_entries)} glossary terms and reusing repeated strings so only {unique_text_count} unique text items need translation..."
        )
    elif glossary_entries:
        status_placeholder.info(
            f"Found {total_text_items} text items. Applying {len(glossary_entries)} glossary terms before translation..."
        )
    elif reused_text_count > 0:
        status_placeholder.info(
            f"Found {total_text_items} text items. Reusing repeated strings so only {unique_text_count} unique text items need translation..."
        )
    else:
        status_placeholder.info(f"Found {total_text_items} text items. Starting translation...")

    translated_items = 0

    for batch in build_translation_batches(text_targets):
        batch_texts = [target.text for target in batch]
        translated_texts, used_fallback = translate_batch(
            batch_texts,
            model,
            translation_guidance=translation_guidance,
            translation_memory=translation_memory,
            glossary_entries=glossary_entries,
        )
        if used_fallback:
            debug_info["fallback_batch_count"] += 1

        for target, translated_text in zip(batch, translated_texts):
            original_text = target.text
            if translated_text == original_text:
                continue

            target.text = translated_text
            if contains_chinese_characters(translated_text):
                apply_translated_font(target)

        translated_items += len(batch)
        debug_info["translated_item_count"] = translated_items
        progress_bar.progress(translated_items / total_text_items)
        status_placeholder.info(f"Translated {translated_items} of {total_text_items} text items...")

    buffer = BytesIO()
    presentation.save(buffer)
    return buffer.getvalue(), debug_info


def build_output_filename(original_name):
    filename_root, extension = os.path.splitext(original_name)
    return f"{filename_root}.translated{extension or '.pptx'}"


def reset_translation_state():
    st.session_state.translated_file_key = None
    st.session_state.translated_file_bytes = None
    st.session_state.translated_file_name = None
    st.session_state.translated_debug_info = None


def reset_term_candidate_state():
    st.session_state.term_candidate_file_key = None
    st.session_state.term_candidate_entries = None
    st.session_state.selected_term_candidates = []
    st.session_state.glossary_builder_message = None
    st.session_state.pending_glossary_lines = None
    st.session_state.pending_selected_term_candidates_reset = False


def show_download_button(file_bytes, file_name):
    st.markdown(
        """
        <style>
        div[data-testid="stDownloadButton"] > button {
            background-color: #15803d !important;
            border-color: #15803d !important;
            color: #ffffff !important;
        }

        div[data-testid="stDownloadButton"] > button:hover {
            background-color: #166534 !important;
            border-color: #166534 !important;
            color: #ffffff !important;
        }

        div[data-testid="stDownloadButton"] > button:active {
            background-color: #14532d !important;
            border-color: #14532d !important;
            color: #ffffff !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
    st.download_button(
        label="Download",
        data=file_bytes,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        type="primary",
        use_container_width=True,
    )


def show_debug_panel(debug_info, used_cached_result):
    with st.expander("Debug details", expanded=True):
        st.caption("Use this panel to confirm how many shapes and text targets were scanned before translation, including repeated strings that can be reused.")

        metric_columns = st.columns(4)
        metric_columns[0].metric("Slides", debug_info["slide_count"])
        metric_columns[1].metric("Shapes", debug_info["shape_count"])
        metric_columns[2].metric("Text Items", debug_info["text_target_count"])
        metric_columns[3].metric("Batches", debug_info["batch_count"])

        metric_columns = st.columns(5)
        metric_columns[0].metric("Container Shapes", debug_info["container_shape_count"])
        metric_columns[1].metric("Unique Texts", debug_info["unique_text_target_count"])
        metric_columns[2].metric("Reused Texts", debug_info["reused_text_target_count"])
        metric_columns[3].metric("Fallback Batches", debug_info["fallback_batch_count"])
        metric_columns[4].metric("Used Cache", "Yes" if used_cached_result else "No")

        metric_columns = st.columns(2)
        metric_columns[0].metric("Passthrough Texts", debug_info["passthrough_text_target_count"])
        metric_columns[1].metric("Glossary Terms", debug_info["glossary_term_count"])

        st.write("Slide scan summary")
        st.dataframe(debug_info["slides"], width="stretch")

        st.write("Batch plan")
        st.dataframe(debug_info["batches"], width="stretch")

# Upload a PowerPoint file
debug_mode = st.toggle(
    "Debug mode",
    value=False,
    help="Show scan counts and batch planning details for the current upload.",
)
translation_guidance = st.text_area(
    "Optional slide context and translation notes",
    placeholder=(
        "Example:\n"
        "Main content: This deck summarizes Taiwan enterprise AI adoption trends.\n"
        "Notes: Keep product names in English. Use a formal business tone. Preserve citations and report codes."
    ),
    help="This text is added to the system prompt for every translation request.",
)
apply_pending_glossary_updates()
translation_glossary = st.text_area(
    "Optional terminology glossary",
    placeholder=(
        "Example:\n"
        "Copilot Studio => Copilot Studio\n"
        "agentic workflow => 代理式工作流程\n"
        "retrieval-augmented generation => 檢索增強生成"
    ),
    help="One term per line using 'source => target'. Matching source terms are replaced with the target translation after the model returns.",
    key="translation_glossary_input",
)
uploaded_file = st.file_uploader("Upload a PowerPoint file", type=["pptx"])

if uploaded_file is None:
    reset_translation_state()
    reset_term_candidate_state()
else:
    try:
        glossary_entries = parse_translation_glossary(translation_glossary)
    except ValueError as exc:
        reset_translation_state()
        st.error(f"Glossary format error: {exc}")
        st.stop()

    uploaded_bytes = uploaded_file.getvalue()
    try:
        load_powerpoint_presentation(uploaded_bytes)
    except InvalidPowerPointFileError as exc:
        reset_translation_state()
        reset_term_candidate_state()
        st.error(str(exc))
        st.stop()

    upload_content_key = build_upload_content_key(uploaded_bytes)
    if (
        st.session_state.term_candidate_file_key != upload_content_key
        or st.session_state.term_candidate_entries is None
    ):
        st.session_state.term_candidate_file_key = upload_content_key
        st.session_state.term_candidate_entries = extract_terminology_candidates(uploaded_bytes)
        st.session_state.selected_term_candidates = []

    if st.session_state.glossary_builder_message:
        st.info(st.session_state.glossary_builder_message)
        st.session_state.glossary_builder_message = None

    existing_glossary_term_keys = {
        canonicalize_term_candidate(entry["source"])
        for entry in glossary_entries
    }
    available_term_candidates = [
        candidate
        for candidate in st.session_state.term_candidate_entries
        if canonicalize_term_candidate(candidate["term"]) not in existing_glossary_term_keys
    ]

    with st.expander("Suggested glossary candidates", expanded=bool(available_term_candidates)):
        st.caption(
            "The candidate list is heuristic. Selected terms are added as 'source => source' glossary lines so you can edit the right-hand side before translating."
        )

        if available_term_candidates:
            candidate_count_map = {
                candidate["term"]: candidate["count"]
                for candidate in available_term_candidates
            }
            candidate_options = [candidate["term"] for candidate in available_term_candidates]
            prepare_selected_term_candidates(candidate_options)
            selected_terms = st.multiselect(
                "Repeated English term candidates",
                options=candidate_options,
                format_func=lambda term: f"{term} ({candidate_count_map[term]} occurrences)",
                placeholder="Select terms to add to the glossary",
                key="selected_term_candidates",
            )
            st.caption(
                f"Showing {len(available_term_candidates)} repeated English term candidates found in this deck."
            )

            if st.button("Add selected terms to glossary"):
                if not selected_terms:
                    st.session_state.pending_glossary_lines = None
                    st.session_state.glossary_builder_message = "No terms were added because nothing was selected."
                else:
                    glossary_lines, added_terms = build_glossary_lines_from_terms(
                        selected_terms,
                        glossary_entries,
                    )
                    if added_terms:
                        st.session_state.pending_glossary_lines = glossary_lines
                        st.session_state.glossary_builder_message = (
                            f"Added {len(added_terms)} terms to the glossary draft."
                        )
                    else:
                        st.session_state.pending_glossary_lines = None
                        st.session_state.glossary_builder_message = (
                            "No new terms were added because the selected terms are already in the glossary."
                        )

                st.session_state.pending_selected_term_candidates_reset = True
                st.rerun()
        else:
            st.caption("No repeated English term candidates were found outside the current glossary.")

    # Include translation guidance in the hash so changing the prompt context
    # reprocesses the same upload instead of returning a stale cached result.
    uploaded_file_key = build_translation_cache_key(
        uploaded_bytes,
        translation_guidance,
        glossary_entries,
    )
    progress_placeholder = st.empty()
    status_placeholder = st.empty()
    used_cached_result = (
        st.session_state.translated_file_key == uploaded_file_key
        and st.session_state.translated_file_bytes is not None
    )
    translate_button_label = "Retranslate PowerPoint" if used_cached_result else "Translate PowerPoint"
    action_columns = st.columns(2)
    with action_columns[0]:
        auth_check_requested = st.button(
            "Check Microsoft Entra ID sign-in",
            use_container_width=True,
            help=(
                "Try to acquire a Microsoft Entra ID token now. If no cached sign-in is "
                "available, your default browser may open so you can sign in before translation starts."
            ),
        )

    if auth_check_requested:
        with st.spinner("Checking Microsoft Entra ID sign-in status..."):
            verify_entra_authentication()

    with action_columns[1]:
        translation_requested = st.button(
            translate_button_label,
            type="primary",
            use_container_width=True,
            disabled=not st.session_state.entra_auth_ready,
            help=(
                "Check Entra sign-in first so the app can prompt for Microsoft Entra authentication before translation."
                if not st.session_state.entra_auth_ready
                else None
            ),
        )

    show_entra_auth_status_panel()

    if translation_requested:
        progress_bar = progress_placeholder.progress(0.0)
        try:
            translated_file_bytes, debug_info = translate_presentation(
                uploaded_bytes,
                model,
                progress_bar,
                status_placeholder,
                translation_guidance=translation_guidance,
                glossary_entries=glossary_entries,
            )
        except Exception as exc:
            progress_placeholder.empty()
            status_placeholder.error(f"Translation failed: {exc}")
            reset_translation_state()
            st.stop()

        st.session_state.translated_file_key = uploaded_file_key
        st.session_state.translated_file_bytes = translated_file_bytes
        st.session_state.translated_debug_info = debug_info
        used_cached_result = False
    elif used_cached_result and st.session_state.translated_debug_info is None:
        st.session_state.translated_debug_info = analyze_presentation(uploaded_bytes)

    current_result_available = (
        st.session_state.translated_file_key == uploaded_file_key
        and st.session_state.translated_file_bytes is not None
    )
    if current_result_available and st.session_state.translated_debug_info is not None:
        st.session_state.translated_debug_info["glossary_term_count"] = len(glossary_entries)

    if current_result_available:
        st.session_state.translated_file_name = build_output_filename(uploaded_file.name)
        progress_placeholder.progress(1.0)
        status_placeholder.success("Translation complete.")
    else:
        progress_placeholder.empty()
        if st.session_state.entra_auth_ready:
            status_placeholder.info(
                "Review the suggested glossary candidates above, then click Translate PowerPoint."
            )
        else:
            status_placeholder.info(
                "Review the suggested glossary candidates above, click Check Microsoft Entra sign-in, then start translation."
            )

    if current_result_available and debug_mode and st.session_state.translated_debug_info:
        show_debug_panel(st.session_state.translated_debug_info, used_cached_result)

    if current_result_available and st.session_state.translated_file_bytes:
        show_download_button(
            st.session_state.translated_file_bytes,
            st.session_state.translated_file_name,
        )


    

