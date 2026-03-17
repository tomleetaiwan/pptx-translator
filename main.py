import streamlit as st
import hashlib
import json
import os
import re
import sys
from importlib.metadata import PackageNotFoundError, version
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
# This app keeps its own cache key in st.session_state rather than using
# Streamlit's built-in st.cache_* decorators.
# Use a neutral version label because this covers any output-affecting change,
# including translation logic and font-writing behavior.
# Bump this when output behavior changes so the same upload is reprocessed
# instead of reusing a cached result from the current session.
TRANSLATION_CACHE_VERSION = "output-v1"

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

# Set up the Microsoft Entra ID token provider for authentication
token_provider = get_bearer_token_provider(
    DefaultAzureCredential(exclude_interactive_browser_credential=False),
    "https://cognitiveservices.azure.com/.default",
)

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


def find_protected_spans(text, protect_numbers=True):
    candidate_spans = []

    for label, pattern in PROTECTED_TEXT_PATTERNS:
        if label == "num" and not protect_numbers:
            continue
        for match in pattern.finditer(text):
            candidate_spans.append((match.start(), match.end(), label))

    for match in PROTECTED_TERM_PATTERN.finditer(text):
        candidate_spans.append((match.start(), match.end(), "term"))

    candidate_spans.sort(key=lambda span: (span[0], -(span[1] - span[0])))

    selected_spans = []
    current_end = -1
    for start, end, label in candidate_spans:
        if start < current_end:
            continue
        selected_spans.append((start, end, label))
        current_end = end

    return selected_spans


def mask_protected_content(text, protect_numbers=True):
    # Replace terms like AI, Microsoft, dates, and reference codes with placeholders
    # so the model can translate surrounding language without damaging those tokens.
    protected_spans = find_protected_spans(text, protect_numbers=protect_numbers)
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
):
    # Single-item translation is used both directly and as a safe fallback when
    # batch output is malformed or leaks placeholders back into the result.
    masked_text, replacements = mask_protected_content(original_text, protect_numbers=protect_numbers)

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

    restored_text = restore_protected_content(chinese_text, replacements)
    if contains_placeholder_tokens(restored_text):
        return original_text

    return restored_text


def translate_to_chinese(original_text, model, max_completion_tokens=512, translation_guidance=""):
    text_profile = classify_text(original_text)
    if text_profile in {"empty", "metric"}:
        return original_text

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
        )

    if text_profile == "stat":
        return translate_text_with_prompt(
            original_text,
            model,
            build_system_prompt(STAT_CAPTION_SYSTEM_PROMPT, translation_guidance),
            STAT_CAPTION_USER_PROMPT_TEMPLATE,
            max_completion_tokens=max_completion_tokens,
            protect_numbers=False,
        )

    return translate_text_with_prompt(
        original_text,
        model,
        build_system_prompt(TRANSLATION_SYSTEM_PROMPT, translation_guidance),
        TRANSLATION_USER_PROMPT_TEMPLATE,
        max_completion_tokens=max_completion_tokens,
        protect_numbers=True,
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


def translate_batch(texts, model, max_completion_tokens=3072, translation_guidance=""):
    if not texts:
        return [], False

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
        grouped_texts[text_profile].append((index, text))

    for text_profile, items in grouped_texts.items():
        if not items:
            continue

        # Batch only similar text profiles together so one prompt does not need to
        # simultaneously handle KPIs, citation rows, and normal body copy.
        if text_profile == "general":
            translated_group, group_used_fallback = translate_text_group(
                items,
                model,
                build_system_prompt(BATCH_TRANSLATION_SYSTEM_PROMPT, translation_guidance),
                BATCH_TRANSLATION_USER_PROMPT_TEMPLATE,
                max_completion_tokens=max_completion_tokens,
                protect_numbers=True,
                translation_guidance=translation_guidance,
            )
        elif text_profile == "stat":
            translated_group, group_used_fallback = translate_text_group(
                items,
                model,
                build_system_prompt(STAT_BATCH_TRANSLATION_SYSTEM_PROMPT, translation_guidance),
                STAT_CAPTION_BATCH_USER_PROMPT_TEMPLATE,
                max_completion_tokens=max_completion_tokens,
                protect_numbers=False,
                translation_guidance=translation_guidance,
            )
        else:
            translated_group, group_used_fallback = translate_text_group(
                items,
                model,
                build_system_prompt(SOURCE_NOTE_BATCH_TRANSLATION_SYSTEM_PROMPT, translation_guidance),
                SOURCE_NOTE_BATCH_USER_PROMPT_TEMPLATE,
                max_completion_tokens=max_completion_tokens,
                protect_numbers=False,
                translation_guidance=translation_guidance,
            )

        used_fallback = used_fallback or group_used_fallback
        for item_index, translated_text in translated_group:
            translation_results[item_index] = translated_text

    return translation_results, used_fallback


def translate_text_group(
    items,
    model,
    system_prompt,
    user_prompt_template,
    max_completion_tokens=3072,
    protect_numbers=True,
    translation_guidance="",
):
    # Build one masked JSON payload for the model, then restore protected tokens
    # item-by-item after the batch response comes back.
    masked_texts = []
    replacement_maps = []
    original_texts = []
    item_indices = []

    for item_index, text in items:
        masked_text, replacements = mask_protected_content(text, protect_numbers=protect_numbers)
        item_indices.append(item_index)
        original_texts.append(text)
        masked_texts.append(masked_text)
        replacement_maps.append(replacements)

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
            (item_index, translate_to_chinese(original_text, model, translation_guidance=translation_guidance))
            for item_index, original_text in items
        ], True

    try:
        translated_texts = parse_batch_translation_response(content, len(items))
        restored_results = []

        for item_index, original_text, translated_text, replacements in zip(
            item_indices,
            original_texts,
            translated_texts,
            replacement_maps,
        ):
            restored_text = restore_protected_content(translated_text, replacements)
            if contains_placeholder_tokens(restored_text):
                restored_results.append(
                    (item_index, translate_to_chinese(original_text, model, translation_guidance=translation_guidance))
                )
            else:
                restored_results.append((item_index, restored_text))

        return restored_results, False
    except (json.JSONDecodeError, ValueError):
        return [
            (item_index, translate_to_chinese(original_text, model, translation_guidance=translation_guidance))
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


def build_translation_cache_key(uploaded_bytes, translation_guidance):
    cache_digest = hashlib.sha256()
    cache_digest.update(uploaded_bytes)
    cache_digest.update(b"\0")
    cache_digest.update(TRANSLATION_CACHE_VERSION.encode("utf-8"))
    cache_digest.update(b"\0")
    cache_digest.update(normalize_translation_guidance(translation_guidance).encode("utf-8"))
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
    presentation = Presentation(BytesIO(uploaded_bytes))
    _, debug_info = collect_translation_targets(presentation)
    return debug_info


def translate_presentation(uploaded_bytes, model, progress_bar, status_placeholder, translation_guidance=""):
    presentation = Presentation(BytesIO(uploaded_bytes))
    text_targets, debug_info = collect_translation_targets(presentation)

    total_text_items = len(text_targets)
    if total_text_items == 0:
        progress_bar.progress(1.0)
        status_placeholder.info("No translatable text found. The original presentation is ready to download.")
        buffer = BytesIO()
        presentation.save(buffer)
        return buffer.getvalue(), debug_info

    status_placeholder.info(f"Found {total_text_items} text items. Starting translation...")
    translated_items = 0

    for batch in build_translation_batches(text_targets):
        batch_texts = [target.text for target in batch]
        translated_texts, used_fallback = translate_batch(
            batch_texts,
            model,
            translation_guidance=translation_guidance,
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


def show_download_button(file_bytes, file_name):
    st.download_button(         
        label="Download",
        data=file_bytes,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    ) 


def show_debug_panel(debug_info, used_cached_result):
    with st.expander("Debug details", expanded=True):
        st.caption("Use this panel to confirm how many shapes and text targets were scanned before translation.")

        metric_columns = st.columns(4)
        metric_columns[0].metric("Slides", debug_info["slide_count"])
        metric_columns[1].metric("Shapes", debug_info["shape_count"])
        metric_columns[2].metric("Text Items", debug_info["text_target_count"])
        metric_columns[3].metric("Batches", debug_info["batch_count"])

        metric_columns = st.columns(3)
        metric_columns[0].metric("Container Shapes", debug_info["container_shape_count"])
        metric_columns[1].metric("Fallback Batches", debug_info["fallback_batch_count"])
        metric_columns[2].metric("Used Cache", "Yes" if used_cached_result else "No")

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
uploaded_file = st.file_uploader("Upload a PowerPoint file", type=["pptx"])

if uploaded_file is None:
    reset_translation_state()
else:
    uploaded_bytes = uploaded_file.getvalue()
    # Include translation guidance in the hash so changing the prompt context
    # reprocesses the same upload instead of returning a stale cached result.
    uploaded_file_key = build_translation_cache_key(uploaded_bytes, translation_guidance)
    progress_placeholder = st.empty()
    status_placeholder = st.empty()
    used_cached_result = (
        st.session_state.translated_file_key == uploaded_file_key
        and st.session_state.translated_file_bytes is not None
    )

    if not used_cached_result:
        progress_bar = progress_placeholder.progress(0.0)
        try:
            translated_file_bytes, debug_info = translate_presentation(
                uploaded_bytes,
                model,
                progress_bar,
                status_placeholder,
                translation_guidance=translation_guidance,
            )
        except Exception as exc:
            progress_placeholder.empty()
            status_placeholder.error(f"Translation failed: {exc}")
            reset_translation_state()
            st.stop()

        st.session_state.translated_file_key = uploaded_file_key
        st.session_state.translated_file_bytes = translated_file_bytes
        st.session_state.translated_debug_info = debug_info
    elif st.session_state.translated_debug_info is None:
        st.session_state.translated_debug_info = analyze_presentation(uploaded_bytes)

    st.session_state.translated_file_name = build_output_filename(uploaded_file.name)
    progress_placeholder.progress(1.0)
    status_placeholder.success("Translation complete.")

    if debug_mode and st.session_state.translated_debug_info:
        show_debug_panel(st.session_state.translated_debug_info, used_cached_result)

    if st.session_state.translated_file_bytes:
        show_download_button(
            st.session_state.translated_file_bytes,
            st.session_state.translated_file_name,
        )


    

