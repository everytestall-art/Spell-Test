import io
import re
import json
import csv
import urllib.request
from pathlib import Path
import base64
import html
import textwrap
from difflib import SequenceMatcher

import docx
import fitz  # PyMuPDF
import streamlit as st
import streamlit.components.v1 as components

from PIL import Image, ImageDraw
from pythainlp.tokenize import word_tokenize, sent_tokenize
from pythainlp.util import dict_trie
from pythainlp.corpus import thai_words
from pythainlp.spell import NorvigSpellChecker

from rapidfuzz import process, fuzz

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    OPENPYXL_AVAILABLE = True
except Exception:
    Workbook = None
    Font = None
    PatternFill = None
    Alignment = None
    get_column_letter = None
    OPENPYXL_AVAILABLE = False

try:
    from wordfreq import zipf_frequency, top_n_list
    WORDFREQ_AVAILABLE = True
except Exception:
    WORDFREQ_AVAILABLE = False
    zipf_frequency = None
    top_n_list = None


# =========================================================
# CONFIGURATION
# =========================================================

st.set_page_config(
    page_title="ระบบตรวจสอบคำถูก-คำผิด ไทย–อังกฤษ",
    layout="wide"
)


# =========================================================
# CUSTOM DICTIONARIES
# =========================================================

CUSTOM_WORDS = {
    "สรรพสามิต",
    "สรรพสามิตสโมสร",
    "กรมสรรพสามิต",
    "สนับสนุน",
    "สิทธิประโยชน์",
    "กองทุนสรรพสามิตสโมสร",
    "ของที่ระลึก",
}

CUSTOM_PHRASES = {
    "สรรพสามิตสโมสร",
    "กองทุนสรรพสามิตสโมสร",
    "กรมสรรพสามิต",
    "สิทธิประโยชน์",
    "ของที่ระลึก",
}

COMMON_MISTAKES = {
    "สรรพมิต": {
        "correct": "สรรพสามิต",
        "reason": "คำว่า 'สรรพสามิต' สะกดไม่ครบ",
        "confidence": 0.99,
    },

    "สรรพมิตสโมสร": {
        "correct": "สรรพสามิตสโมสร",
        "reason": "คำว่า 'สรรพสามิต' ในวลีสะกดไม่ครบ",
        "confidence": 0.99,
    },
}


# =========================================================
# GOOGLE SHEETS / JSON CORRECT-WORD DICTIONARY (WHITELIST)
# =========================================================

# ใส่ลิงก์ Export CSV ของ Google Sheets ที่นี่ (ต้องเปิดสิทธิ์ Anyone with the link)
# ตัวอย่าง: "https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID/export?format=csv"
GOOGLE_SHEETS_CSV_URL = "https://docs.google.com/spreadsheets/d/1KIR5OTpTEWfwQ2W6KziQUWUim1aA9__L/export?format=csv&gid=673336886"

# ไฟล์สำรองกรณีโหลดจากเน็ตไม่สำเร็จ
JSON_DICTIONARY_PATH = Path(
    r"C:\Users\ADMIN\OneDrive\Desktop\SpellCheck\dictionary.json"
)


def load_json_correct_words(json_path: Path, sheet_url: str = "") -> tuple:
    """
    โหลดคำถูกจาก Google Sheets ก่อน ถ้ายกเลิก/เชื่อมต่อไม่ได้ ค่อยโหลดจาก dictionary.json (Fallback)
    """
    raw_words = []

    # 1. พยายามโหลดจาก Google Sheets (ถ้ามีลิงก์)
    if sheet_url:
        try:
            req = urllib.request.Request(sheet_url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req, timeout=5) as response:
                decoded = response.read().decode('utf-8-sig')
                
                # อ่าน CSV โดยสนใจคอลัมน์แรกสุด
                reader = csv.reader(io.StringIO(decoded))
                for row in reader:
                    if row and str(row[0]).strip():
                        raw_words.append(str(row[0]).strip())
                        
            correct_words = {w for w in raw_words if w}
            return (
                correct_words,
                f"โหลดคลังคำจาก Google Sheets สำเร็จ ({len(correct_words):,} คำ)",
                True,
            )
        except Exception as e:
            print(f"Warning: Failed to load from Google Sheets ({e}). Falling back to local JSON.")
            # ถ้าโหลดไม่ได้ ก็ปล่อยให้ตกลงไปทำ Fallback ด้านล่าง

    # 2. Fallback: โหลดจากไฟล์ JSON ภายในเครื่อง
    try:
        raw_json = json_path.read_text(encoding="utf-8-sig")

        try:
            data = json.loads(raw_json)
        except json.JSONDecodeError:
            cleaned_json = re.sub(r",\s*([}\]])", r"\1", raw_json)
            data = json.loads(cleaned_json)

    except FileNotFoundError:
        return (set(), f"ไม่พบไฟล์คลังคำสำรอง: {json_path}", False)
    except json.JSONDecodeError as exc:
        return (set(), f"ไฟล์ dictionary.json มีรูปแบบไม่ถูกต้อง: {exc}", False)
    except Exception as exc:
        return (set(), f"ไม่สามารถอ่าน dictionary.json ได้: {exc}", False)

    if isinstance(data, list):
        raw_words = data
    elif isinstance(data, dict):
        for key in ("words", "correct_words", "dictionary", "correct", "whitelist"):
            value = data.get(key)
            if isinstance(value, list):
                raw_words = value
                break
        else:
            raw_words = list(data.keys())
    else:
        return (set(), "dictionary.json ต้องเป็น JSON Array หรือ JSON Object", False)

    correct_words = {
        str(word).strip()
        for word in raw_words
        if isinstance(word, (str, int, float)) and str(word).strip()
    }

    return (
        correct_words,
        f"โหลดคลังคำสำรองจาก JSON แล้ว {len(correct_words):,} คำ",
        True,
    )


JSON_CORRECT_WORDS, JSON_DICTIONARY_STATUS, JSON_DICTIONARY_OK = (
    load_json_correct_words(JSON_DICTIONARY_PATH, GOOGLE_SHEETS_CSV_URL)
)


def is_json_correct_word(word: str) -> bool:
    """True เมื่อคำถูกระบุไว้ใน dictionary.json แบบตรงตัว"""
    return bool(word and word.strip() in JSON_CORRECT_WORDS)


def is_json_correct_occurrence(word: str, context: str = "") -> bool:
    """
    ตรวจ whitelist แบบเข้าใจบริบท เพื่อไม่ให้ Engine ฟ้องบางส่วนของวลีที่ถูกต้อง
    เช่น JSON มี "กองทุนสรรพสามิตสงเคราะห์" แต่ Engine จับ "กองทุนสรรพสามิต"
    """
    if not word:
        return False

    normalized_word = re.sub(r"\s+", "", str(word).strip())
    normalized_context = re.sub(r"\s+", "", str(context or ""))

    if not normalized_word:
        return False

    for correct_word in JSON_CORRECT_WORDS:
        normalized_correct = re.sub(r"\s+", "", str(correct_word).strip())
        if not normalized_correct:
            continue
        if normalized_word == normalized_correct:
            return True
        if normalized_word in normalized_correct and normalized_correct in normalized_context:
            return True

    return False


# =========================================================
# THAI CHARACTER / TEXT UTILITIES
# =========================================================

THAI_PATTERN = re.compile(
    r"[\u0E00-\u0E7F]+"
)


def contains_thai(text: str) -> bool:
    return bool(
        text and THAI_PATTERN.search(text)
    )


def normalize_text(text: str) -> str:

    if not text:
        return ""

    text = text.replace(
        "\u200b",
        ""
    )

    text = text.replace(
        "\ufeff",
        ""
    )

    text = re.sub(
        r"\s+",
        " ",
        text
    )

    return text.strip()


def thai_only(text: str) -> str:

    return "".join(
        re.findall(
            r"[\u0E00-\u0E7F]",
            text or ""
        )
    )


# =========================================================
# EDIT DISTANCE
# =========================================================

def levenshtein_distance(
    a: str,
    b: str
) -> int:

    if a == b:
        return 0

    if not a:
        return len(b)

    if not b:
        return len(a)

    previous = list(
        range(
            len(b) + 1
        )
    )

    for i, char_a in enumerate(
        a,
        start=1
    ):

        current = [i]

        for j, char_b in enumerate(
            b,
            start=1
        ):

            insert_cost = (
                current[j - 1] + 1
            )

            delete_cost = (
                previous[j] + 1
            )

            replace_cost = (
                previous[j - 1]
                + (char_a != char_b)
            )

            current.append(
                min(
                    insert_cost,
                    delete_cost,
                    replace_cost
                )
            )

        previous = current

    return previous[-1]


def similarity_score(
    a: str,
    b: str
) -> float:

    if not a or not b:
        return 0.0

    return (
        SequenceMatcher(
            None,
            a,
            b
        ).ratio()
        * 100
    )


# =========================================================
# DICTIONARY / SPELL CHECKER
# =========================================================

@st.cache_resource
def load_dictionary_and_checker(json_correct_words: tuple):

    base_words = set(
        thai_words()
    )

    full_dict = (
        base_words
        .union(CUSTOM_WORDS)
        .union(CUSTOM_PHRASES)
        .union(json_correct_words)
    )

    custom_trie = dict_trie(
        full_dict
    )

    custom_checker = NorvigSpellChecker(
        custom_dict=full_dict
    )

    return (
        full_dict,
        custom_trie,
        custom_checker
    )


THAI_DICT, CUSTOM_TRIE, SPELL_CHECKER = (
    load_dictionary_and_checker(
        tuple(sorted(JSON_CORRECT_WORDS))
    )
)


# =========================================================
# FUZZY DICTIONARY
# =========================================================

FUZZY_DICTIONARY = sorted(
    set(
        CUSTOM_WORDS
        .union(CUSTOM_PHRASES)
        .union(
            {
                item["correct"]
                for item in COMMON_MISTAKES.values()
            }
        )
    )
)


# =========================================================
# SESSION STATE
# =========================================================

if "analyzed_data" not in st.session_state:
    st.session_state.analyzed_data = None

if "file_key" not in st.session_state:
    st.session_state.file_key = None


if "preview_refresh_token" not in st.session_state:
    st.session_state.preview_refresh_token = 0


def reset_preview_focus_state():
    """รีเซ็ต Focus Mode เมื่อเปลี่ยนขอบเขตการแสดงผล"""
    st.session_state.preview_refresh_token += 1


# =========================================================
# SENTENCE / CONTEXT
# =========================================================

def get_sentence_context(
    full_text: str,
    target_word: str
) -> str:
    if not full_text:
        return ""

    text = str(full_text)

    chunks = re.split(r"[\r\n]+", text)
    for chunk in chunks:
        chunk = chunk.strip()
        if target_word and target_word in chunk:
            parts = re.split(r"(?<=[.!?。！？])\s+|[;；]+", chunk)
            for part in parts:
                if target_word in part:
                    return part.strip()
            return chunk

    try:
        sentences = sent_tokenize(text)
        for sent in sentences:
            if target_word in sent:
                return sent.strip()
    except Exception:
        pass

    index = text.find(target_word) if target_word else -1
    if index >= 0:
        left = max(0, index - 80)
        right = min(len(text), index + len(target_word) + 80)
        return text[left:right].strip()

    return text[:180].strip()

def get_context_for_phrase(
    full_text: str,
    phrase: str
) -> str:

    return get_sentence_context(
        full_text,
        phrase
    )


# =========================================================
# HTML HIGHLIGHT
# =========================================================

def highlight_word_in_html(
    sentence: str,
    word: str
) -> str:

    if not sentence or not word:
        return html.escape(
            sentence or ""
        )

    escaped_sentence = html.escape(
        sentence
    )

    escaped_word = html.escape(
        word
    )

    pattern = re.escape(
        escaped_word
    )

    return re.sub(
        f"({pattern})",
        r'<mark style="'
        r'background-color:#ffcccc;'
        r'color:#b30000;'
        r'padding:2px 4px;'
        r'border-radius:4px;'
        r'font-weight:bold;'
        r'">\1</mark>',
        escaped_sentence
    )


# =========================================================
# PDF BOUNDING BOX ENGINE
# =========================================================

def normalize_for_bbox(
    text: str
) -> str:

    if not text:
        return ""

    return re.sub(
        r"\s+",
        "",
        text
    )


def find_word_bboxes(
    page,
    target_word: str
) -> list:

    """
    ค้นหาตำแหน่งคำผิดบน PDF

    รองรับ:
    1. คำผิดอยู่ใน PDF word เดียว
    2. คำผิดประกอบด้วยหลาย PDF words
    3. phrase เช่น สรรพมิตสโมสร
    """

    if not target_word:
        return []

    target = normalize_for_bbox(
        target_word
    )

    if not target:
        return []

    words = page.get_text(
        "words"
    )

    if not words:
        return []

    bboxes = []

    # -----------------------------------------------------
    # CASE 1
    # คำอยู่ใน word เดียว
    # -----------------------------------------------------

    for inst in words:

        if len(inst) < 5:
            continue

        extracted = str(
            inst[4]
        )

        normalized = normalize_for_bbox(
            extracted
        )

        if target in normalized:

            x0, y0, x1, y1 = (
                inst[:4]
            )

            bboxes.append(
                (
                    float(x0),
                    float(y0),
                    float(x1),
                    float(y1)
                )
            )

    if bboxes:
        return bboxes

    # -----------------------------------------------------
    # CASE 2
    # phrase อยู่หลาย words
    # -----------------------------------------------------

    for start_idx in range(
        len(words)
    ):

        combined = ""

        selected = []

        for end_idx in range(
            start_idx,
            min(
                len(words),
                start_idx + 8
            )
        ):

            word_text = str(
                words[end_idx][4]
            )

            combined += normalize_for_bbox(
                word_text
            )

            selected.append(
                words[end_idx]
            )

            if target in combined:

                x0 = min(
                    float(item[0])
                    for item in selected
                )

                y0 = min(
                    float(item[1])
                    for item in selected
                )

                x1 = max(
                    float(item[2])
                    for item in selected
                )

                y1 = max(
                    float(item[3])
                    for item in selected
                )

                return [
                    (
                        x0,
                        y0,
                        x1,
                        y1
                    )
                ]

            # ถ้ายาวเกิน target มาก
            if len(combined) > len(target) + 10:
                break

    return []


# =========================================================
# CREATE ERROR ID
# =========================================================

def create_error_id(
    page_num: int,
    index: int
) -> str:

    return (
        f"error-page-{page_num}-"
        f"item-{index}"
    )


# =========================================================
# PDF PREVIEW HTML
# =========================================================

def image_to_base64(
    image: Image.Image
) -> str:

    buffer = io.BytesIO()

    image.save(
        buffer,
        format="PNG"
    )

    return base64.b64encode(
        buffer.getvalue()
    ).decode(
        "utf-8"
    )


def build_interactive_pdf_preview(
    page,
    error_items: list,
    page_num: int
) -> str:

    """
    สร้าง PDF Preview ที่กรอบแดงสามารถกดได้

    Click
      ↓
    #error-page-X-item-Y
      ↓
    scroll ไปยังกล่องวิเคราะห์
    """

    pix = page.get_pixmap(
        dpi=150,
        alpha=False
    )

    img_bytes = pix.tobytes(
        "png"
    )

    image = Image.open(
        io.BytesIO(img_bytes)
    )

    image_width = image.width
    image_height = image.height

    page_rect = page.rect

    base64_image = base64.b64encode(
        img_bytes
    ).decode(
        "utf-8"
    )

    overlay_html = []

    for item in error_items:

        error_id = item.get(
            "error_id"
        )

        bboxes = item.get(
            "bboxes",
            []
        )

        if not bboxes:
            continue

        for bbox in bboxes:

            x0, y0, x1, y1 = bbox

            left = (
                x0 /
                page_rect.width
                * 100
            )

            top = (
                y0 /
                page_rect.height
                * 100
            )

            width = (
                (x1 - x0)
                /
                page_rect.width
                * 100
            )

            height = (
                (y1 - y0)
                /
                page_rect.height
                * 100
            )

            is_uncertain = item.get("status") == "uncertain"
            border_color = "#2563eb" if is_uncertain else "#ff2b2b"
            hover_border = "#1d4ed8" if is_uncertain else "#b00000"
            background = "rgba(37,99,235,0.10)" if is_uncertain else "rgba(255,0,0,0.10)"
            hover_background = "rgba(37,99,235,0.24)" if is_uncertain else "rgba(255,0,0,0.25)"

            tooltip_text = (
                f"{item['word']} · ยังไม่แน่ใจ"
                if is_uncertain
                else f"{item['word']} → {item.get('suggested', '')}"
            )
            tooltip = html.escape(tooltip_text)

            overlay_html.append(
                f"""
                <button
                    type="button"
                    class="preview-error-box"
                    data-error-id="{error_id}"
                    title="{tooltip}"
                    style="
                        position:absolute;
                        left:{left:.4f}%;
                        top:{top:.4f}%;
                        width:{max(width, 1):.4f}%;
                        height:{max(height, 1.5):.4f}%;

                        border:3px solid {border_color};
                        background:{background};

                        box-sizing:border-box;
                        border-radius:4px;

                        cursor:pointer;
                        z-index:10;
                        padding:0;

                        transition:all .15s ease;
                    "

                    onmouseover="
                        this.style.background='{hover_background}';
                        this.style.borderColor='{hover_border}';
                    "

                    onmouseout="
                        this.style.background='{background}';
                        this.style.borderColor='{border_color}';
                    "
                >
                </button>
                """
            )

    preview_script = """
    <script>
    (function () {
        function resizePreviewFrame() {
            try {
                const viewport = window.parent.innerHeight || 900;
                const targetHeight = Math.max(560, Math.min(980, viewport - 190));

                if (window.frameElement) {
                    window.frameElement.style.height = targetHeight + "px";
                }

                document.body.style.margin = "0";
                document.body.style.overflowY = "auto";

                // Sticky toolbar ของ Preview
                const stickyToolbar = document.getElementById("preview-sticky-toolbar");
                if (stickyToolbar) {
                    stickyToolbar.style.position = "sticky";
                    stickyToolbar.style.top = "0px";
                    stickyToolbar.style.zIndex = "100";
                }
            } catch (err) {
                // ใช้ความสูงจาก Streamlit เป็น fallback
            }
        }

        resizePreviewFrame();
        window.addEventListener("resize", resizePreviewFrame);

        // -------------------------------------------------
        // RESET STALE FOCUS MODE
        // หลัง Streamlit rerun บางครั้ง DOM เดิมยังมี display:none
        // จากการกดกรอบ Preview ก่อนหน้า
        // -------------------------------------------------
        try {
            const parentDoc = window.parent.document;

            parentDoc
                .querySelectorAll('[data-focus-hidden="1"]')
                .forEach(function (el) {
                    el.style.display = "";
                    delete el.dataset.focusHidden;
                });

            parentDoc
                .querySelectorAll("#analysis-focus-panel.active")
                .forEach(function (panel) {
                    panel.classList.remove("active");
                });
        } catch (err) {
            // fallback: ปล่อยให้ Streamlit rerender ตามปกติ
        }

        // -------------------------------------------------
        // ZOOM CONTROLS
        // ซูมได้ 50% - 250% ครั้งละ 25%
        // กรอบคำผิด/ไม่แน่ใจอยู่ใน stage เดียวกับรูป
        // จึงขยาย/ย่อตามเอกสารพร้อมกัน
        // -------------------------------------------------
        let zoomLevel = 100;
        const minZoom = 50;
        const maxZoom = 250;
        const zoomStep = 25;

        const stage = document.getElementById("preview-page-stage");
        const zoomValue = document.getElementById("preview-zoom-value");
        const zoomIn = document.getElementById("preview-zoom-in");
        const zoomOut = document.getElementById("preview-zoom-out");
        const zoomReset = document.getElementById("preview-zoom-reset");

        function applyZoom(nextZoom) {
            zoomLevel = Math.max(minZoom, Math.min(maxZoom, nextZoom));

            if (stage) {
                // ใช้ width แทน transform เพื่อให้ scroll area
                // ขยายตามขนาดจริงเมื่อซูมเกิน 100%
                stage.style.width = zoomLevel + "%";
            }

            if (zoomValue) {
                zoomValue.textContent = zoomLevel + "%";
            }

            if (zoomOut) {
                zoomOut.disabled = zoomLevel <= minZoom;
                zoomOut.style.opacity = zoomOut.disabled ? "0.4" : "1";
            }

            if (zoomIn) {
                zoomIn.disabled = zoomLevel >= maxZoom;
                zoomIn.style.opacity = zoomIn.disabled ? "0.4" : "1";
            }
        }

        if (zoomIn) {
            zoomIn.addEventListener("click", function () {
                applyZoom(zoomLevel + zoomStep);
            });
        }

        if (zoomOut) {
            zoomOut.addEventListener("click", function () {
                applyZoom(zoomLevel - zoomStep);
            });
        }

        if (zoomReset) {
            zoomReset.addEventListener("click", function () {
                applyZoom(100);
            });
        }

        applyZoom(100);

        const boxes = document.querySelectorAll(".preview-error-box");

        boxes.forEach(function (box) {
            box.addEventListener("click", function (event) {
                event.preventDefault();
                event.stopPropagation();

                const errorId = box.getAttribute("data-error-id");
                if (!errorId) return;

                try {
                    const parentDoc = window.parent.document;
                    const target = parentDoc.getElementById(errorId);

                    if (!target) return;

                    const resultColumn =
                        target.closest('[data-testid="column"]') ||
                        parentDoc;

                    const panel = resultColumn.querySelector("#analysis-focus-panel");
                    if (!panel) return;

                    const getData = function (name) {
                        return target.getAttribute("data-" + name) || "";
                    };

                    const setText = function (selector, value) {
                        const el = panel.querySelector(selector);
                        if (el) el.textContent = value;
                    };

                    setText(".focus-location", getData("location"));
                    setText(".focus-word", getData("word"));
                    setText(".focus-suggested", getData("suggested"));
                    setText(".focus-detail", getData("detail"));
                    setText(".focus-sentence", getData("sentence"));
                    setText(".focus-confidence", getData("confidence"));
                    setText(".focus-method", getData("method"));

                    panel.classList.add("active");

                    // ผูกปุ่มปิด Focus Mode จาก iframe Preview ไปยัง DOM หลัก
                    // (script ที่ฝังด้วย st.markdown จะไม่ถูกใช้)
                    const closeFocusMode = function () {
                        panel.classList.remove("active");

                        resultColumn
                            .querySelectorAll('[data-focus-hidden="1"]')
                            .forEach(function (el) {
                                el.style.display = "";
                                delete el.dataset.focusHidden;
                            });
                    };

                    const closeBtn = panel.querySelector("#analysis-focus-close");
                    const backBtn = panel.querySelector("#analysis-focus-back");

                    if (closeBtn && closeBtn.dataset.bound !== "1") {
                        closeBtn.dataset.bound = "1";
                        closeBtn.addEventListener("click", closeFocusMode);
                    }

                    if (backBtn && backBtn.dataset.bound !== "1") {
                        backBtn.dataset.bound = "1";
                        backBtn.addEventListener("click", closeFocusMode);
                    }

                    // ซ่อนรายการยาวทั้งหมดในฝั่งผลวิเคราะห์
                    // เหลือเฉพาะผลที่สัมพันธ์กับกรอบแดง 1:1
                    resultColumn
                        .querySelectorAll('div[data-testid="stExpander"]')
                        .forEach(function (el) {
                            el.dataset.focusHidden = "1";
                            el.style.display = "none";
                        });

                    resultColumn
                        .querySelectorAll('[data-baseweb="tab-list"]')
                        .forEach(function (el) {
                            el.dataset.focusHidden = "1";
                            el.style.display = "none";
                        });

                    // ไม่ใช้ scrollIntoView อีกแล้ว
                    // ผู้ใช้จึงไม่ถูกพาไปท้ายหน้า
                } catch (err) {
                    console.warn("ไม่สามารถเปิดผลวิเคราะห์แบบ 1:1 ได้", err);
                }
            });
        });
    })();
    </script>
    """

    preview_html = f"""
    <div
        style="
            width:100%;
            background:#f3f4f6;
            border:1px solid #d1d5db;
            border-radius:12px;
            padding:16px;
            box-sizing:border-box;
        "
    >

        <div
            id="preview-sticky-toolbar"
            style="
                position:sticky;
                top:0;
                z-index:100;
                margin:-16px -16px 12px -16px;
                padding:14px 16px 12px 16px;
                display:flex;
                justify-content:space-between;
                align-items:center;
                gap:12px;
                flex-wrap:wrap;
                background:rgba(243,244,246,0.98);
                border-bottom:1px solid #d8dee6;
                box-shadow:0 2px 6px rgba(15,23,42,.06);
                backdrop-filter:blur(6px);
                -webkit-backdrop-filter:blur(6px);
            "
        >
            <div>
                <strong
                    style="
                        font-size:16px;
                        color:#111827;
                    "
                >
                    ตัวอย่างเอกสาร หน้า {page_num}
                </strong>

                <div
                    style="
                        font-size:13px;
                        color:#6b7280;
                        margin-top:3px;
                    "
                >
                    กดกรอบสีแดงหรือสีน้ำเงินเพื่อดูรายการวิเคราะห์
                </div>
            </div>

            <div
                style="
                    display:flex;
                    align-items:center;
                    gap:10px;
                    flex-wrap:wrap;
                    justify-content:flex-end;
                "
            >
                <div
                    style="
                        font-size:13px;
                        color:#b91c1c;
                        font-weight:600;
                    "
                >
                    {len(error_items)} จุดที่พบ
                </div>

                <div
                    class="preview-zoom-controls"
                    style="
                        display:flex;
                        align-items:center;
                        gap:4px;
                        background:#ffffff;
                        border:1px solid #cbd5e1;
                        border-radius:7px;
                        padding:3px;
                        line-height:1;
                    "
                >
                    <button
                        type="button"
                        id="preview-zoom-out"
                        title="ซูมออก"
                        style="
                            width:30px;
                            height:30px;
                            border:0;
                            border-radius:5px;
                            background:#f8fafc;
                            color:#334155;
                            font-size:20px;
                            cursor:pointer;
                        "
                    >−</button>

                    <span
                        id="preview-zoom-value"
                        style="
                            min-width:52px;
                            text-align:center;
                            font-size:13px;
                            font-weight:700;
                            color:#334155;
                        "
                    >100%</span>

                    <button
                        type="button"
                        id="preview-zoom-in"
                        title="ซูมเข้า"
                        style="
                            width:30px;
                            height:30px;
                            border:0;
                            border-radius:5px;
                            background:#f8fafc;
                            color:#334155;
                            font-size:20px;
                            cursor:pointer;
                        "
                    >+</button>

                    <button
                        type="button"
                        id="preview-zoom-reset"
                        title="กลับเป็น 100%"
                        style="
                            min-width:42px;
                            height:30px;
                            border:0;
                            border-radius:5px;
                            background:#f8fafc;
                            color:#334155;
                            font-size:12px;
                            cursor:pointer;
                            padding:0 7px;
                        "
                    >100</button>
                </div>
            </div>
        </div>

        <div
            id="preview-scroll-container"
            style="
                width:100%;
                overflow:auto;
                border-radius:8px;
                background:#e5e7eb;
                box-shadow:0 2px 8px rgba(0,0,0,.08);
            "
        >
            <div
                id="preview-page-stage"
                style="
                    position:relative;
                    width:100%;
                    min-width:100%;
                    line-height:0;
                    background:white;
                    transform-origin:top left;
                "
            >
                <img
                    src="data:image/png;base64,{base64_image}"
                    style="
                        display:block;
                        width:100%;
                        height:auto;
                    "
                >

                {''.join(overlay_html)}
            </div>
        </div>

        {preview_script}

    </div>
    """

    return preview_html


# =========================================================
# DRAW RED BOXES
# =========================================================

def draw_red_boxes_on_image(
    image: Image.Image,
    page_fitz,
    error_items: list
) -> Image.Image:

    if image is None:
        return image

    img_copy = image.copy()

    draw = ImageDraw.Draw(
        img_copy
    )

    rect = page_fitz.rect

    scale_x = (
        img_copy.width /
        rect.width
    )

    scale_y = (
        img_copy.height /
        rect.height
    )

    for item in error_items:

        bboxes = item.get(
            "bboxes",
            []
        )

        for bbox in bboxes:

            x0, y0, x1, y1 = bbox

            box = [
                x0 * scale_x - 4,
                y0 * scale_y - 4,
                x1 * scale_x + 4,
                y1 * scale_y + 4
            ]

            outline_color = (
                "blue" if item.get("status") == "uncertain" else "red"
            )
            draw.rectangle(
                box,
                outline=outline_color,
                width=4
            )

    return img_copy


# =========================================================
# COMMON MISTAKE DETECTION
# =========================================================

def detect_known_mistakes(
    text: str
) -> list:

    results = []

    if not text:
        return results

    for wrong, info in (
        COMMON_MISTAKES.items()
    ):

        # dictionary.json เป็น whitelist: ถ้าผู้ใช้ระบุว่าคำนี้ถูก ให้ข้ามทันที
        if is_json_correct_word(wrong):
            continue

        start = 0

        while True:

            index = text.find(
                wrong,
                start
            )

            if index == -1:
                break

            results.append({
                "wrong": wrong,
                "correct": info["correct"],
                "reason": info["reason"],
                "confidence": info["confidence"],
                "index": index,
            })

            start = (
                index +
                max(
                    1,
                    len(wrong)
                )
            )

    return results


# =========================================================
# PHRASE DETECTION
# =========================================================

def generate_candidate_phrases(
    text: str
) -> list:

    candidates = set()

    if not text:
        return []

    for phrase in CUSTOM_PHRASES:

        if phrase in text:

            candidates.add(
                phrase
            )

    try:

        tokens = word_tokenize(
            text,
            engine="newmm",
            custom_dict=CUSTOM_TRIE
        )

        clean_tokens = [
            token.strip()
            for token in tokens
            if contains_thai(token)
        ]

        for n in range(
            2,
            6
        ):

            for i in range(
                len(clean_tokens) - n + 1
            ):

                candidate = "".join(
                    clean_tokens[
                        i:i+n
                    ]
                )

                if len(candidate) >= 4:

                    candidates.add(
                        candidate
                    )

    except Exception:
        pass

    return sorted(
        candidates,
        key=len,
        reverse=True
    )


# =========================================================
# FUZZY MATCHING
# =========================================================

def fuzzy_find_best_match(
    word: str,
    choices=None,
    score_cutoff=80
):

    if not word:
        return None

    if choices is None:

        choices = FUZZY_DICTIONARY

    if not choices:
        return None

    result = process.extractOne(
        word,
        choices,
        scorer=fuzz.ratio,
        score_cutoff=score_cutoff
    )

    if not result:
        return None

    matched_word = result[0]
    score = result[1]

    if matched_word == word:
        return None

    return {
        "word": matched_word,
        "score": float(score)
    }


# =========================================================
# FUZZY PHRASE DETECTION
# =========================================================

def detect_fuzzy_phrases(
    text: str
) -> list:

    results = []

    if not text:
        return results

    candidates = generate_candidate_phrases(
        text
    )

    for candidate in candidates:

        if candidate in CUSTOM_PHRASES:
            continue

        if candidate in THAI_DICT:
            continue

        if not contains_thai(
            candidate
        ):
            continue

        match = fuzzy_find_best_match(
            candidate,
            CUSTOM_PHRASES,
            score_cutoff=82
        )

        if not match:
            continue

        matched = match["word"]
        score = match["score"]

        if matched == candidate:
            continue

        distance = levenshtein_distance(
            candidate,
            matched
        )

        if score >= 92:

            confidence = 0.97

        elif score >= 88:

            confidence = 0.93

        elif score >= 84:

            confidence = 0.88

        else:

            confidence = 0.80

        results.append({
            "word": candidate,
            "suggested": matched,
            "score": score,
            "distance": distance,
            "confidence": confidence,
        })

    return results


# =========================================================
# CONTEXT PROBABILITY ENGINE
# =========================================================
# ใช้ "ความน่าจะเป็นตามบริบท" ร่วมกับความคล้ายของคำ
# เพื่อจับกรณีที่คำสะกดถูกตามพจนานุกรม แต่ใช้ผิดบริบท
# เช่น "ฉันไปโรงเรียร ตอนสวย" -> "ฉันไปโรงเรียน ตอนสาย"
#
# หมายเหตุ:
# - Engine นี้ไม่แทนที่ระบบเดิม แต่ทำงานเพิ่มหลังการตรวจแบบคำเดี่ยว
# - ถ้าความน่าจะเป็นสูงมาก จะจัดเป็นคำผิด
# - ถ้าคะแนนยังไม่เด็ดขาด จะจัดเป็น "ยังไม่แน่ใจ" (กรอบน้ำเงิน)
# - สามารถขยาย CONTEXT_NEXT_WORD_PRIORS เพิ่มได้ภายหลัง

CONTEXT_NEXT_WORD_PRIORS = {
    # คำบอกช่วงเวลา
    "ตอน": {
        "เช้า": 0.96,
        "สาย": 0.99,
        "บ่าย": 0.96,
        "เย็น": 0.96,
        "ค่ำ": 0.93,
        "กลางคืน": 0.90,
    },
    "ช่วง": {
        "เช้า": 0.96,
        "สาย": 0.93,
        "บ่าย": 0.96,
        "เย็น": 0.96,
        "ค่ำ": 0.90,
    },
}


def _context_tokens(text: str) -> list:
    try:
        return [
            token.strip()
            for token in word_tokenize(
                text,
                engine="newmm",
                custom_dict=CUSTOM_TRIE
            )
            if token and token.strip()
        ]
    except Exception:
        return re.findall(r"[\u0E00-\u0E7F]+", text or "")


def _context_candidate_score(
    observed: str,
    candidate: str,
    context_probability: float,
) -> tuple:
    """
    คืนค่า (final_probability, similarity, distance)

    final_probability ผสมจาก:
    - ความน่าจะเป็นของคำในบริบท 78%
    - ความคล้ายด้านการสะกด 22%

    เพื่อไม่ให้ Context อย่างเดียวแก้คำที่รูปต่างกันมากเกินไป
    """
    similarity = similarity_score(observed, candidate) / 100.0
    distance = levenshtein_distance(observed, candidate)

    # ลดน้ำหนัก candidate ที่ต่างจากคำเดิมมากเกินไป
    distance_penalty = 1.0
    if distance == 2:
        distance_penalty = 0.88
    elif distance >= 3:
        distance_penalty = 0.55

    probability = (
        (0.78 * float(context_probability))
        + (0.22 * similarity)
    ) * distance_penalty

    return probability, similarity, distance


def detect_context_probability_errors(text: str) -> tuple:
    """
    ตรวจคำที่อาจ "ถูกสะกด แต่ผิดบริบท" ด้วยคะแนนความน่าจะเป็น

    คืนค่า:
        (context_errors, context_uncertain)
    """
    context_errors = []
    context_uncertain = []

    if not text:
        return context_errors, context_uncertain

    tokens = _context_tokens(text)
    if len(tokens) < 2:
        return context_errors, context_uncertain

    for index in range(1, len(tokens)):
        previous_word = tokens[index - 1]
        observed = tokens[index]

        if not contains_thai(observed) or len(observed) <= 1:
            continue

        priors = CONTEXT_NEXT_WORD_PRIORS.get(previous_word)
        if not priors:
            continue

        # ถ้าคำปัจจุบันเป็นคำที่เข้ากับบริบทอยู่แล้ว ไม่ต้องแก้
        if observed in priors:
            continue

        best = None

        for candidate, context_probability in priors.items():
            probability, similarity, distance = _context_candidate_score(
                observed,
                candidate,
                context_probability,
            )

            # Context correction ควรยังใกล้กับคำเดิมพอสมควร
            # เพื่อป้องกันการเดาคำใหม่แบบไกลเกินไป
            if distance > 2:
                continue

            item = {
                "word": observed,
                "suggested": candidate,
                "context_word": previous_word,
                "probability": probability,
                "similarity": similarity * 100.0,
                "distance": distance,
            }

            if best is None or item["probability"] > best["probability"]:
                best = item

        if not best:
            continue

        sentence = get_sentence_context(text, observed)
        probability = best["probability"]

        base = {
            "word": observed,
            "suggested": best["suggested"],
            "sentence": sentence,
            "distance": best["distance"],
            "similarity": best["similarity"],
            "confidence": probability,
            "context_probability": probability,
        }

        # >= 0.90: มีหลักฐานบริบท + รูปคำใกล้เคียงมากพอ -> คำผิด
        if probability >= 0.90:
            base.update({
                "detail": (
                    f"คำนี้สะกดได้ แต่มีความเป็นไปได้สูงว่าใช้ผิดบริบท "
                    f"หลังคำว่า '{previous_word}'"
                ),
                "method": "วิเคราะห์ความน่าจะเป็นตามบริบท",
            })
            context_errors.append(base)

        # 0.78 - 0.90: ยังไม่ชัดพอ -> ให้ผู้ใช้ตรวจสอบ
        elif probability >= 0.78:
            base.update({
                "reason": (
                    f"บริบทหลังคำว่า '{previous_word}' มีแนวโน้มเป็น "
                    f"'{best['suggested']}' มากกว่า"
                ),
                "detail": (
                    f"บริบทหลังคำว่า '{previous_word}' มีแนวโน้มเป็น "
                    f"'{best['suggested']}' มากกว่า"
                ),
                "method": "ประเมินความน่าจะเป็นตามบริบท",
            })
            context_uncertain.append(base)

    return context_errors, context_uncertain



def detect_generic_thai_context_errors(text: str) -> tuple:
    """
    ตรวจคำไทยที่สะกดถูก แต่มีแนวโน้มใช้ผิดบริบท
    โดยเปรียบเทียบ phrase frequency ของคำเดิมกับ candidate ใกล้เคียง
    ทำงานเมื่อมี wordfreq เท่านั้น
    """
    errors = []
    uncertain = []

    if not text or not WORDFREQ_AVAILABLE:
        return errors, uncertain

    tokens = _context_tokens(text)

    for i, observed in enumerate(tokens):
        if not contains_thai(observed) or len(observed) <= 1:
            continue
        if observed not in THAI_DICT:
            continue
        if is_json_correct_occurrence(observed, text):
            continue

        previous_word = tokens[i - 1] if i > 0 else ""
        next_word = tokens[i + 1] if i + 1 < len(tokens) else ""

        if not previous_word and not next_word:
            continue

        observed_context = _phrase_frequency_score(
            previous_word,
            observed,
            next_word,
            "th",
        )

        try:
            matches = process.extract(
                observed,
                THAI_DICT,
                scorer=fuzz.ratio,
                limit=12,
                score_cutoff=62,
            )
        except Exception:
            matches = []

        best = None

        for candidate, fuzzy_score, *_ in matches:
            if candidate == observed:
                continue

            distance = levenshtein_distance(observed, candidate)
            if distance > 1:
                continue

            similarity = float(fuzzy_score) / 100.0
            if similarity < 0.68:
                continue

            candidate_context = _phrase_frequency_score(
                previous_word,
                candidate,
                next_word,
                "th",
            )

            delta = candidate_context - observed_context

            item = {
                "word": observed,
                "suggested": candidate,
                "distance": distance,
                "similarity": similarity * 100.0,
                "candidate_context": candidate_context,
                "context_delta": delta,
            }

            if best is None or delta > best["context_delta"]:
                best = item

        if not best:
            continue

        delta = best["context_delta"]

        if delta < 0.14 or best["candidate_context"] < 0.28:
            continue

        sentence = get_sentence_context(text, observed)

        confidence = min(
            0.94,
            0.58
            + 0.70 * min(delta, 0.40)
            + 0.08 * (best["similarity"] / 100.0),
        )

        base = {
            "word": observed,
            "suggested": best["suggested"],
            "sentence": sentence,
            "distance": best["distance"],
            "similarity": best["similarity"],
            "confidence": confidence,
            "context_probability": best["candidate_context"],
            "context_delta": delta,
        }

        if delta >= 0.24 and confidence >= 0.78:
            base.update({
                "detail": (
                    "คำนี้สะกดถูกตามพจนานุกรม แต่ candidate ใกล้เคียง "
                    "มีความน่าจะเป็นในบริบทสูงกว่าอย่างชัดเจน"
                ),
                "method": "Phrase Context Probability",
            })
            errors.append(base)
        else:
            base.update({
                "reason": (
                    "คำสะกดถูก แต่บริบทมี candidate อื่นที่เป็นไปได้มากกว่า"
                ),
                "detail": (
                    "คำสะกดถูก แต่บริบทมี candidate อื่นที่เป็นไปได้มากกว่า"
                ),
                "method": "Phrase Context Probability",
            })
            uncertain.append(base)

    return errors, uncertain


# =========================================================
# BILINGUAL PROBABILITY / CANDIDATE RANKING ENGINE
# =========================================================

ENGLISH_FALLBACK_WORDS = {
    "a","about","after","again","all","also","am","an","and","are","as","at",
    "be","because","been","before","but","by","can","come","could","day","did",
    "do","does","done","down","each","even","first","for","from","get","go",
    "good","had","has","have","he","hello","her","here","him","his","how","i",
    "if","in","into","is","it","its","japan","just","know","like","look","made",
    "make","many","me","more","most","my","name","new","no","not","now","of",
    "on","one","only","or","other","our","out","over","people","right","said",
    "see","she","so","some","take","than","that","the","their","them","then",
    "there","these","they","thing","think","this","time","to","two","up","us",
    "use","very","want","was","way","we","well","were","what","when","where",
    "which","who","will","with","work","would","year","yes","you","your",
}

if WORDFREQ_AVAILABLE:
    try:
        ENGLISH_WORDS = set(top_n_list("en", 50000))
    except Exception:
        ENGLISH_WORDS = set(ENGLISH_FALLBACK_WORDS)
else:
    ENGLISH_WORDS = set(ENGLISH_FALLBACK_WORDS)

ENGLISH_WORDS.update(
    w.lower()
    for w in JSON_CORRECT_WORDS
    if re.fullmatch(r"[A-Za-z][A-Za-z'\-]*", str(w).strip())
)


def _zipf(word: str, lang: str) -> float:
    if not word:
        return 0.0
    if WORDFREQ_AVAILABLE:
        try:
            return max(0.0, float(zipf_frequency(word, lang)))
        except Exception:
            pass
    return 0.0


def _candidate_frequency_score(word: str, lang: str) -> float:
    z = _zipf(word, lang)
    if z <= 0:
        return 0.0
    return min(1.0, z / 7.0)


def _thai_candidate_pool(word: str, limit: int = 14) -> list:
    if not word:
        return []

    try:
        matches = process.extract(
            word,
            THAI_DICT,
            scorer=fuzz.ratio,
            limit=limit,
            score_cutoff=55,
        )
    except Exception:
        matches = []

    candidates = []
    for matched, score, *_ in matches:
        if not matched or matched == word:
            continue
        distance = levenshtein_distance(word, matched)
        max_distance = 2 if len(word) <= 8 else 3
        if distance > max_distance:
            continue

        candidates.append({
            "word": matched,
            "similarity": float(score) / 100.0,
            "distance": distance,
        })

    return candidates


def _english_candidate_pool(word: str, limit: int = 12) -> list:
    lower = word.lower()
    if not lower:
        return []

    try:
        matches = process.extract(
            lower,
            ENGLISH_WORDS,
            scorer=fuzz.ratio,
            limit=limit,
            score_cutoff=55,
        )
    except Exception:
        matches = []

    candidates = []
    for matched, score, *_ in matches:
        if matched == lower:
            continue
        distance = levenshtein_distance(lower, matched)
        max_distance = 1 if len(lower) <= 4 else 2
        if distance > max_distance:
            continue

        candidates.append({
            "word": matched,
            "similarity": float(score) / 100.0,
            "distance": distance,
        })

    return candidates


def _common_prefix_ratio(a: str, b: str) -> float:
    if not a or not b:
        return 0.0
    count = 0
    for x, y in zip(a, b):
        if x != y:
            break
        count += 1
    return count / max(len(a), len(b), 1)


def _common_suffix_ratio(a: str, b: str) -> float:
    if not a or not b:
        return 0.0
    count = 0
    for x, y in zip(reversed(a), reversed(b)):
        if x != y:
            break
        count += 1
    return count / max(len(a), len(b), 1)


def _word_structure_score(observed: str, candidate: str) -> float:
    prefix = _common_prefix_ratio(observed, candidate)
    suffix = _common_suffix_ratio(observed, candidate)
    length_ratio = min(len(observed), len(candidate)) / max(
        len(observed), len(candidate), 1
    )
    return min(
        1.0,
        0.50 * prefix
        + 0.20 * suffix
        + 0.30 * length_ratio
    )


def _phrase_frequency_score(
    previous_word: str,
    candidate: str,
    next_word: str,
    lang: str,
) -> float:
    if not WORDFREQ_AVAILABLE or not candidate:
        return 0.0

    phrases = []

    if previous_word:
        if lang == "th":
            phrases.extend([
                previous_word + candidate,
                previous_word + " " + candidate,
            ])
        else:
            phrases.append(previous_word + " " + candidate)

    if next_word:
        if lang == "th":
            phrases.extend([
                candidate + next_word,
                candidate + " " + next_word,
            ])
        else:
            phrases.append(candidate + " " + next_word)

    scores = []
    for phrase in phrases:
        z = _zipf(phrase, lang)
        if z > 0:
            scores.append(min(1.0, z / 7.0))

    return max(scores, default=0.0)


def _neighbor_context_score(
    previous_word: str,
    candidate: str,
    next_word: str,
    lang: str,
) -> float:
    lexical_score = _candidate_frequency_score(candidate, lang)
    phrase_score = _phrase_frequency_score(
        previous_word,
        candidate,
        next_word,
        lang,
    )

    if phrase_score > 0:
        score = 0.35 * lexical_score + 0.65 * phrase_score
    else:
        score = lexical_score

    if lang == "th" and previous_word:
        priors = CONTEXT_NEXT_WORD_PRIORS.get(previous_word, {})
        if candidate in priors:
            score = max(score, float(priors[candidate]))

    return min(1.0, score)

def rank_spelling_candidates(
    observed: str,
    previous_word: str = "",
    next_word: str = "",
    lang: str = "th",
):
    if lang == "en":
        pool = _english_candidate_pool(observed)
    else:
        pool = _thai_candidate_pool(observed)

    ranked = []

    for item in pool:
        distance = item["distance"]
        similarity = item["similarity"]

        if distance <= 1:
            distance_score = 1.0
        elif distance == 2:
            distance_score = 0.72
        else:
            distance_score = 0.48

        structure_score = _word_structure_score(
            observed,
            item["word"],
        )

        context_score = _neighbor_context_score(
            previous_word,
            item["word"],
            next_word,
            lang,
        )

        if lang == "th":
            final_score = (
                0.46 * similarity
                + 0.17 * distance_score
                + 0.22 * structure_score
                + 0.15 * context_score
            )
        else:
            final_score = (
                0.58 * similarity
                + 0.17 * distance_score
                + 0.10 * structure_score
                + 0.15 * context_score
            )

        ranked.append({
            **item,
            "structure_score": structure_score,
            "context_score": context_score,
            "score": final_score,
        })

    ranked.sort(key=lambda x: x["score"], reverse=True)

    if not ranked:
        return None

    best = ranked[0]
    second_score = ranked[1]["score"] if len(ranked) > 1 else 0.0
    best["margin"] = best["score"] - second_score
    return best

def _thai_analysis_tokens(text: str) -> list:
    try:
        return [
            t.strip()
            for t in word_tokenize(
                text,
                engine="newmm",
                custom_dict=CUSTOM_TRIE,
            )
            if t and t.strip()
        ]
    except Exception:
        return re.findall(r"[\u0E00-\u0E7F]+|[A-Za-z]+", text or "")


def _english_tokens(text: str) -> list:
    return re.findall(r"\b[A-Za-z][A-Za-z'\-]*\b", text or "")


def detect_english_errors(text: str) -> tuple:
    errors = []
    uncertain = []
    tokens = _english_tokens(text)
    seen = set()

    for i, raw_word in enumerate(tokens):
        lower = raw_word.lower()
        if lower in seen or len(lower) <= 1:
            continue
        seen.add(lower)

        if is_json_correct_word(raw_word) or is_json_correct_word(lower):
            continue

        if lower in ENGLISH_WORDS:
            continue

        prev_word = tokens[i - 1].lower() if i > 0 else ""
        next_word = tokens[i + 1].lower() if i + 1 < len(tokens) else ""

        best = rank_spelling_candidates(
            lower,
            prev_word,
            next_word,
            lang="en",
        )

        sentence = get_sentence_context(text, raw_word)

        if not best:
            uncertain.append({
                "word": raw_word,
                "reason": "ไม่พบคำภาษาอังกฤษนี้ในพจนานุกรม",
                "sentence": sentence,
                "suggested": "ควรตรวจสอบ",
                "confidence": 0.45,
                "method": "ตรวจภาษาอังกฤษ",
            })
            continue

        confidence = min(
            0.98,
            0.56 + 0.34 * best["similarity"] + 0.08 * min(1.0, best["margin"] * 5),
        )

        result = {
            "word": raw_word,
            "suggested": best["word"],
            "sentence": sentence,
            "distance": best["distance"],
            "similarity": best["similarity"] * 100.0,
            "confidence": confidence,
            "method": "ตรวจสะกดภาษาอังกฤษด้วย Candidate Ranking",
        }

        if best["distance"] == 1 and best["similarity"] >= 0.72:
            result["detail"] = "สะกดคำภาษาอังกฤษไม่ถูกต้อง"
            errors.append(result)
        elif best["score"] >= 0.72 and best["margin"] >= 0.03:
            result["detail"] = "คำภาษาอังกฤษมีแนวโน้มสะกดไม่ถูกต้อง"
            errors.append(result)
        else:
            result["reason"] = "พบคำภาษาอังกฤษที่ควรตรวจสอบเพิ่มเติม"
            result["detail"] = result["reason"]
            uncertain.append(result)

    return errors, uncertain


def detect_thai_ranked_errors(text: str) -> tuple:
    errors = []
    uncertain = []
    tokens = _thai_analysis_tokens(text)

    covered_indexes = set()

    # A) รวม token 2-3 ชิ้นก่อน เช่น โรง + เรียร -> โรงเรียร
    for width in (3, 2):
        for i in range(0, len(tokens) - width + 1):
            indexes = list(range(i, i + width))

            if any(idx in covered_indexes for idx in indexes):
                continue

            parts = tokens[i:i + width]

            if not all(
                contains_thai(part)
                for part in parts
            ):
                continue

            observed = "".join(parts)

            if not (4 <= len(observed) <= 24):
                continue
            if observed in THAI_DICT:
                continue
            if is_json_correct_occurrence(observed, text):
                continue

            previous_word = tokens[i - 1] if i > 0 else ""
            next_word = (
                tokens[i + width]
                if i + width < len(tokens)
                else ""
            )

            best = rank_spelling_candidates(
                observed,
                previous_word,
                next_word,
                lang="th",
            )

            if not best:
                continue

            strong = (
                best["distance"] <= 1
                and best["similarity"] >= 0.84
                and best.get("structure_score", 0) >= 0.60
                and abs(len(observed) - len(best["word"])) <= 1
            )

            if not strong:
                continue

            sentence = get_sentence_context(text, observed)

            confidence = min(
                0.99,
                0.70
                + 0.20 * best["similarity"]
                + 0.07 * best.get("structure_score", 0)
                + 0.02 * min(1.0, best.get("margin", 0) * 5),
            )

            errors.append({
                "word": observed,
                "suggested": best["word"],
                "sentence": sentence,
                "distance": best["distance"],
                "similarity": best["similarity"] * 100.0,
                "confidence": confidence,
                "candidate_margin": best.get("margin", 0),
                "detail": (
                    "ตรวจพบว่าตัวแบ่งคำแยกคำสะกดผิดออกเป็นหลายส่วน "
                    "จึงรวมคำก่อนจัดอันดับคำแนะนำ"
                ),
                "method": (
                    "รวม Token + Candidate Ranking + Context Probability"
                ),
            })

            covered_indexes.update(indexes)

    # B) ตรวจ token เดี่ยวที่ไม่ถูก reconstruction ครอบคลุม
    seen = set()

    for i, word in enumerate(tokens):
        if i in covered_indexes:
            continue
        if not contains_thai(word) or len(word) <= 1:
            continue
        if word in seen:
            continue
        seen.add(word)

        if is_json_correct_occurrence(word, text):
            continue
        if word in THAI_DICT:
            continue
        if word in COMMON_MISTAKES:
            continue

        previous_word = tokens[i - 1] if i > 0 else ""
        next_word = tokens[i + 1] if i + 1 < len(tokens) else ""

        best = rank_spelling_candidates(
            word,
            previous_word,
            next_word,
            lang="th",
        )

        sentence = get_sentence_context(text, word)

        if not best:
            uncertain.append({
                "word": word,
                "reason": (
                    "คำนี้ไม่พบในพจนานุกรมและยังไม่มีคำแนะนำ "
                    "ที่ชนะ candidate อื่นชัดเจน"
                ),
                "sentence": sentence,
                "suggested": "ควรตรวจสอบ",
                "confidence": 0.45,
                "method": "Candidate Ranking ภาษาไทย",
            })
            continue

        confidence = min(
            0.97,
            0.47
            + 0.34 * best["similarity"]
            + 0.10 * best.get("structure_score", 0)
            + 0.06 * min(1.0, best.get("margin", 0) * 5),
        )

        item = {
            "word": word,
            "suggested": best["word"],
            "sentence": sentence,
            "distance": best["distance"],
            "similarity": best["similarity"] * 100.0,
            "confidence": confidence,
            "candidate_margin": best.get("margin", 0),
            "method": (
                "จัดอันดับคำแนะนำด้วยรูปคำ ความคล้าย "
                "และความน่าจะเป็นตามบริบท"
            ),
        }

        if (
            best["distance"] == 1
            and best["similarity"] >= 0.82
            and best.get("structure_score", 0) >= 0.55
            and best.get("margin", 0) >= 0.015
        ):
            item["detail"] = (
                "สะกดคำไม่ถูกต้อง โดย candidate อันดับสูงสุด "
                "มีรูปคำและโครงสร้างใกล้เคียง"
            )
            errors.append(item)
        elif (
            best["score"] >= 0.78
            and best.get("margin", 0) >= 0.045
        ):
            item["detail"] = (
                "คำนี้มีแนวโน้มสะกดไม่ถูกต้องจากคะแนนรวมหลายปัจจัย"
            )
            errors.append(item)
        else:
            item["reason"] = (
                "มี candidate หลายคำที่คะแนนใกล้กัน จึงจัดเป็นยังไม่แน่ใจ"
            )
            item["detail"] = item["reason"]
            uncertain.append(item)

    return errors, uncertain


# =========================================================
# TOKEN SPELL CHECK
# =========================================================

def detect_token_errors(
    text: str
) -> tuple:
    errors = []
    uncertain = []

    if not text:
        return errors, uncertain

    try:
        tokens = word_tokenize(
            text,
            engine="newmm",
            custom_dict=CUSTOM_TRIE,
        )
    except Exception:
        tokens = re.findall(r"[\u0E00-\u0E7F]+", text)

    seen = set()
    for token in tokens:
        word = token.strip()
        if (
            not word
            or word in seen
            or not contains_thai(word)
            or len(word) <= 1
        ):
            continue

        seen.add(word)

        if is_json_correct_occurrence(word, text):
            continue
        if word in THAI_DICT:
            continue

        if word in COMMON_MISTAKES:
            info = COMMON_MISTAKES[word]
            errors.append({
                "word": word,
                "suggested": info["correct"],
                "detail": info["reason"],
                "confidence": info["confidence"],
                "method": "ตรวจจากคลังคำผิดที่รู้จัก",
                "sentence": get_sentence_context(text, word),
            })

    ranked_errors, ranked_uncertain = detect_thai_ranked_errors(text)

    known_words = {e["word"] for e in errors}
    errors.extend(
        e for e in ranked_errors
        if e.get("word") not in known_words
    )
    uncertain.extend(
        u for u in ranked_uncertain
        if u.get("word") not in known_words
    )

    return errors, uncertain


# =========================================================
# DEDUPLICATE
# =========================================================

def deduplicate_errors(
    errors: list
) -> list:

    unique = {}

    for error in errors:

        key = (
            error.get(
                "word",
                ""
            ),
            error.get(
                "suggested",
                ""
            ),
            error.get(
                "sentence",
                ""
            )
        )

        if key not in unique:

            unique[key] = error

        else:

            old = unique[key]

            if (
                error.get(
                    "confidence",
                    0
                )
                >
                old.get(
                    "confidence",
                    0
                )
            ):

                unique[key] = error

    return list(
        unique.values()
    )


# =========================================================
# MAIN ANALYSIS ENGINE
# =========================================================

def analyze_text_tokens(
    text: str,
    location: str,
    page_num: int,
    pdf_page=None
):

    text = normalize_text(
        text
    )

    page_errors = []
    page_uncertain = []

    if not text:

        return (
            page_errors,
            page_uncertain
        )

    # =====================================================
    # ENGINE 1
    # Known Mistakes
    # =====================================================

    known_mistakes = detect_known_mistakes(
        text
    )

    for item in known_mistakes:

        word = item["wrong"]

        suggested = item["correct"]

        sentence = get_sentence_context(
            text,
            word
        )

        highlighted = highlight_word_in_html(
            sentence,
            word
        )

        page_errors.append({
            "page_num": page_num,
            "location": location,
            "word": word,
            "suggested": suggested,
            "detail": item["reason"],
            "confidence": item["confidence"],
            "method": "ตรวจจากคลังคำผิดที่รู้จัก",
            "sentence": sentence,
            "highlighted_sentence": highlighted,
            "bboxes": (
                find_word_bboxes(
                    pdf_page,
                    word
                )
                if pdf_page is not None
                else []
            ),
        })

    # =====================================================
    # ENGINE 2
    # Phrase Fuzzy
    # =====================================================

    fuzzy_phrases = detect_fuzzy_phrases(
        text
    )

    for item in fuzzy_phrases:

        word = item["word"]

        suggested = item["suggested"]

        already_exists = any(
            e["word"] == word
            and e["suggested"] == suggested
            for e in page_errors
        )

        if already_exists:
            continue

        sentence = get_context_for_phrase(
            text,
            word
        )

        highlighted = highlight_word_in_html(
            sentence,
            word
        )

        distance = item["distance"]

        if distance == 1:

            detail = (
                "พบการสะกดผิดเล็กน้อย "
                "โดยมีความแตกต่างจากคำที่ถูกต้อง 1 ตำแหน่ง"
            )

        elif distance <= 2:

            detail = (
                "พบคำที่มีรูปแบบใกล้เคียง "
                "กับคำเฉพาะทางในระบบ"
            )

        else:

            detail = (
                "พบวลีที่มีความคล้ายกับ "
                "วลีเฉพาะทางในระบบ"
            )

        page_errors.append({
            "page_num": page_num,
            "location": location,
            "word": word,
            "suggested": suggested,
            "detail": detail,
            "confidence": item["confidence"],
            "method": "จับคู่วลีใกล้เคียง",
            "similarity": item["score"],
            "distance": distance,
            "sentence": sentence,
            "highlighted_sentence": highlighted,
            "bboxes": (
                find_word_bboxes(
                    pdf_page,
                    word
                )
                if pdf_page is not None
                else []
            ),
        })

    # =====================================================
    # ENGINE 3
    # Token Spell Check
    # =====================================================

    token_errors, token_uncertain = (
        detect_token_errors(
            text
        )
    )

    for error in token_errors:

        word = error["word"]

        suggested = error["suggested"]

        sentence = error["sentence"]

        highlighted = highlight_word_in_html(
            sentence,
            word
        )

        page_errors.append({
            "page_num": page_num,
            "location": location,
            "word": word,
            "suggested": suggested,
            "detail": error["detail"],
            "confidence": error.get(
                "confidence",
                0
            ),
            "method": error.get(
                "method",
                "ตรวจการสะกดคำ"
            ),
            "distance": error.get(
                "distance"
            ),
            "sentence": sentence,
            "highlighted_sentence": highlighted,
            "bboxes": (
                find_word_bboxes(
                    pdf_page,
                    word
                )
                if pdf_page is not None
                else []
            ),
        })

    # =====================================================
    # ENGINE 4
    # Context Probability
    # =====================================================

    context_errors, context_uncertain = (
        detect_context_probability_errors(text)
    )

    for error in context_errors:
        word = error["word"]
        suggested = error["suggested"]

        # ถ้า Engine เดิมพบคำเดียวกันและแนะนำเหมือนกันอยู่แล้ว ไม่เพิ่มซ้ำ
        already_exists = any(
            e.get("word") == word
            and e.get("suggested") == suggested
            for e in page_errors
        )
        if already_exists:
            continue

        sentence = error["sentence"]
        highlighted = highlight_word_in_html(sentence, word)

        page_errors.append({
            "page_num": page_num,
            "location": location,
            "word": word,
            "suggested": suggested,
            "detail": error["detail"],
            "confidence": error.get("confidence", 0),
            "method": error.get(
                "method",
                "วิเคราะห์ความน่าจะเป็นตามบริบท"
            ),
            "distance": error.get("distance"),
            "similarity": error.get("similarity"),
            "sentence": sentence,
            "highlighted_sentence": highlighted,
            "bboxes": (
                find_word_bboxes(pdf_page, word)
                if pdf_page is not None
                else []
            ),
        })

    # เพิ่มผลที่ Context Engine ยังไม่มั่นใจ เข้า token_uncertain
    # เพื่อให้ใช้ UI / กรอบสีน้ำเงินเดิมต่อได้ทันที
    for item in context_uncertain:
        if any(e.get("word") == item.get("word") for e in page_errors):
            continue
        if any(u.get("word") == item.get("word") for u in token_uncertain):
            continue

        token_uncertain.append({
            "word": item["word"],
            "reason": item["reason"],
            "sentence": item["sentence"],
            "suggested": item.get("suggested", "ควรตรวจสอบ"),
            "confidence": item.get("confidence", 0),
            "method": item.get(
                "method",
                "ประเมินความน่าจะเป็นตามบริบท"
            ),
            "distance": item.get("distance"),
        })

    # =====================================================
    # =====================================================
    # ENGINE 4B
    # Generic Thai Phrase Context Probability
    # =====================================================

    generic_context_errors, generic_context_uncertain = (
        detect_generic_thai_context_errors(text)
    )

    for error in generic_context_errors:
        word = error["word"]
        suggested = error["suggested"]

        if any(
            e.get("word") == word
            and e.get("suggested") == suggested
            for e in page_errors
        ):
            continue

        sentence = error["sentence"]

        page_errors.append({
            "page_num": page_num,
            "location": location,
            "word": word,
            "suggested": suggested,
            "detail": error["detail"],
            "confidence": error.get("confidence", 0),
            "method": error.get(
                "method",
                "Phrase Context Probability",
            ),
            "distance": error.get("distance"),
            "similarity": error.get("similarity"),
            "sentence": sentence,
            "highlighted_sentence": highlight_word_in_html(
                sentence,
                word,
            ),
            "bboxes": (
                find_word_bboxes(pdf_page, word)
                if pdf_page is not None
                else []
            ),
        })

    for item in generic_context_uncertain:
        if any(
            e.get("word") == item.get("word")
            for e in page_errors
        ):
            continue
        token_uncertain.append(item)

    # ENGINE 5
    # English Spell Check
    # =====================================================

    english_errors, english_uncertain = detect_english_errors(text)

    for error in english_errors:
        word = error["word"]

        if any(
            e.get("word", "").lower() == word.lower()
            for e in page_errors
        ):
            continue

        sentence = error["sentence"]
        page_errors.append({
            "page_num": page_num,
            "location": location,
            "word": word,
            "suggested": error["suggested"],
            "detail": error["detail"],
            "confidence": error.get("confidence", 0),
            "method": error.get("method", "ตรวจภาษาอังกฤษ"),
            "distance": error.get("distance"),
            "similarity": error.get("similarity"),
            "sentence": sentence,
            "highlighted_sentence": highlight_word_in_html(sentence, word),
            "bboxes": (
                find_word_bboxes(pdf_page, word)
                if pdf_page is not None
                else []
            ),
        })

    for item in english_uncertain:
        if any(
            e.get("word", "").lower() == item.get("word", "").lower()
            for e in page_errors
        ):
            continue
        token_uncertain.append(item)

    # =====================================================
    # UNCERTAIN
    # =====================================================

    for uncertain in token_uncertain:

        word = uncertain["word"]

        exists_as_error = any(
            e["word"] == word
            for e in page_errors
        )

        if exists_as_error:
            continue

        sentence = uncertain[
            "sentence"
        ]

        highlighted = highlight_word_in_html(
            sentence,
            word
        )

        page_uncertain.append({
            "page_num": page_num,
            "location": location,
            "word": word,
            "reason": uncertain["reason"],
            "detail": uncertain["reason"],
            "suggested": uncertain.get("suggested", "ควรตรวจสอบ"),
            "confidence": uncertain.get("confidence", 0.0),
            "method": uncertain.get("method", "ยังไม่แน่ใจ"),
            "distance": uncertain.get("distance"),
            "status": "uncertain",
            "sentence": sentence,
            "highlighted_sentence": highlighted,
            "bboxes": (
                find_word_bboxes(pdf_page, word)
                if pdf_page is not None
                else []
            ),
        })

    # =====================================================
    # JSON WHITELIST OVERRIDE
    # =====================================================

    # ไม่เปลี่ยน Logic ของ Engine เดิม แต่กรองผลขั้นสุดท้าย:
    # คำที่ผู้ใช้กำหนดไว้ใน dictionary.json จะถือว่าเป็นคำถูกเสมอ
    page_errors = [
        error
        for error in page_errors
        if not is_json_correct_occurrence(
            error.get("word", ""),
            error.get("sentence", text)
        )
    ]

    page_uncertain = [
        item
        for item in page_uncertain
        if not is_json_correct_occurrence(
            item.get("word", ""),
            item.get("sentence", text)
        )
    ]

    # =====================================================
    # DEDUPLICATE
    # =====================================================

    page_errors = deduplicate_errors(
        page_errors
    )

    # =====================================================
    # ADD ERROR ID
    # =====================================================

    for index, error in enumerate(
        page_errors
    ):

        error["error_id"] = (
            create_error_id(
                page_num,
                index
            )
        )
        error["status"] = "error"

    for index, item in enumerate(page_uncertain):
        item["error_id"] = f"uncertain-page-{page_num}-item-{index}"
        item["status"] = "uncertain"

    return (
        page_errors,
        page_uncertain
    )


# =========================================================
# PDF PROCESSING
# =========================================================

def process_full_pdf(
    file_bytes
):

    doc = fitz.open(
        stream=file_bytes,
        filetype="pdf"
    )

    pages_data = []

    all_errors = []
    all_uncertain = []

    for i, page in enumerate(
        doc
    ):

        text = page.get_text() or ""

        location = f"หน้า {i + 1}"

        page_num = i + 1

        pix = page.get_pixmap(
            dpi=150
        )

        img_bytes = pix.tobytes(
            "png"
        )

        orig_image = Image.open(
            io.BytesIO(img_bytes)
        )

        page_errors = []
        page_uncertain = []

        if text.strip():

            (
                page_errors,
                page_uncertain
            ) = analyze_text_tokens(
                text,
                location,
                page_num,
                pdf_page=page
            )

            all_errors.extend(
                page_errors
            )

            all_uncertain.extend(
                page_uncertain
            )

        annotated_image = (
            draw_red_boxes_on_image(
                orig_image,
                page,
                page_errors
            )
        )

        pages_data.append({
            "page_num": page_num,
            "location": location,
            "text": text,
            "image": annotated_image,
            "errors": page_errors,
            "uncertain": page_uncertain,
            "pdf_page": page,
        })

    # ไม่ปิด doc ตรงนี้
    # เพราะ preview ใช้ page object
    # หลังจาก analysis เสร็จ

    return (
        pages_data,
        all_errors,
        all_uncertain
    )



# =========================================================
# EXPORT ANNOTATED PDF
# =========================================================

def export_annotated_pdf(
    original_pdf_bytes: bytes,
    all_errors: list,
    all_uncertain: list,
) -> bytes:
    """
    สร้าง PDF สำเนาจากไฟล์ต้นฉบับ แล้ววาดกรอบถาวรลงใน PDF

    สี:
    - คำผิด = สีแดง
    - ยังไม่แน่ใจ = สีน้ำเงิน

    ใช้ bbox ที่ได้จากขั้นตอนวิเคราะห์เดิม จึงไม่เปลี่ยนข้อความ
    หรือ Layout ของ PDF ต้นฉบับ
    """
    if not original_pdf_bytes:
        return b""

    doc = fitz.open(
        stream=original_pdf_bytes,
        filetype="pdf",
    )

    try:
        items_by_page = {}

        for item in all_errors:
            page_num = int(item.get("page_num", 0) or 0)
            if page_num <= 0:
                continue

            export_item = dict(item)
            export_item["status"] = "error"

            items_by_page.setdefault(
                page_num,
                []
            ).append(export_item)

        for item in all_uncertain:
            page_num = int(item.get("page_num", 0) or 0)
            if page_num <= 0:
                continue

            export_item = dict(item)
            export_item["status"] = "uncertain"

            items_by_page.setdefault(
                page_num,
                []
            ).append(export_item)

        for page_num, items in items_by_page.items():

            page_index = page_num - 1

            if (
                page_index < 0
                or page_index >= len(doc)
            ):
                continue

            page = doc[page_index]

            for item in items:

                bboxes = item.get(
                    "bboxes",
                    []
                )

                if not bboxes:
                    continue

                is_uncertain = (
                    item.get("status")
                    == "uncertain"
                )

                # PyMuPDF RGB ใช้ช่วง 0..1
                stroke_color = (
                    (0.145, 0.388, 0.922)
                    if is_uncertain
                    else (0.937, 0.161, 0.161)
                )

                for bbox in bboxes:

                    if (
                        not bbox
                        or len(bbox) < 4
                    ):
                        continue

                    x0, y0, x1, y1 = bbox[:4]

                    # ขยายกรอบเล็กน้อยเพื่อไม่ทับตัวอักษร
                    rect = fitz.Rect(
                        float(x0) - 1.5,
                        float(y0) - 1.5,
                        float(x1) + 1.5,
                        float(y1) + 1.5,
                    )

                    # จำกัดไม่ให้กรอบเลยขอบหน้า
                    rect = rect & page.rect

                    if rect.is_empty:
                        continue

                    page.draw_rect(
                        rect,
                        color=stroke_color,
                        width=1.7,
                        overlay=True,
                    )

        output = io.BytesIO()

        doc.save(
            output,
            garbage=4,
            deflate=True,
            clean=True,
        )

        return output.getvalue()

    finally:
        doc.close()



# =========================================================
# EXPORT ANALYSIS REPORT — CSV / EXCEL / JSON
# =========================================================

def _report_status_text(status: str) -> str:
    return "คำผิด" if status == "error" else "ยังไม่แน่ใจ"


def build_analysis_report_rows(
    errors: list,
    uncertain: list,
) -> list:
    """
    รวมผลวิเคราะห์เป็นตารางกลางสำหรับ Export CSV / Excel / JSON
    """
    rows = []

    def add_rows(items, status):
        for item in items:
            confidence = item.get("confidence", 0)

            try:
                confidence_pct = round(float(confidence) * 100, 1)
            except Exception:
                confidence_pct = 0.0

            suggested = (
                item.get("suggested")
                or item.get("correct")
                or ""
            )

            detail = (
                item.get("detail")
                or item.get("reason")
                or ""
            )

            rows.append({
                "สถานะ": _report_status_text(status),
                "หน้า": item.get("page_num", ""),
                "ตำแหน่ง": item.get("location", ""),
                "คำที่พบ": item.get("word", ""),
                "คำแนะนำ": suggested,
                "รายละเอียด": detail,
                "ความมั่นใจ (%)": confidence_pct,
                "วิธีตรวจพบ": item.get("method", ""),
                "ข้อความบริบท": item.get("sentence", ""),
            })

    add_rows(errors, "error")
    add_rows(uncertain, "uncertain")

    return rows


def export_report_csv(rows: list) -> bytes:
    """
    CSV ใช้ UTF-8 BOM เพื่อให้ Excel บน Windows เปิดภาษาไทยได้ถูกต้อง
    """
    output = io.StringIO()

    headers = [
        "สถานะ",
        "หน้า",
        "ตำแหน่ง",
        "คำที่พบ",
        "คำแนะนำ",
        "รายละเอียด",
        "ความมั่นใจ (%)",
        "วิธีตรวจพบ",
        "ข้อความบริบท",
    ]

    writer = csv.DictWriter(
        output,
        fieldnames=headers,
        extrasaction="ignore",
    )

    writer.writeheader()

    for row in rows:
        writer.writerow(row)

    return output.getvalue().encode("utf-8-sig")


def export_report_json(rows: list) -> bytes:
    """
    JSON ภาษาไทยแบบอ่านง่าย
    """
    return json.dumps(
        {
            "summary": {
                "total": len(rows),
                "errors": sum(
                    1 for row in rows
                    if row.get("สถานะ") == "คำผิด"
                ),
                "uncertain": sum(
                    1 for row in rows
                    if row.get("สถานะ") == "ยังไม่แน่ใจ"
                ),
            },
            "results": rows,
        },
        ensure_ascii=False,
        indent=2,
    ).encode("utf-8")


def export_report_excel(rows: list) -> bytes:
    """
    Excel พร้อมหัวตาราง, Freeze header และ Auto Filter
    """
    if not OPENPYXL_AVAILABLE:
        return b""

    wb = Workbook()
    ws = wb.active
    ws.title = "ผลการตรวจคำ"

    headers = [
        "สถานะ",
        "หน้า",
        "ตำแหน่ง",
        "คำที่พบ",
        "คำแนะนำ",
        "รายละเอียด",
        "ความมั่นใจ (%)",
        "วิธีตรวจพบ",
        "ข้อความบริบท",
    ]

    ws.append(headers)

    for row in rows:
        ws.append([
            row.get(header, "")
            for header in headers
        ])

    # Header
    header_fill = PatternFill(
        "solid",
        fgColor="0F4C81",
    )

    header_font = Font(
        bold=True,
        color="FFFFFF",
    )

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
        )

    # สีสถานะ
    error_fill = PatternFill(
        "solid",
        fgColor="FDE8E8",
    )
    uncertain_fill = PatternFill(
        "solid",
        fgColor="E7F0FF",
    )

    for row_idx in range(2, ws.max_row + 1):
        status_cell = ws.cell(row=row_idx, column=1)
        if status_cell.value == "คำผิด":
            status_cell.fill = error_fill
        elif status_cell.value == "ยังไม่แน่ใจ":
            status_cell.fill = uncertain_fill

        for col_idx in range(1, len(headers) + 1):
            ws.cell(
                row=row_idx,
                column=col_idx,
            ).alignment = Alignment(
                vertical="top",
                wrap_text=True,
            )

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    widths = {
        1: 15,
        2: 9,
        3: 18,
        4: 22,
        5: 22,
        6: 38,
        7: 18,
        8: 28,
        9: 55,
    }

    for col_idx, width in widths.items():
        ws.column_dimensions[
            get_column_letter(col_idx)
        ].width = width

    ws.row_dimensions[1].height = 26

    output = io.BytesIO()
    wb.save(output)

    return output.getvalue()


# =========================================================
# DOCX PROCESSING
# =========================================================

def process_full_docx(
    file_bytes
):

    doc = docx.Document(
        io.BytesIO(file_bytes)
    )

    pages_data = []

    all_errors = []
    all_uncertain = []

    for i, para in enumerate(
        doc.paragraphs
    ):

        text = para.text.strip()

        if not text:
            continue

        location = (
            f"ย่อหน้า {i + 1}"
        )

        page_num = i + 1

        (
            page_errors,
            page_uncertain
        ) = analyze_text_tokens(
            text,
            location,
            page_num
        )

        all_errors.extend(
            page_errors
        )

        all_uncertain.extend(
            page_uncertain
        )

        pages_data.append({
            "page_num": page_num,
            "location": location,
            "text": text,
            "image": None,
            "errors": page_errors,
            "uncertain": page_uncertain,
        })

    return (
        pages_data,
        all_errors,
        all_uncertain
    )


# =========================================================
# DISPLAY HELPERS
# =========================================================

def confidence_label(
    confidence
):

    if confidence >= 0.95:
        return "สูงมาก"

    if confidence >= 0.90:
        return "สูง"

    if confidence >= 0.80:
        return "ปานกลาง"

    return "ต่ำ"


def render_error_card(
    err
):

    error_id = err.get(
        "error_id",
        ""
    )

    confidence = err.get(
        "confidence",
        0
    )

    method = err.get(
        "method",
        "-"
    )

    # -----------------------------------------------------
    # Anchor
    # -----------------------------------------------------

    st.markdown(
        f"""
        <div
            id="{error_id}"
            style="
                scroll-margin-top:100px;
                height:1px;
            "
        ></div>
        """,
        unsafe_allow_html=True
    )

    st.markdown(
        "**ประโยคต้นฉบับ:**"
    )

    st.markdown(
        f"> {err['highlighted_sentence']}",
        unsafe_allow_html=True
    )

    st.markdown(
        f"- **คำที่พบ:** `{err['word']}`"
    )

    st.markdown(
        f"- **คำที่ถูกต้อง:** "
        f":green[**{err['suggested']}**]"
    )

    st.markdown(
        f"- **รายละเอียด:** "
        f"{err['detail']}"
    )

    st.markdown(
        f"- **วิธีตรวจพบ:** "
        f"`{method}`"
    )

    if "similarity" in err:

        st.markdown(
            f"- **ความคล้าย:** "
            f"{err['similarity']:.1f}%"
        )

    if err.get("distance") is not None:

        st.markdown(
            f"- **ระยะความแตกต่างของคำ:** "
            f"{err['distance']}"
        )

    st.markdown(
        f"- **ความมั่นใจ:** "
        f"**{confidence_label(confidence)}** "
        f"({confidence * 100:.0f}%)"
    )

    # -----------------------------------------------------
    # Preview Link
    # -----------------------------------------------------

    if err.get("bboxes"):

        st.markdown(
            f"""
            <a
                href="#preview-{error_id}"
                style="
                    display:inline-block;
                    margin-top:8px;
                    padding:6px 10px;
                    border-radius:6px;
                    background:#fef2f2;
                    color:#b91c1c;
                    text-decoration:none;
                    font-size:13px;
                    font-weight:600;
                "
            >
                ดูตำแหน่งคำนี้บนเอกสาร
            </a>
            """,
            unsafe_allow_html=True
        )



# =========================================================
# รูปแบบหน้าจอ
# =========================================================

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Kanit:wght@300;400;500;600;700;800&display=swap');

    :root {
        --navy:#063b6d;
        --navy2:#0b4f87;
        --line:#d7e0ea;
        --soft:#f4f6f8;
        --danger:#c81e1e;
        --text:#1f2937;
        --muted:#64748b;
    }

    html { scroll-behavior:smooth; }

    html, body, [class*="css"], [data-testid="stAppViewContainer"],
    [data-testid="stSidebar"], button, input, textarea, select {
        font-family:"Kanit","Noto Sans Thai",sans-serif !important;
    }

    [data-testid="stAppViewContainer"] { background:#f5f7f9; }
    [data-testid="stHeader"] { background:transparent; height:0; }
    [data-testid="stToolbar"] { visibility:hidden; }

    [data-testid="stSidebar"] {
        background:#f4f6f8;
        border-right:1px solid var(--line);
    }

    [data-testid="stSidebar"] .block-container {
        padding-top:1rem;
        padding-left:.9rem;
        padding-right:.9rem;
    }

    .block-container {
        max-width:1900px;
        padding-top:.75rem;
        padding-bottom:2rem;
        padding-left:1.15rem;
        padding-right:1.15rem;
    }

    .topbar {
        background:#fff;
        border:1px solid var(--line);
        margin:-.75rem -1.15rem .9rem -1.15rem;
        padding:.72rem 1.15rem;
        display:flex;
        align-items:center;
        justify-content:space-between;
    }

    .brand-wrap { display:flex; align-items:center; gap:1.35rem; min-width:0; }
    .brand { color:var(--navy); font-size:1.3rem; font-weight:800; white-space:nowrap; }
    .nav { display:flex; gap:1.35rem; font-size:.9rem; }
    .nav span { color:#334155; padding:.35rem 0; }
    .nav .active { color:var(--navy); font-weight:700; border-bottom:2px solid var(--navy); }

    .side-brand { color:var(--navy); font-size:1.05rem; font-weight:800; margin:.15rem 0 1rem; }

    .file-card {
        background:#fff;
        border:1px solid #b9c9d9;
        border-radius:8px;
        padding:.8rem;
        margin:.25rem 0 .65rem;
    }

    .file-name { color:var(--navy); font-weight:700; word-break:break-word; }
    .tiny { color:#64748b; font-size:.76rem; line-height:1.55; }
    .menu-title { color:#41566d; font-size:.78rem; font-weight:700; margin:.9rem 0 .35rem; }
    .side-item { padding:.62rem .65rem; border-radius:6px; margin:.15rem 0; color:#23384d; }
    .side-item.active {
        background:#dde5ec;
        border-left:3px solid var(--navy);
        font-weight:700;
        color:var(--navy);
    }

    .count-badge {
        float:right;
        min-width:24px;
        text-align:center;
        border-radius:999px;
        padding:.05rem .4rem;
        font-size:.72rem;
        background:#fee2e2;
        color:#c81e1e;
    }

    .count-badge.warn { background:#ffedd5; color:#a94c00; }

    .intro-card,
    .ready-card {
        max-width:760px;
        margin:2rem auto 1rem;
        background:#fff;
        border:1px solid #c9d3df;
        border-radius:10px;
        padding:1.4rem 1.6rem;
    }

    .intro-title,
    .ready-title {
        text-align:center;
        color:var(--navy);
        font-size:1.45rem;
        font-weight:800;
        margin-bottom:.2rem;
    }

    .intro-sub,
    .ready-sub {
        text-align:center;
        color:#526579;
        font-size:.88rem;
    }

    .doc-toolbar {
        background:#fff;
        border:1px solid var(--line);
        padding:.58rem .75rem;
        display:flex;
        align-items:center;
        justify-content:space-between;
        color:#334155;
        font-size:.82rem;
        position:sticky;
        top:0;
        z-index:3;
    }

    .page-chip { font-weight:700; color:#1e293b; }

    .result-head {
        background:#fff;
        border-bottom:1px solid var(--line);
        padding:.7rem .15rem .55rem;
        margin-bottom:.65rem;
        position:sticky;
        top:0;
        z-index:4;
    }

    .analysis-focus-panel {
        display:none;
        margin:.65rem 0 .8rem;
    }

    .analysis-focus-panel.active {
        display:block;
    }

    .analysis-focus-card {
        background:#ffffff;
        border:1px solid #dbe3ec;
        border-radius:10px;
        padding:16px;
        box-shadow:0 4px 14px rgba(15,23,42,.06);
    }

    .analysis-focus-header {
        display:flex;
        justify-content:space-between;
        align-items:flex-start;
        gap:12px;
        margin-bottom:14px;
    }

    .analysis-focus-kicker {
        font-size:12px;
        font-weight:600;
        color:#64748b;
    }

    .analysis-focus-title {
        margin-top:2px;
        font-size:16px;
        font-weight:700;
        color:var(--navy);
        word-break:break-word;
    }

    .analysis-focus-close-icon {
        width:34px;
        height:34px;
        flex:0 0 34px;
        border:1px solid #d6dee8;
        border-radius:7px;
        background:#ffffff;
        color:#64748b;
        font-size:22px;
        line-height:1;
        cursor:pointer;
        padding:0;
    }

    .analysis-focus-context {
        background:#f8fafc;
        border:1px solid #e2e8f0;
        border-radius:8px;
        padding:12px;
        margin-bottom:12px;
    }

    .analysis-focus-label {
        font-size:12px;
        font-weight:600;
        color:#64748b;
        margin-bottom:4px;
    }

    .focus-sentence {
        font-size:14px;
        line-height:1.7;
        color:#334155;
        word-break:break-word;
    }

    .analysis-focus-grid {
        display:grid;
        grid-template-columns:1fr 1fr;
        gap:10px;
        margin-bottom:12px;
    }

    .analysis-focus-box {
        border:1px solid #e2e8f0;
        border-radius:8px;
        padding:12px;
        min-width:0;
    }

    .analysis-focus-found {
        font-size:16px;
        font-weight:700;
        color:#c2410c;
        word-break:break-word;
    }

    .analysis-focus-suggested {
        font-size:16px;
        font-weight:700;
        color:#15803d;
        word-break:break-word;
    }

    .analysis-focus-meta {
        border-top:1px solid #e2e8f0;
        padding-top:10px;
        display:flex;
        flex-direction:column;
        gap:7px;
        margin-bottom:14px;
    }

    .analysis-focus-meta-row {
        display:grid;
        grid-template-columns:90px 1fr;
        gap:8px;
        font-size:13px;
        line-height:1.5;
    }

    .analysis-focus-meta-label {
        color:#64748b;
        font-weight:600;
    }

    .analysis-focus-back {
        width:100%;
        min-height:42px;
        border:1px solid #cbd5e1;
        border-radius:7px;
        background:#ffffff;
        color:#334155;
        font-family:"Kanit","Noto Sans Thai",sans-serif;
        font-weight:600;
        cursor:pointer;
    }

    .analysis-focus-back:hover,
    .analysis-focus-close-icon:hover {
        background:#f8fafc;
    }

    @media (max-width:560px) {
        .analysis-focus-grid {
            grid-template-columns:1fr;
        }

        .analysis-focus-meta-row {
            grid-template-columns:1fr;
            gap:2px;
        }

        .analysis-focus-grid {
            grid-template-columns:1fr;
        }
    }

    .result-title { color:var(--navy); font-size:1.22rem; font-weight:800; }
    .result-sub { color:#4b6073; font-size:.78rem; }

    div[data-testid="stExpander"] {
        border-radius:8px;
        border:1px solid #cbd5e1;
        margin-bottom:.7rem;
        background:#fff;
        overflow:hidden;
    }

    .stTabs [data-baseweb="tab-list"] {
        gap:.25rem;
        border-bottom:1px solid #cbd5e1;
        flex-wrap:wrap;
    }

    .stTabs [data-baseweb="tab"] { color:#40556a; font-weight:700; }
    .stTabs [aria-selected="true"] { color:var(--navy) !important; }

    /* Primary */
    .stButton > button[kind="primary"],
    .stDownloadButton > button[kind="primary"] {
        background:var(--navy) !important;
        color:#fff !important;
        border:1px solid var(--navy) !important;
    }

    /* Secondary */
    .stButton > button:not([kind="primary"]),
    .stDownloadButton > button:not([kind="primary"]) {
        background:#fff !important;
        color:#334155 !important;
        border:1px solid #cbd5e1 !important;
    }

    .stButton > button,
    .stDownloadButton > button {
        border-radius:7px;
        min-height:2.5rem;
        font-weight:700;
        width:100%;
    }

    /* =====================================================
       File uploader — Simple / Clear UX
       ===================================================== */

    .upload-section-title {
        display:flex;
        align-items:center;
        gap:9px;
        margin:.25rem 0 .55rem;
        color:#0f2742;
        font-size:.98rem;
        font-weight:800;
    }

    .upload-step {
        width:26px;
        height:26px;
        display:inline-flex;
        align-items:center;
        justify-content:center;
        border-radius:50%;
        background:#0f3d68;
        color:#fff;
        font-size:.78rem;
        font-weight:800;
        flex:0 0 auto;
    }

    .upload-helper {
        margin:-.25rem 0 .65rem 35px;
        color:#64748b;
        font-size:.78rem;
    }

    .stFileUploader {
        margin-bottom:.35rem;
    }

    .stFileUploader [data-testid="stFileUploaderDropzone"] {
        min-height:150px !important;
        padding:1.15rem 1.25rem !important;
        border:1.5px dashed #9bb3c9 !important;
        background:#ffffff !important;
        border-radius:12px !important;
        display:flex !important;
        flex-direction:column !important;
        align-items:center !important;
        justify-content:center !important;
        gap:.65rem !important;
        transition:border-color .15s ease, background .15s ease;
    }

    .stFileUploader [data-testid="stFileUploaderDropzone"]:hover {
        border-color:#477ca8 !important;
        background:#f8fbfe !important;
    }

    /* ซ่อนข้อความเดิมของ Streamlit ทั้งหมด เพื่อไม่ให้ข้อความซ้อนกัน */
    .stFileUploader [data-testid="stFileUploaderDropzoneInstructions"] > div,
    .stFileUploader [data-testid="stFileUploaderDropzoneInstructions"] small {
        display:none !important;
    }

    .stFileUploader [data-testid="stFileUploaderDropzoneInstructions"] {
        width:100% !important;
        text-align:center !important;
        display:block !important;
        padding:0 !important;
        margin:0 !important;
    }

    .stFileUploader [data-testid="stFileUploaderDropzoneInstructions"]::before {
        content:"ลากไฟล์มาวางที่นี่";
        display:block;
        color:#16324f;
        font-weight:800;
        font-size:1rem;
        line-height:1.45;
        margin-bottom:.28rem;
    }

    .stFileUploader [data-testid="stFileUploaderDropzoneInstructions"]::after {
        content:"รองรับไฟล์ PDF และ DOCX";
        display:block;
        color:#6b7f92;
        font-size:.78rem;
        font-weight:500;
        line-height:1.4;
    }

    .stFileUploader [data-testid="baseButton-secondary"] {
        min-width:112px !important;
        min-height:38px !important;
        padding:.45rem .9rem !important;
        border:1px solid #c5d1dc !important;
        border-radius:8px !important;
        background:#fff !important;
        color:#18324c !important;
        box-shadow:none !important;
        font-size:0 !important;
        font-weight:700 !important;
    }

    .stFileUploader [data-testid="baseButton-secondary"]::after {
        content:"เลือกไฟล์";
        font-size:.85rem;
        font-family:"Kanit","Noto Sans Thai",sans-serif;
    }

    .stFileUploader [data-testid="baseButton-secondary"]:hover {
        border-color:#6f91af !important;
        background:#f7fafc !important;
    }

    /* รายชื่อไฟล์หลังเลือกแล้ว */
    .stFileUploader [data-testid="stFileUploaderFile"] {
        border:1px solid #dbe4ec !important;
        border-radius:9px !important;
        background:#fff !important;
    }

    /* ให้ preview component ดูเป็นพื้นที่งานที่เลื่อนได้ในตัว */
    iframe[title="streamlit_components.v1.components.html"] {
        width:100% !important;
        border:1px solid var(--line) !important;
        border-radius:8px !important;
        background:#eef2f6 !important;
    }

    /* Responsive */
    @media (max-width: 900px) {
        .block-container {
            padding-left:.65rem;
            padding-right:.65rem;
            padding-top:.5rem;
        }

        .topbar {
            margin:-.5rem -.65rem .7rem -.65rem;
            padding:.65rem .75rem;
        }

        .brand { font-size:1.05rem; }
        .nav { display:none; }

        .intro-card,
        .ready-card {
            margin:1rem auto .75rem;
            padding:1rem;
        }

        .intro-title,
        .ready-title {
            font-size:1.2rem;
        }

        .doc-toolbar {
            font-size:.76rem;
            padding:.5rem .6rem;
        }

        .result-title {
            font-size:1.08rem;
        }

        .stButton > button,
        .stDownloadButton > button {
            min-height:44px;
            font-size:.92rem;
        }

        [data-testid="stFileUploaderDropzone"] {
            min-height:132px !important;
            padding:.9rem !important;
        }

        .upload-helper {
            margin-left:0;
        }
    }

    @media (max-width: 560px) {
        .brand { font-size:1rem; }
        .block-container { padding-bottom:1rem; }

        .stTabs [data-baseweb="tab"] {
            font-size:.82rem;
            padding-left:.45rem;
            padding-right:.45rem;
        }

        div[data-testid="stExpander"] summary {
            font-size:.86rem !important;
        }

        iframe[title="streamlit_components.v1.components.html"] {
            min-height:540px !important;
        }
    }
    
    /* =====================================================
       Accessibility — Larger, easier-to-read typography
       รองรับผู้มีสายตาเลือนราง / ผู้สูงอายุ
       ===================================================== */

    html, body, [class*="css"], .stApp {
        font-size:18px !important;
    }

    /* เนื้อหาทั่วไป */
    .stMarkdown,
    .stMarkdown p,
    .stMarkdown li,
    .stCaption,
    label,
    [data-testid="stWidgetLabel"] p {
        font-size:1rem !important;
        line-height:1.65 !important;
    }

    /* หัวข้อ */
    h1 {
        font-size:2rem !important;
        line-height:1.3 !important;
    }

    h2 {
        font-size:1.65rem !important;
        line-height:1.35 !important;
    }

    h3 {
        font-size:1.35rem !important;
        line-height:1.4 !important;
    }

    /* Header / Navigation */
    .top-title,
    .side-brand {
        font-size:1.15rem !important;
    }

    .top-nav {
        font-size:1rem !important;
    }

    /* Upload */
    .upload-section-title {
        font-size:1.15rem !important;
    }

    .upload-step {
        width:32px !important;
        height:32px !important;
        font-size:.95rem !important;
    }

    .upload-helper {
        font-size:.95rem !important;
        line-height:1.55 !important;
    }

    .stFileUploader [data-testid="stFileUploaderDropzoneInstructions"]::before {
        font-size:1.15rem !important;
        line-height:1.55 !important;
    }

    .stFileUploader [data-testid="stFileUploaderDropzoneInstructions"]::after {
        font-size:.95rem !important;
        line-height:1.55 !important;
    }

    .stFileUploader [data-testid="baseButton-secondary"] {
        min-width:132px !important;
        min-height:46px !important;
    }

    .stFileUploader [data-testid="baseButton-secondary"]::after {
        font-size:1rem !important;
    }

    /* ปุ่มทั้งหมด */
    .stButton button,
    .stDownloadButton button {
        min-height:46px !important;
        font-size:1rem !important;
        font-weight:600 !important;
        padding:.6rem 1rem !important;
    }

    /* Selectbox / Input */
    [data-baseweb="select"] > div,
    .stTextInput input,
    .stNumberInput input {
        min-height:46px !important;
        font-size:1rem !important;
    }

    [data-baseweb="select"] span,
    [role="option"] {
        font-size:1rem !important;
        line-height:1.5 !important;
    }

    /* Tabs */
    button[data-baseweb="tab"] {
        min-height:48px !important;
    }

    button[data-baseweb="tab"] p {
        font-size:1rem !important;
        font-weight:700 !important;
    }

    /* Result cards */
    .result-title {
        font-size:1.35rem !important;
        line-height:1.4 !important;
    }

    .result-sub,
    .intro-sub {
        font-size:.95rem !important;
        line-height:1.6 !important;
    }

    /* Expander */
    [data-testid="stExpander"] summary p {
        font-size:1rem !important;
        line-height:1.55 !important;
    }

    /* Metric */
    [data-testid="stMetricValue"] {
        font-size:1.75rem !important;
    }

    [data-testid="stMetricLabel"] {
        font-size:.95rem !important;
    }

    /* Caption เดิมของ Streamlit เล็กเกินไป */
    [data-testid="stCaptionContainer"] p {
        font-size:.9rem !important;
        line-height:1.55 !important;
    }

    /* Focus analysis */
    .focus-location {
        font-size:.95rem !important;
    }

    .focus-word,
    .focus-suggested {
        font-size:1.25rem !important;
    }

    .focus-detail,
    .focus-sentence,
    .focus-confidence,
    .focus-method {
        font-size:1rem !important;
        line-height:1.6 !important;
    }

    /* จอเล็กยังคงอ่านง่าย แต่ไม่ใหญ่จน UI แตก */
    @media (max-width: 768px) {
        html, body, [class*="css"], .stApp {
            font-size:17px !important;
        }

        .upload-section-title {
            font-size:1.1rem !important;
        }

        .stFileUploader [data-testid="stFileUploaderDropzoneInstructions"]::before {
            font-size:1.08rem !important;
        }

        .stButton button,
        .stDownloadButton button {
            min-height:48px !important;
        }
    }


    /* =====================================================
       Sticky Preview
       ล็อก Preview ฝั่งซ้ายไว้ ขณะเลื่อนดูผลวิเคราะห์ฝั่งขวา
       ===================================================== */

    div[data-testid="stColumn"]:has(#sticky-preview-anchor) {
        position: sticky !important;
        top: 10px !important;
        align-self: flex-start !important;
        z-index: 20 !important;
        height: fit-content !important;
    }

    /* ให้ Preview ไม่ถูกบังจาก element อื่น */
    div[data-testid="stColumn"]:has(#sticky-preview-anchor)
    iframe[title="streamlit_components.v1.components.html"] {
        position: relative;
        z-index: 1;
    }

    /* บน Tablet / Mobile ให้กลับเป็น layout ปกติ
       เพราะ Streamlit จะเรียงคอลัมน์บนลงล่าง */
    @media (max-width: 900px) {
        div[data-testid="stColumn"]:has(#sticky-preview-anchor) {
            position: static !important;
            top: auto !important;
            z-index: auto !important;
        }
    }


    /* Scope control: วาง “แสดงทั้งเอกสาร” ข้างตัวเลือกหน้า */
    div[data-testid="stHorizontalBlock"]:has(
        [data-testid="stSelectbox"]
    ) [data-testid="stToggle"] {
        min-height:46px;
        display:flex;
        align-items:center;
        justify-content:flex-end;
        padding:.2rem 0;
    }

    div[data-testid="stHorizontalBlock"]:has(
        [data-testid="stSelectbox"]
    ) [data-testid="stToggle"] label {
        font-size:1rem !important;
        font-weight:600 !important;
        color:#243b53 !important;
        white-space:nowrap;
    }

    @media (max-width: 900px) {
        div[data-testid="stHorizontalBlock"]:has(
            [data-testid="stSelectbox"]
        ) [data-testid="stToggle"] {
            justify-content:flex-start;
        }
    }

</style>
    """,
    unsafe_allow_html=True
)

# =========================================================
# ส่วนหัว
# =========================================================

st.markdown(
    """
    <div class="topbar">
      <div class="brand-wrap">
        <div class="brand">A̲&nbsp; ระบบตรวจคำไทย</div>
        <div class="nav">
          <span class="active">แดชบอร์ด</span>
        </div>
      </div>
    </div>
    """,
    unsafe_allow_html=True
)

# =========================================================
# อัปโหลดไฟล์
# =========================================================

st.markdown(
    """
    <div class="upload-section-title">
        <span class="upload-step">1</span>
        <span>อัปโหลดเอกสาร</span>
    </div>
    <div class="upload-helper">
        ลากไฟล์ลงในช่องด้านล่าง หรือกด “เลือกไฟล์”
    </div>
    """,
    unsafe_allow_html=True
)

uploaded_file = st.file_uploader(
    "อัปโหลดเอกสาร",
    type=["pdf", "docx"],
    label_visibility="collapsed",
    key="main_uploader"
)

if uploaded_file is None:
    with st.sidebar:
        st.markdown('<div class="side-brand">📄 ข้อมูลเอกสาร</div>', unsafe_allow_html=True)
        if JSON_DICTIONARY_OK:
            st.success(JSON_DICTIONARY_STATUS, icon="✅")
        else:
            st.warning(JSON_DICTIONARY_STATUS, icon="⚠️")
        st.markdown(
            '<div class="tiny">สถานะ</div>'
            '<div style="font-weight:700;color:#334155;margin-bottom:.8rem;">รออัปโหลดไฟล์</div>',
            unsafe_allow_html=True
        )

    st.markdown(
        """
        <div class="intro-card" style="padding:.9rem 1rem;">
          <div class="intro-sub" style="margin:0;">
            เมื่อเลือกไฟล์แล้ว ระบบจะแสดงข้อมูลเอกสารและปุ่มเริ่มตรวจสอบ
          </div>
        </div>
        """,
        unsafe_allow_html=True
    )
    st.stop()

file_bytes = uploaded_file.getvalue()
file_type = uploaded_file.name.rsplit(".", 1)[-1].lower()
current_key = f"{uploaded_file.name}_{len(file_bytes)}"

if st.session_state.file_key != current_key:
    st.session_state.file_key = current_key
    st.session_state.analyzed_data = None

# =========================================================
# แถบด้านซ้ายเมื่อมีไฟล์
# =========================================================

with st.sidebar:
    st.markdown('<div class="side-brand">A̲&nbsp; ระบบตรวจคำไทย</div>', unsafe_allow_html=True)
    if JSON_DICTIONARY_OK:
        st.success(JSON_DICTIONARY_STATUS, icon="✅")
    else:
        st.warning(JSON_DICTIONARY_STATUS, icon="⚠️")
    st.markdown(
        f"""
        <div class="file-card">
          <div class="file-name">📄 {html.escape(uploaded_file.name)}</div>
          <div class="tiny">ขนาด: {len(file_bytes)/1024/1024:.2f} เมกะไบต์</div>
          <div class="tiny">ชนิดไฟล์: {file_type.upper()}</div>
        </div>
        """,
        unsafe_allow_html=True
    )

    if st.session_state.analyzed_data is None:
        st.markdown('<div class="side-item active">▣ เอกสารที่เลือก</div>', unsafe_allow_html=True)
        st.markdown('<div class="side-item">○ รอผลการตรวจสอบ</div>', unsafe_allow_html=True)
    else:
        ae = st.session_state.analyzed_data
        st.markdown('<div class="menu-title">ผลการตรวจสอบ</div>', unsafe_allow_html=True)
        st.markdown(
            f'<div class="side-item active">▣ ผลทั้งหมด '
            f'<span class="count-badge">{len(ae["all_errors"]) + len(ae["all_uncertain"])}</span></div>',
            unsafe_allow_html=True
        )
        st.markdown(
            f'<div class="side-item">ⓘ คำผิด '
            f'<span class="count-badge">{len(ae["all_errors"])}</span></div>',
            unsafe_allow_html=True
        )
        st.markdown(
            f'<div class="side-item">? ไม่แน่ใจ '
            f'<span class="count-badge warn">{len(ae["all_uncertain"])}</span></div>',
            unsafe_allow_html=True
        )

# =========================================================
# เริ่มตรวจสอบ
# =========================================================

if st.session_state.analyzed_data is None:
    st.markdown(
        f"""
        <div class="ready-card">
          <div class="ready-title">เอกสารพร้อมตรวจสอบ</div>
          <div class="ready-sub">
            {html.escape(uploaded_file.name)} · {len(file_bytes)/1024:.0f} กิโลไบต์
          </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    c1, c2, c3 = st.columns([1, 1.2, 1])
    with c2:
        start_analysis = st.button(
            "เริ่มตรวจสอบเอกสาร",
            type="primary",
            use_container_width=True
        )

    if not start_analysis:
        st.stop()

    with st.spinner("กำลังตรวจสอบเอกสาร กรุณารอสักครู่..."):
        if file_type == "pdf":
            pages_data, all_errors, all_uncertain = process_full_pdf(file_bytes)
        else:
            pages_data, all_errors, all_uncertain = process_full_docx(file_bytes)

        st.session_state.analyzed_data = {
            "pages_data": pages_data,
            "all_errors": all_errors,
            "all_uncertain": all_uncertain
        }

    st.rerun()

# =========================================================
# หน้าผลการตรวจสอบ
# =========================================================

analyzed = st.session_state.analyzed_data
pages_data = analyzed["pages_data"]
all_errors = analyzed["all_errors"]
all_uncertain = analyzed["all_uncertain"]

exportable_error_count = sum(
    1
    for item in all_errors
    if item.get("bboxes")
)

exportable_uncertain_count = sum(
    1
    for item in all_uncertain
    if item.get("bboxes")
)

if not pages_data:
    st.warning("ไม่พบข้อความที่สามารถตรวจสอบได้ในเอกสารนี้")
    st.stop()

page_label_map = {}
page_labels = []

for p in pages_data:
    error_count = len(p.get("errors", []))
    uncertain_count = len(p.get("uncertain", []))
    # ข้อความของ Dropdown ใช้ข้อความปกติ
    # แล้วใช้ JavaScript ด้านล่าง Highlight เฉพาะ "ตัวเลขจำนวน"
    label = (
        f"{p['location']} · ผิด {error_count} คำ"
        f" · ไม่แน่ใจ {uncertain_count} คำ"
    )
    page_labels.append(label)
    page_label_map[label] = p

page_control_left, page_control_right = st.columns(
    [4.6, 1.4],
    gap="medium"
)

with page_control_left:
    selected_label = st.selectbox(
        "เลือกหน้าเอกสาร",
        page_labels,
        label_visibility="collapsed",
        key="page_selector"
    )

with page_control_right:
    show_all_document = st.toggle(
        "แสดงทั้งเอกสาร",
        value=False,
        key="show_all_document",
        on_change=reset_preview_focus_state,
    )

selected_page = page_label_map[selected_label]

# ส่งค่าที่เลือกจริงจาก Python เข้า JavaScript โดยตรง
# เพื่อให้สถานะปิดของ Dropdown ไม่ต้องเดาค่าจาก DOM ของ Streamlit
_selected_label_js = json.dumps(selected_label, ensure_ascii=False)

# =========================================================
# HIGHLIGHT COUNTS INSIDE PAGE DROPDOWN
# =========================================================
# 1) ตอนเปิด dropdown: Highlight ตัวเลขในทุก option
# 2) ตอนปิด dropdown (static): สร้าง overlay บนค่าที่เลือกอยู่
#    เพราะ BaseWeb บางเวอร์ชันเก็บ selected value ใน input/value
#    จึงไม่สามารถ Highlight บางส่วนด้วย Text Node อย่างเดียวได้
components.html(
    fr"""
    <script>
    (function () {{
        const parentDoc = window.parent.document;
        const HIGHLIGHT_CLASS = "spell-count-highlight";
        const STATIC_OVERLAY_CLASS = "spell-static-count-overlay";
        const SELECTED_LABEL = {_selected_label_js};
        const pattern = /(ผิด\\s*)(\\d+)(\\s*คำ\\s*·\\s*ไม่แน่ใจ\\s*)(\\d+)(\\s*คำ)/;

        function makeHighlight(value) {{
            const span = parentDoc.createElement("span");
            span.className = HIGHLIGHT_CLASS;
            span.textContent = value;
            span.style.background = "#fff200";
            span.style.color = "#111827";
            span.style.fontWeight = "800";
            span.style.padding = "0 3px";
            span.style.margin = "0 1px";
            span.style.borderRadius = "2px";
            span.style.lineHeight = "1.35";
            span.style.display = "inline-block";
            return span;
        }}

        function appendHighlightedLabel(container, fullText) {{
            const match = (fullText || "").match(pattern);
            if (!match) return false;

            const before = fullText.slice(0, match.index);
            const after = fullText.slice(match.index + match[0].length);

            container.appendChild(
                parentDoc.createTextNode(before + match[1])
            );
            container.appendChild(makeHighlight(match[2]));
            container.appendChild(
                parentDoc.createTextNode(match[3])
            );
            container.appendChild(makeHighlight(match[4]));
            container.appendChild(
                parentDoc.createTextNode(match[5] + after)
            );
            return true;
        }}

        function highlightTextNode(textNode) {{
            if (!textNode || !textNode.parentElement) return;
            if (textNode.parentElement.closest("." + HIGHLIGHT_CLASS)) return;
            if (textNode.parentElement.closest("." + STATIC_OVERLAY_CLASS)) return;

            const text = textNode.nodeValue || "";
            if (!text.includes("ผิด") || !text.includes("ไม่แน่ใจ")) return;

            const match = text.match(pattern);
            if (!match) return;

            const fragment = parentDoc.createDocumentFragment();
            const before = text.slice(0, match.index);
            const after = text.slice(match.index + match[0].length);

            fragment.appendChild(parentDoc.createTextNode(before + match[1]));
            fragment.appendChild(makeHighlight(match[2]));
            fragment.appendChild(parentDoc.createTextNode(match[3]));
            fragment.appendChild(makeHighlight(match[4]));
            fragment.appendChild(parentDoc.createTextNode(match[5] + after));

            textNode.parentNode.replaceChild(fragment, textNode);
        }}

        function highlightInside(root) {{
            if (!root) return;
            const walker = parentDoc.createTreeWalker(root, NodeFilter.SHOW_TEXT);
            const textNodes = [];
            let node;
            while ((node = walker.nextNode())) textNodes.push(node);
            textNodes.forEach(highlightTextNode);
        }}

        function getSelectedLabel(box) {{
            // ค่าที่มาจาก Python เป็น source of truth
            if (SELECTED_LABEL && pattern.test(SELECTED_LABEL)) {{
                return SELECTED_LABEL;
            }}

            const input = box.querySelector('input');
            if (input && input.value && pattern.test(input.value)) {{
                return input.value;
            }}

            const text = (box.textContent || "").replace(/\\s+/g, " ").trim();
            const pageMatch = text.match(/หน้า\\s*\\d+\\s*·\\s*ผิด\\s*\\d+\\s*คำ\\s*·\\s*ไม่แน่ใจ\\s*\\d+\\s*คำ/);
            return pageMatch ? pageMatch[0] : "";
        }}

        function decorateStaticSelect(box) {{
            if (!box) return;
            const label = getSelectedLabel(box);
            if (!label) return;

            const control =
                box.querySelector('[data-baseweb="select"]') ||
                box.querySelector('[role="combobox"]') ||
                box;

            let overlay = control.querySelector("." + STATIC_OVERLAY_CLASS);
            if (overlay && overlay.dataset.label === label) return;
            if (overlay) overlay.remove();

            const input = box.querySelector('input');

            const candidates = Array.from(control.querySelectorAll('div, span'))
                .filter(function (el) {{
                    if (el.closest("." + STATIC_OVERLAY_CLASS)) return false;
                    const t = (el.textContent || "").replace(/\s+/g, " ").trim();
                    return t === label;
                }})
                .sort(function (a, b) {{
                    return a.querySelectorAll('*').length - b.querySelectorAll('*').length;
                }});

            const computed = window.parent.getComputedStyle(control);
            control.style.position = "relative";

            overlay = parentDoc.createElement("div");
            overlay.className = STATIC_OVERLAY_CLASS;
            overlay.dataset.label = label;
            overlay.style.position = "absolute";
            overlay.style.left = "12px";
            overlay.style.right = "42px";
            overlay.style.top = "50%";
            overlay.style.transform = "translateY(-50%)";
            overlay.style.pointerEvents = "none";
            overlay.style.zIndex = "5";
            overlay.style.whiteSpace = "nowrap";
            overlay.style.overflow = "hidden";
            overlay.style.textOverflow = "ellipsis";
            overlay.style.font = computed.font;
            overlay.style.color = "#111827";
            overlay.style.lineHeight = "1.35";

            overlay.style.background =
                computed.backgroundColor &&
                computed.backgroundColor !== "rgba(0, 0, 0, 0)"
                    ? computed.backgroundColor
                    : "#f3f4f6";

            // ไม่ซ่อน native text อีกต่อไป
            // overlay ที่มีพื้นหลังจะทับข้อความเดิมอย่างปลอดภัย
            if (appendHighlightedLabel(overlay, label)) {{
                control.appendChild(overlay);
            }}
        }}

        function scan() {{
            try {{
                const selectboxes = Array.from(
                    parentDoc.querySelectorAll('[data-testid="stSelectbox"]')
                );

                let targetSelect = null;

                if (window.frameElement && selectboxes.length) {{
                    const frameTop =
                        window.frameElement.getBoundingClientRect().top;

                    let bestDistance = Number.POSITIVE_INFINITY;

                    selectboxes.forEach(function (box) {{
                        const rect = box.getBoundingClientRect();
                        const distance = frameTop - rect.bottom;

                        if (distance >= -10 && distance < bestDistance) {{
                            bestDistance = distance;
                            targetSelect = box;
                        }}
                    }});
                }}

                if (!targetSelect && selectboxes.length) {{
                    targetSelect = selectboxes[0];
                }}

                if (
                    targetSelect &&
                    SELECTED_LABEL &&
                    pattern.test(SELECTED_LABEL)
                ) {{
                    decorateStaticSelect(targetSelect);
                }}

                parentDoc
                    .querySelectorAll('[role="option"]')
                    .forEach(function (option) {{
                        const text = option.textContent || "";
                        if (text.includes("ผิด") && text.includes("ไม่แน่ใจ")) {{
                            highlightInside(option);
                        }}
                    }});

                parentDoc
                    .querySelectorAll('[role="listbox"]')
                    .forEach(highlightInside);
            }} catch (err) {{
                console.warn("ไม่สามารถ Highlight จำนวนใน Dropdown ได้", err);
            }}
        }}

        scan();

        const observer = new MutationObserver(function () {{
            window.requestAnimationFrame(scan);
        }});

        observer.observe(parentDoc.body, {{
            childList: true,
            subtree: true
        }});

        setTimeout(scan, 50);
        setTimeout(scan, 150);
        setTimeout(scan, 300);
        setTimeout(scan, 600);
        setTimeout(scan, 1000);
    }})();
    </script>
    """,
    height=0,
    scrolling=False
)

left, right = st.columns([2.65, 1.35], gap="large")

with left:
    # Anchor สำหรับ Sticky Preview:
    # CSS จะล็อกเฉพาะคอลัมน์นี้ไว้เมื่อผู้ใช้เลื่อนลงดูผลวิเคราะห์
    st.markdown(
        '<div id="sticky-preview-anchor" aria-hidden="true"></div>',
        unsafe_allow_html=True
    )

    st.markdown(
        f"""
        <div class="doc-toolbar">
          <span class="page-chip">{html.escape(selected_page["location"])}</span>
          <span>ใช้ปุ่ม − / + ในตัวอย่างเพื่อซูม</span>
        </div>
        """,
        unsafe_allow_html=True
    )

    if selected_page.get("pdf_page") is not None:
        preview_items = selected_page["errors"] + selected_page["uncertain"]
        preview_html = build_interactive_pdf_preview(
            selected_page["pdf_page"],
            preview_items,
            selected_page["page_num"]
        )

        # ทำให้ HTML signature เปลี่ยนเมื่อเปิด/ปิด “แสดงทั้งเอกสาร”
        # เพื่อให้ Preview component refresh และเคลียร์ Focus Mode เดิม
        preview_html += (
            f"<!-- preview-refresh:"
            f"{st.session_state.preview_refresh_token} -->"
        )
        # ใช้ components.html แทน st.markdown สำหรับ HTML ที่มีรูป Base64 ขนาดใหญ่
        # เพื่อป้องกัน Streamlit แสดง source HTML ออกมาทั้งก้อน
        preview_height = 900
        try:
            page_rect = selected_page["pdf_page"].rect
            if page_rect.width:
                preview_height = int(
                    min(
                        max((page_rect.height / page_rect.width) * 720 + 120, 700),
                        1400
                    )
                )
        except Exception:
            pass

        # ความสูงเริ่มต้นสำหรับ desktop/mobile
        preview_height = 820

        try:
            components.html(
                preview_html,
                height=preview_height,
                scrolling=True,
            )
        except Exception as preview_error:
            st.warning("ไม่สามารถแสดงตัวอย่างแบบโต้ตอบได้ จึงแสดงเป็นภาพแทน")
            try:
                fallback_img = draw_red_boxes_on_image(
                    selected_page["image"],
                    selected_page["pdf_page"],
                    preview_items,
                )
                st.image(fallback_img, use_container_width=True)
            except Exception:
                st.error(f"เกิดข้อผิดพลาดในการแสดงตัวอย่าง: {preview_error}")
    else:
        st.text_area(
            "เนื้อหาเอกสาร",
            value=selected_page["text"],
            height=620,
            disabled=True,
            label_visibility="collapsed"
        )

with right:
    display_errors = (
        all_errors
        if show_all_document
        else selected_page["errors"]
    )

    display_uncertain = (
        all_uncertain
        if show_all_document
        else selected_page["uncertain"]
    )

    if show_all_document:
        st.caption(
            f"กำลังแสดงผลทั้งเอกสาร · "
            f"คำผิด {len(display_errors)} รายการ · "
            f"ยังไม่แน่ใจ {len(display_uncertain)} รายการ"
        )

    scope_text = (
        "ทั้งเอกสาร"
        if show_all_document
        else selected_page["location"]
    )

    st.markdown(
        f"""
        <div class="result-head">
          <div class="result-title">ผลการวิเคราะห์</div>
          <div class="result-sub">
            {html.escape(scope_text)} · {html.escape(uploaded_file.name)}
          </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    # =====================================================
    # EXPORT PDF พร้อมกรอบผลตรวจ
    # =====================================================

    if file_type == "pdf":
        try:
            annotated_pdf_bytes = export_annotated_pdf(
                file_bytes,
                all_errors,
                all_uncertain,
            )

            original_stem = (
                uploaded_file.name.rsplit(".", 1)[0]
            )

            export_name = (
                f"{original_stem}_ตรวจคำแล้ว.pdf"
            )

            st.download_button(
                label="⬇️ Export PDF พร้อมกรอบผลตรวจ",
                data=annotated_pdf_bytes,
                file_name=export_name,
                mime="application/pdf",
                use_container_width=True,
                type="primary",
                key="download_annotated_pdf",
            )

            st.caption(
                "🔴 คำผิด "
                f"{exportable_error_count} รายการ"
                " · 🔵 ยังไม่แน่ใจ "
                f"{exportable_uncertain_count} รายการ"
            )

        except Exception as export_error:
            st.warning(
                "ไม่สามารถสร้าง PDF สำหรับ Export ได้: "
                f"{export_error}"
            )

    else:
        st.info(
            "การ Export PDF พร้อมกรอบตำแหน่งคำ "
            "รองรับไฟล์ต้นฉบับ PDF ก่อน เนื่องจาก DOCX "
            "ยังไม่มีพิกัดคำระดับหน้าเอกสาร"
        )

    # =====================================================
    # EXPORT REPORT — CSV / EXCEL / JSON
    # =====================================================

    report_rows = build_analysis_report_rows(
        all_errors,
        all_uncertain,
    )

    report_stem = uploaded_file.name.rsplit(".", 1)[0]

    with st.expander(
        "📄 Export รายงานผลตรวจ",
        expanded=False,
    ):
        st.caption(
            f"รายงานทั้งเอกสาร · "
            f"คำผิด {len(all_errors)} รายการ · "
            f"ยังไม่แน่ใจ {len(all_uncertain)} รายการ"
        )

        export_col_csv, export_col_excel, export_col_json = st.columns(
            3,
            gap="small",
        )

        with export_col_csv:
            st.download_button(
                label="CSV",
                data=export_report_csv(report_rows),
                file_name=f"{report_stem}_รายงานผลตรวจ.csv",
                mime="text/csv",
                use_container_width=True,
                key="download_report_csv",
            )

        with export_col_excel:
            if OPENPYXL_AVAILABLE:
                st.download_button(
                    label="Excel",
                    data=export_report_excel(report_rows),
                    file_name=f"{report_stem}_รายงานผลตรวจ.xlsx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "spreadsheetml.sheet"
                    ),
                    use_container_width=True,
                    key="download_report_excel",
                )
            else:
                st.button(
                    "Excel",
                    disabled=True,
                    use_container_width=True,
                    key="download_report_excel_disabled",
                )
                st.caption("ต้องติดตั้ง openpyxl")

        with export_col_json:
            st.download_button(
                label="JSON",
                data=export_report_json(report_rows),
                file_name=f"{report_stem}_รายงานผลตรวจ.json",
                mime="application/json",
                use_container_width=True,
                key="download_report_json",
            )

        st.caption(
            "ข้อมูลในรายงาน: สถานะ · หน้า · ตำแหน่ง · คำที่พบ · "
            "คำแนะนำ · รายละเอียด · ความมั่นใจ · วิธีตรวจพบ · บริบท"
        )

    # IMPORTANT: dedent เพื่อไม่ให้ Markdown มอง HTML ที่เยื้องเป็น code block
    focus_panel_html = textwrap.dedent("""
    <div id="analysis-focus-panel" class="analysis-focus-panel">
        <div class="analysis-focus-card">
            <div class="analysis-focus-header">
                <div>
                    <div class="analysis-focus-kicker">ผลวิเคราะห์จากตำแหน่งที่เลือก</div>
                    <div class="analysis-focus-title focus-location"></div>
                </div>
                <button
                    type="button"
                    id="analysis-focus-close"
                    class="analysis-focus-close-icon"
                    aria-label="ปิดผลวิเคราะห์"
                >×</button>
            </div>

            <div class="analysis-focus-context">
                <div class="analysis-focus-label">ข้อความที่พบ</div>
                <div class="focus-sentence"></div>
            </div>

            <div class="analysis-focus-grid">
                <div class="analysis-focus-box analysis-focus-found-box">
                    <div class="analysis-focus-label">คำที่พบ</div>
                    <div class="analysis-focus-found focus-word"></div>
                </div>
                <div class="analysis-focus-box analysis-focus-suggested-box">
                    <div class="analysis-focus-label">คำแนะนำ</div>
                    <div class="analysis-focus-suggested focus-suggested"></div>
                </div>
            </div>

            <div class="analysis-focus-meta">
                <div class="analysis-focus-meta-row">
                    <span class="analysis-focus-meta-label">ประเภท</span>
                    <span class="focus-detail"></span>
                </div>
                <div class="analysis-focus-meta-row">
                    <span class="analysis-focus-meta-label">ความมั่นใจ</span>
                    <span class="focus-confidence"></span>
                </div>
                <div class="analysis-focus-meta-row">
                    <span class="analysis-focus-meta-label">วิธีตรวจพบ</span>
                    <span class="focus-method"></span>
                </div>
            </div>

            <button
                type="button"
                class="analysis-focus-back"
                id="analysis-focus-back"
            >← กลับไปดูรายการทั้งหมด</button>
        </div>
    </div>
    """).strip()

    # IMPORTANT:
    # st.markdown() ยังสามารถตีความ HTML หลายบรรทัดเป็น Markdown/code block ได้
    # ใช้ st.html() โดยตรงเพื่อ render HTML ใน DOM หลักของ Streamlit
    # ซึ่งยังทำให้ JavaScript จาก Preview iframe เข้าถึง #analysis-focus-panel ได้
    if hasattr(st, "html"):
        st.html(focus_panel_html)
    else:
        # Streamlit รุ่นเก่ามากที่ยังไม่มี st.html:
        # บีบ HTML เป็นบรรทัดเดียวเพื่อลดโอกาสถูก Markdown ตีความเป็น code block
        focus_panel_html_fallback = " ".join(
            line.strip()
            for line in focus_panel_html.splitlines()
            if line.strip()
        )
        st.markdown(
            focus_panel_html_fallback,
            unsafe_allow_html=True
        )

    tab_errors, tab_uncertain = st.tabs([
        f"คำผิด ({len(display_errors)})",
        f"ยังไม่แน่ใจ ({len(display_uncertain)})"
    ])

    with tab_errors:
        if display_errors:
            for err in display_errors:
                confidence = err.get("confidence", 0)
                error_id = err.get("error_id", "")

                st.markdown(
                    f"""
                    <div
                        id="{error_id}"
                        data-location="{html.escape(str(err.get('location', '')), quote=True)}"
                        data-word="{html.escape(str(err.get('word', '')), quote=True)}"
                        data-suggested="{html.escape(str(err.get('suggested', '')), quote=True)}"
                        data-detail="{html.escape(str(err.get('detail', 'คำสะกดไม่ถูกต้อง')), quote=True)}"
                        data-sentence="{html.escape(str(err.get('sentence', '')), quote=True)}"
                        data-confidence="{confidence * 100:.0f}%"
                        data-method="{html.escape(str(err.get('method', 'ตรวจการสะกดคำ')), quote=True)}"
                        style="height:1px;"
                    ></div>
                    """,
                    unsafe_allow_html=True
                )

                with st.expander(
                    f"{err['location']} · {err['word']} → {err['suggested']}",
                    expanded=True
                ):
                    st.markdown(
                        f"**ประเภท:** {err.get('detail', 'คำสะกดไม่ถูกต้อง')}  \n"
                        f"**ข้อความ:** {err['highlighted_sentence']}",
                        unsafe_allow_html=True
                    )

                    c_a, c_b = st.columns(2)

                    with c_a:
                        st.markdown(
                            f"**คำที่พบ**  \n"
                            f":orange[**{err['word']}**]"
                        )

                    with c_b:
                        st.markdown(
                            f"**คำแนะนำ**  \n"
                            f":green[**{err['suggested']}**]"
                        )

                    st.caption(
                        f"ความมั่นใจ {confidence * 100:.0f}% · "
                        f"วิธีตรวจพบ: {err.get('method', 'ตรวจการสะกดคำ')}"
                    )
        else:
            st.success("ไม่พบคำผิดในส่วนที่เลือก")

    with tab_uncertain:
        if display_uncertain:
            for unc in display_uncertain:
                uncertain_id = unc.get("error_id", "")
                uncertain_anchor_html = (
                    f'<div id="{html.escape(str(uncertain_id), quote=True)}" '
                    f'data-location="{html.escape(str(unc.get("location", "")), quote=True)}" '
                    f'data-word="{html.escape(str(unc.get("word", "")), quote=True)}" '
                    f'data-suggested="ควรตรวจสอบ" '
                    f'data-detail="{html.escape(str(unc.get("reason", "คำที่ควรตรวจสอบเพิ่มเติม")), quote=True)}" '
                    f'data-sentence="{html.escape(str(unc.get("sentence", "")), quote=True)}" '
                    f'data-confidence="ยังไม่ระบุ" '
                    f'data-method="ยังไม่แน่ใจ" '
                    f'style="height:1px;"></div>'
                )
                if hasattr(st, "html"):
                    st.html(uncertain_anchor_html)
                else:
                    st.markdown(uncertain_anchor_html, unsafe_allow_html=True)

                with st.expander(
                    f"{unc['location']} · {unc['word']}"
                ):
                    st.markdown(
                        f"**ข้อความ:** {unc['highlighted_sentence']}",
                        unsafe_allow_html=True
                    )
                    st.markdown(
                        f"**คำที่ควรตรวจสอบ:** `{unc['word']}`"
                    )
                    st.markdown(
                        f"**สาเหตุ:** {unc['reason']}"
                    )
        else:
            st.success("ไม่พบคำที่ควรตรวจสอบเพิ่มเติมในส่วนที่เลือก")

