import pytest
from xml.etree.ElementTree import Element
from docx_processor import DocxProcessor
import random
import string

import sys, os
sys.path.append(os.path.dirname(os.path.dirname(__file__)))
# ===== Helper giả lập paragraph / run =====
class FakeRun:
    def __init__(self, text="", bold=False, italic=False, underline=False, subscript=False, superscript=False):
        self.text = text
        self.bold = bold
        self.italic = italic
        self.underline = underline
        class Font:
            pass
        self.font = Font()
        self.font.subscript = subscript
        self.font.superscript = superscript

        # 🔧 Thêm mock element để tránh AttributeError
        self._element = type("E", (), {"xpath": lambda *a, **kw: []})()

class FakeParagraph:
    def __init__(self, text="", runs=None, alignment=0):
        self.text = text
        self.runs = runs or [FakeRun(text)]
        self.alignment = alignment
        self._p = type("P", (), {"xpath": lambda *a, **kw: []})()

# =================================================================

@pytest.fixture
def processor():
    return DocxProcessor()

# ---- Test escape_html ----
def test_escape_html(processor):
    text = '<tag> "quotes" & more'
    escaped = processor.escape_html(text)
    assert escaped == '&lt;tag&gt; &quot;quotes&quot; &amp; more'

# ---- Test wrap_style ----
@pytest.mark.parametrize("style,expected", [
    ((True, False, False, False, False), "<strong>abc</strong>"),
    ((False, True, False, False, False), "<i>abc</i>"),
    ((False, False, True, False, False), "<u>abc</u>"),
    ((False, False, False, True, False), "<sup>abc</sup>"),
    ((False, False, False, False, True), "<sub>abc</sub>")
])
def test_wrap_style(processor, style, expected):
    assert processor.wrap_style("abc", style) == expected

# ---- Test convert_paragraph_to_html ----
def test_convert_paragraph_to_html_basic(processor):
    p = FakeParagraph("Hello", [FakeRun("Hello", bold=True)])
    html = processor.convert_paragraph_to_html(p)
    assert "<strong>Hello</strong>" in html
    assert html.startswith("<p")

# ---- Test strip_html ----
def test_strip_html(processor):
    html = "<p><strong>Hello</strong></p>"
    assert processor.strip_html(html) == "Hello"

# ---- Test convert_content_to_html with list ----
def test_convert_content_to_html_flat(processor):
    p1 = FakeParagraph("A")
    p2 = FakeParagraph("B")
    html = processor.convert_content_to_html([p1, p2])
    assert "A" in html and "B" in html

# ---- Test dang_tn xử lý đúng đáp án ----
def test_dang_tn_correct_answer(processor):
    para_q = FakeParagraph("Câu 1: Nội dung câu hỏi")
    para_a = [
        FakeParagraph("A. Đáp án 1"),
        FakeParagraph("B. Đáp án 2"),
        FakeParagraph("C. Đáp án 3"),
        FakeParagraph("D. Đáp án 4")
    ]
    para_hdg = [FakeParagraph("Đáp án đúng là 2")]
    cau_sau_xu_ly = [[para_q] + para_a, [para_hdg]]
    xml = Element("question")

    processor.dang_tn(cau_sau_xu_ly, xml, audio=[])

    listanswers = xml.find("listanswers")
    assert listanswers is not None
    corrects = [a.find("isanswer").text for a in listanswers.findall("answer")]
    assert corrects.count("TRUE") == 1
    assert corrects[1] == "TRUE"  # Đáp án thứ 2 đúng

# ---- Test hdg_tn ----
def test_hdg_tn_generates_explain(processor):
    xml = Element("question")
    p = FakeParagraph("Đáp án đúng là 3. Giải thích: do ABC")
    processor.hdg_tn([p], xml)
    explain = xml.find("explainquestion")
    assert explain is not None
    assert "Đáp án đúng" in explain.text or "Giải thích" in explain.text

# ---- Test dang_dt (điền đáp án) ----
def test_dang_dt_extracts_answers(processor):
    xml = Element("question")
    para_q = [FakeParagraph("Câu 1: 2 + 2 = [[4]]")]
    cau_sau_xu_ly = [para_q, []]
    processor.dang_dt(cau_sau_xu_ly, xml, "TOANTHPT")

    listanswers = xml.find("listanswers")
    assert listanswers is not None
    assert listanswers.find("answer").find("content").text == "4"

# ---- Test post_process_xml giữ cấu trúc ----
def test_post_process_xml_basic(processor, tmp_path):
    xml = Element("root")
    xml_str = processor.prettify_xml(xml)
    result = processor.post_process_xml(xml_str)
    assert isinstance(result, str)
    assert "<?xml" in result

@pytest.mark.parametrize("n", range(10))
def test_random_escape_and_wrap_style(processor, n):
    # Sinh ngẫu nhiên text có ký tự đặc biệt
    text = "".join(random.choice(string.ascii_letters + string.punctuation + " <>/&") for _ in range(30))
    escaped = processor.escape_html(text)
    assert isinstance(escaped, str)
    # Đảm bảo không còn ký tự HTML chưa escape
    assert "<" not in escaped or "&lt;" in escaped
    assert ">" not in escaped or "&gt;" in escaped

    # Random style tuple (bold, italic, underline, sup, sub)
    style = tuple(random.choice([True, False]) for _ in range(5))
    wrapped = processor.wrap_style("abc", style)
    assert isinstance(wrapped, str)
    # Nếu có style True thì phải chứa tag tương ứng
    if style[0]: assert "<strong>" in wrapped
    if style[1]: assert "<i>" in wrapped
    if style[2]: assert "<u>" in wrapped
    if style[3]: assert "<sup>" in wrapped
    if style[4]: assert "<sub>" in wrapped

# ---- Random test convert_paragraph_to_html ----
@pytest.mark.parametrize("n", range(10))
def test_random_convert_paragraph_to_html(processor, n):
    text = "".join(random.choice(string.ascii_letters + " ") for _ in range(random.randint(5, 40)))
    runs = [
        FakeRun(text[random.randint(0, len(text)-1):], bold=random.choice([True, False]), italic=random.choice([True, False]))
        for _ in range(random.randint(1, 3))
    ]
    p = FakeParagraph(text, runs)
    html = processor.convert_paragraph_to_html(p)
    assert isinstance(html, str)
    assert html.startswith("<p")
    assert "</p>" in html

# ---- Random test convert_content_to_html ----
@pytest.mark.parametrize("n", range(10))
def test_random_convert_content_to_html(processor, n):
    # tạo danh sách 3 đoạn văn random
    paras = [FakeParagraph("Random " + "".join(random.choices(string.ascii_letters, k=10))) for _ in range(3)]
    html = processor.convert_content_to_html(paras)
    assert isinstance(html, str)
    assert all(p.text.split()[1] in html for p in paras)

# ---- Random test strip_html với HTML lồng nhau ----
@pytest.mark.parametrize("n", range(10))
def test_random_strip_html(processor, n):
    html = "".join(random.choice([
        "<p>abc</p>", "<strong>x</strong>", "<i>y</i>", "<u>z</u>"
    ]) for _ in range(random.randint(2, 5)))
    result = processor.strip_html(html)
    assert isinstance(result, str)
    assert "<" not in result and ">" not in result
