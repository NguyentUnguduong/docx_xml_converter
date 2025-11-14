
# docx_processor.py
import re
import base64
from io import BytesIO
from docx import Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import Table as DocxTable, _Cell
from docx.table import Table 
from docx.text.paragraph import Paragraph
from docx.text.paragraph import Paragraph as DocxParagraph
from xml.etree.ElementTree import Element, SubElement, tostring
from xml.dom import minidom
# from tinhoc_processor import TinHocProcessor # Bỏ import nếu chưa có
from typing import List, Union, Any, Iterable, Optional
import traceback
from PIL import Image
from io import BytesIO
from bs4 import BeautifulSoup

# Giả sử TinHocProcessor chưa được định nghĩa, ta tạo một lớp giả lập
# hoặc đảm bảo nó có thể được import. Nếu không, bỏ qua phần xử lý Tin học.
# class TinHocProcessor:
#     def __init__(self): pass
#     def dang_ds_tinhoc(self, cau_sau_xu_ly, xml, audio, doc): pass
#     def dang_tn_tinhoc(self, cau_sau_xu_ly, xml, audio, doc): pass
#     def dang_dt(self, cau_sau_xu_ly, xml, subject): pass
#     def dang_tl(self, cau_sau_xu_ly, xml, audio): pass

# Thử import, nếu không có thì tạo lớp giả lập
try:
    from tinhoc_processor import TinHocProcessor
except ImportError:
    class TinHocProcessor:
        def __init__(self): pass
        def dang_ds_tinhoc(self, cau_sau_xu_ly, xml, audio, doc): pass
        def dang_tn_tinhoc(self, cau_sau_xu_ly, xml, audio, doc): pass
        def dang_dt(self, cau_sau_xu_ly, xml, subject): pass
        def dang_tl(self, cau_sau_xu_ly, xml, audio): pass


class DocxProcessor:
    """Class chính xử lý DOCX"""
    def __init__(self):
        self.subjects_with_default_titles = [
            "TOANTHPT", "VATLITHPT2", "HOATHPT2", "SINHTHPT2",
            "LICHSUTHPT", "DIALITHPT", "GDCDTHPT2", "NGUVANTHPT",
            "TOANTHCS2", "KHTN", "KHXHTHCS", "GDCDTHCS2", "NGUVANTHCS2", "DGNLDHQGHN"
        ]
        self.tinhoc_subjects = ['TINHOC', 'TINHOCTHCS', 'TINHOCTHPT', 'TINHOC3']
        self.index_question = 0
        self.tinhoc_processor = TinHocProcessor()
        self.nsmap = {
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'v': 'urn:schemas-microsoft-com:vml',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    }

    def process_docx(self, file_path):
        """Xử lý file DOCX và trả về XML string hoặc danh sách lỗi"""
        # Thêm danh sách lỗi cho file này
        errors = []
        doc = Document(file_path)
        self.doc = doc  # Lưu document để truy cập parts
        self.tinhoc_processor.doc = self.doc
        body = doc.element.body
        # Parse các elements
        paragraphs = []
        for child in body:
            if isinstance(child, CT_P):
                paragraphs.append(Paragraph(child, doc))
            elif isinstance(child, CT_Tbl):
                paragraphs.append(Table(child, doc))

        # Phân tích cấu trúc
        list_hl = []
        group_of_questions = []
        current_tag = None
        current_table = None
        content_hl = False
        for idx, para in enumerate(paragraphs):
            is_table = isinstance(para, Table)
            # Xử lý table
            if is_table:
                current_table = para
                if group_of_questions and group_of_questions[-1]['questions']:
                    group_of_questions[-1]['questions'].append(current_table)
                continue
            # Paragraph rỗng
            if len(para.runs) == 0:
                continue
            text = para.text.strip()
            # Phát hiện header [tag, posttype, level]
            if re.match(r'^\[.*\]$', text):
                header = text.replace('[', '').replace(']', '')
                fields = [f.strip() for f in header.split(',')]
                if len(fields) != 3:
                    # raise ValueError(f"Sai format header: {text}")
                    errors.append(f"Sai format header: {text}")
                    continue # Bỏ qua header lỗi, tiếp tục
                dvkt, posttype, knowledge = fields
                current_tag = dvkt
                cap_do = ['NB', 'TH', 'VD', 'VDC']
                knowledge_upper = knowledge.upper()
                level = cap_do.index(knowledge_upper) if knowledge_upper in cap_do else 0
                group = {
                    'subject': dvkt.split('_')[0],
                    'tag': dvkt,
                    'original_tag': dvkt,
                    'posttype': posttype,
                    'knowledgelevel': knowledge_upper if knowledge_upper in cap_do else 'NB',
                    'level': level,
                    'questions': []
                }
                # Kiểm tra trùng lặp
                is_duplicate = any(
                    g['subject'] == group['subject'] and
                    g['tag'] == group['tag'] and
                    g['posttype'] == group['posttype'] and
                    g['knowledgelevel'] == group['knowledgelevel']
                    for g in group_of_questions
                )
                if not is_duplicate:
                    group_of_questions.append(group)
                continue
            # Phát hiện học liệu
            if text.startswith('HL:'):
                if list_hl:
                    prev_group = group_of_questions[-1]
                    group_of_questions = [{
                        'subject': prev_group['subject'],
                        'tag': prev_group['tag'],
                        'posttype': prev_group['posttype'],
                        'knowledgelevel': prev_group['knowledgelevel'],
                        'level': prev_group['level'],
                        'questions': []
                    }]
                hoc_lieu = {
                    'content': [para],
                    'groupOfQ': group_of_questions
                }
                content_hl = True
                list_hl.append(hoc_lieu)
                continue
            # Phát hiện câu hỏi
            if re.match(r'^C[ââ]u.\d', text, re.IGNORECASE):
                content_hl = False
            # Thêm vào content
            if content_hl and list_hl:
                list_hl[-1]['content'].append(para)
                continue
            if group_of_questions:
                # Gán tag cho question
                if hasattr(para, 'current_tag'):
                    para.current_tag = current_tag
                else:
                    para.current_tag = current_tag
                group_of_questions[-1]['questions'].append(para)

        # Tạo XML
        if list_hl:
            # Có học liệu
            root = Element('itemDocuments')
            for idx_hl, hoc_lieu in enumerate(list_hl):
                item_doc = self.create_hoc_lieu_xml(hoc_lieu, idx_hl)
                root.append(item_doc)
        else:
            # Chỉ có câu hỏi
            root = Element('questions')
            self.index_question = 0
            for group in group_of_questions:
                # Gọi format_questions với danh sách lỗi
                self.format_questions(group, root, errors)

        # Convert sang string
        xml_str = self.prettify_xml(root)
        xml_str = self.post_process_xml(xml_str)
        # Trả về cả XML và danh sách lỗi
        return xml_str, errors


    def create_hoc_lieu_xml(self, hoc_lieu, index_hl):
        """Tạo XML cho học liệu"""
        item_doc = Element('itemDocument')
        questions_hl = [g for g in hoc_lieu['groupOfQ'] if g['questions']]
        sub_id = SubElement(item_doc, 'subjectId')
        sub_id.text = questions_hl[0]['subject'] if questions_hl else ''
        know_id = SubElement(item_doc, 'knowledgeId')
        know_id.text = questions_hl[0]['tag'] if questions_hl else ''
        group_material = SubElement(item_doc, 'groupQuestionMaterial')
        group_material.text = str(index_hl)
        content_html = SubElement(item_doc, 'contentHtml')
        html_content = self.xu_ly_hl(hoc_lieu['content'])
        content_html.text = html_content
        list_question = SubElement(item_doc, 'listQuestion')
        for group in questions_hl:
            # Gọi format_questions với danh sách lỗi
            self.format_questions(group, list_question, [])
        return item_doc

    def xu_ly_hl(self, content):
        """
        Xử lý nội dung học liệu (HL) thành HTML hoàn chỉnh.
        - Hỗ trợ Paragraph (bold/italic/underline/sub/sup)
        - Hỗ trợ Ảnh (DrawingML / VML)
        - Hỗ trợ Bảng (bao gồm nested tables)
        - Chạy được với cả Document, _Body hoặc list phần tử
        """
        print("[DEBUG] === BẮT ĐẦU HÀM xu_ly_hl ===")
        # ... (phần code cũ không thay đổi) ...
        # =================== HELPER: EXTRACT ELEMENTS ===================
        def extract_elements(container: Any) -> List[Union[Paragraph, DocxTable]]:
            elements = []
            print(f"[DEBUG] extract_elements: container={type(container)}")
            try:
                # Nếu container có cả paragraphs và tables → dùng cách chuẩn
                if hasattr(container, "paragraphs") or hasattr(container, "tables"):
                    paragraphs = list(getattr(container, "paragraphs", []))
                    tables = list(getattr(container, "tables", []))
                    print(f"[DEBUG] Có {len(paragraphs)} paragraphs, {len(tables)} tables")
                    # Tạo list giữ thứ tự xuất hiện thật trong XML
                    body_elem = getattr(container, "_element", None)
                    if body_elem is None and hasattr(container, "_body"):
                        body_elem = getattr(container._body, "_element", None)
                    if body_elem is not None:
                        for child in body_elem.iterchildren():
                            if isinstance(child, CT_P):
                                para = Paragraph(child, container)
                                elements.append(para)
                            elif isinstance(child, CT_Tbl):
                                tbl = DocxTable(child, container)
                                elements.append(tbl)
                        print(f"[DEBUG] Trích xuất trực tiếp từ XML body: {len(elements)} phần tử")
                        return elements
                    else:
                        # fallback: nối paragraphs và tables nếu không xác định được thứ tự
                        elements = paragraphs + tables
                        print("[WARN] Không xác định được body element, nối thẳng paragraphs+tables")
                        return elements
            except Exception as e:
                print(f"[ERROR] extract_elements lỗi: {e}")
                traceback.print_exc()
            # fallback cuối cùng (cũ)
            try:
                for child in container._element.iterchildren():
                    if isinstance(child, CT_P):
                        elements.append(Paragraph(child, container))
                    elif isinstance(child, CT_Tbl):
                        elements.append(DocxTable(child, container))
            except Exception as e:
                print(f"[WARN] fallback extract_elements lỗi: {e}")
                traceback.print_exc()
            return elements
        # =================== HELPER: CONVERT PARAGRAPH ===================
        def convert_paragraph_to_html(p: Paragraph) -> str:
            """Chuyển 1 Paragraph thành HTML, đồng thời loại bỏ prefix 'HL:' nếu có."""
            html = ""
            try:
                runs = p.runs
                print(f"[DEBUG] → Paragraph có {len(runs)} runs")
                # ✅ Ghép toàn bộ text để dò prefix HL:
                full_text = "".join(run.text or "" for run in runs)
                import re
                # ✅ Regex: dò cụm 'HL:' ở đầu dòng (chấp nhận cách viết linh hoạt)
                hl_match = re.match(r"^\s*(H\s*L\s*[:：\-]\s*)", full_text, re.IGNORECASE)
                hl_cut_pos = hl_match.end() if hl_match else 0
                current_pos = 0
                for i, run in enumerate(runs):
                    text = run.text or ""
                    # ✅ Nếu có HL:, bỏ đoạn tương ứng
                    if hl_cut_pos:
                        if current_pos + len(text) <= hl_cut_pos:
                            current_pos += len(text)
                            continue  # toàn bộ run nằm trong vùng HL:
                        elif current_pos < hl_cut_pos:
                            text = text[hl_cut_pos - current_pos:]  # cắt phần HL:
                            current_pos = hl_cut_pos
                        else:
                            current_pos += len(text)
                    else:
                        current_pos += len(text)
                    # 1️⃣ Xử lý ảnh trong run (nếu có)
                    try:
                        imgs = self._get_image_tags_from_run(run)
                        if imgs:
                            for it in imgs:
                                html += it
                    except Exception as e:
                        print(f"[WARN] Lỗi khi lấy ảnh từ run {i}: {e}")
                    # 2️⃣ Bỏ qua nếu text rỗng
                    if not text.strip():
                        continue
                    # 3️⃣ Áp dụng format
                    pieces = self.escape_html(text)
                    if run.bold:
                        pieces = f"<b>{pieces}</b>"
                    if run.italic:
                        pieces = f"<i>{pieces}</i>"
                    if run.underline:
                        pieces = f"<u>{pieces}</u>"
                    if getattr(run.font, "subscript", False):
                        pieces = f"<sub>{pieces}</sub>"
                    if getattr(run.font, "superscript", False):
                        pieces = f"<sup>{pieces}</sup>"
                    html += pieces
                # 4️⃣ Xử lý ảnh inline / ngoài runs
                try:
                    inline_imgs = p._element.xpath(".//a:blip/@r:embed")
                    for rId in inline_imgs:
                        tag = self._make_img_tag_from_rid(rId)
                        if tag:
                            html += tag
                except Exception as e:
                    print(f"[WARN] Lỗi khi xử lý ảnh inline: {e}")
            except Exception as e:
                print(f"[ERROR] Lỗi convert_paragraph_to_html: {e}")
                traceback.print_exc()
            return f"<p>{html}</p>  "
        # =================== CHUẨN BỊ DANH SÁCH PHẦN TỬ ===================
        if isinstance(content, list):
            all_elements = content
            print(f"[DEBUG] Đầu vào là list, số phần tử: {len(all_elements)}")
        elif hasattr(content, "_element"):
            all_elements = extract_elements(content)
            print(f"[DEBUG] Đầu vào là document/body, trích xuất {len(all_elements)} phần tử")
        else:
            print(f"[WARN] Loại đầu vào không hỗ trợ: {type(content)}")
            return ""
        # =================== DUYỆT TOÀN BỘ PHẦN TỬ ===================
        html_parts = []
        for i, el in enumerate(all_elements):
            print(f"[DEBUG] --- Xử lý phần tử {i}: {type(el).__name__}")
            try:
                if isinstance(el, Paragraph):
                    html_parts.append(convert_paragraph_to_html(el))
                elif isinstance(el, DocxTable):
                    html_parts.append(self.convert_table_to_html(el))
                else:
                    print(f"[WARN] Bỏ qua phần tử loại: {type(el)}")
            except Exception as e:
                print(f"[ERROR] Lỗi xử lý phần tử {i}: {e}")
                traceback.print_exc()
                html_parts.append(f"<!-- ERROR tại phần tử {i} -->")
        # =================== KẾT THÚC ===================
        html = "".join(html_parts)
        print("[DEBUG] === KẾT THÚC HÀM xu_ly_hl ===")
        return html

    def convert_table_to_html(self, table: DocxTable) -> str:
        """
        Convert table sang HTML (hỗ trợ nested table).
        NOTE: Bỏ colspan (theo yêu cầu).
        """
        print("[DEBUG][convert_table_to_html] === BẮT ĐẦU XỬ LÝ TABLE ===")
        html = "<table class='table-material-question'>"
        try:
            # Dùng API python-docx chính thức: table.rows, cell.paragraphs, cell.tables
            for r_idx, row in enumerate(table.rows):
                html += "<tr>"
                print(f"[DEBUG] → Row {r_idx}, số ô: {len(row.cells)}")
                for c_idx, cell in enumerate(row.cells):
                    print(f"[DEBUG]   → Cell ({r_idx},{c_idx}) bắt đầu xử lý")
                    parts: List[str] = []
                    # 1) Nếu cell có nested tables theo python-docx -> xử lý trước
                    try:
                        if hasattr(cell, "tables") and cell.tables:
                            print(f"[DEBUG]   Nested tables count in cell ({r_idx},{c_idx}): {len(cell.tables)}")
                            for nt_idx, nested in enumerate(cell.tables):
                                try:
                                    parts.append(self.convert_table_to_html(nested))
                                except Exception as e:
                                    print(f"[ERROR] Lỗi xử lý nested table ({r_idx},{c_idx},{nt_idx}): {e}")
                                    traceback.print_exc()
                                    parts.append("<!-- ERROR nested -->")
                    except Exception as e:
                        print(f"[WARN] Không thể đọc cell.tables tại ({r_idx},{c_idx}): {e}")
                    # 2) Thêm các paragraph trong cell (theo thứ tự)
                    try:
                        if hasattr(cell, "paragraphs"):
                            for p_idx, p in enumerate(cell.paragraphs):
                                try:
                                    # convert_paragraph_to_html đã tồn tại và trả về <p>..</p>
                                    para_html = self.convert_content_to_html(p)
                                    parts.append(para_html)
                                except Exception as e:
                                    print(f"[WARN] Lỗi convert paragraph trong cell ({r_idx},{c_idx},p{p_idx}): {e}")
                                    traceback.print_exc()
                                    # fallback: raw text
                                    try:
                                        parts.append(p.text)
                                    except Exception:
                                        parts.append("")
                    except Exception as e:
                        print(f"[WARN] Không thể đọc cell.paragraphs tại ({r_idx},{c_idx}): {e}")
                    # 3) Join parts, trim; nếu rỗng -> dùng &nbsp;
                    cell_html = "".join(parts).strip()
                    if not cell_html:
                        cell_html = "&nbsp;"
                    # 4) **Không sinh colspan nữa** (user yêu cầu xóa colspan)
                    html += f"<td>{cell_html}</td>"
                html += "</tr>"
        except Exception as e:
            print(f"[ERROR] convert_table_to_html gặp lỗi tổng thể: {e}")
            traceback.print_exc()
        html += "</table><br>"
        print("[DEBUG][convert_table_to_html] === KẾT THÚC ===")
        return html

    def wrap_style(self, text, style):
        """Đóng gói text với style tuple"""
        bold, italic, underline, sup, sub = style
        if bold:
            text = f"<strong>{text}</strong>"
        if italic:
            text = f"<i>{text}</i>"
        if underline:
            text = f"<u>{text}</u>"
        if sup:
            text = f"<sup>{text}</sup>"
        if sub:
            text = f"<sub>{text}</sub>"
        return text

    def format_questions(self, group, questions_xml, errors):
        """Format các câu hỏi, nhận thêm danh sách errors để ghi lỗi"""
        group_of_q = []
        for para in group['questions']:
            if isinstance(para, Table):
                if group_of_q and group_of_q[-1]:
                    group_of_q[-1]['items'].append(para)
                continue
            text = para.text.strip().lower()
            # Phát hiện câu hỏi mới
            if re.match(r'^c[ââ]u.\d', text):
                question_tag = getattr(para, 'current_tag', None) or group.get('original_tag') or group['tag']
                question = {
                    'items': [para],
                    'question_tag': question_tag
                }
                group_of_q.append(question)
            elif group_of_q:
                group_of_q[-1]['items'].append(para)

        # Xử lý từng câu hỏi
        for idx, question_dict in enumerate(group_of_q):
            each_question_xml = Element('question')
            # Metadata
            SubElement(each_question_xml, 'indexGroupQuestionMaterial').text = str(self.index_question)
            SubElement(each_question_xml, 'subject').text = group['subject']
            question_tag = question_dict['question_tag']
            SubElement(each_question_xml, 'tag').text = question_tag
            SubElement(each_question_xml, 'posttype').text = group['posttype']
            SubElement(each_question_xml, 'knowledgelevel').text = group['knowledgelevel']
            SubElement(each_question_xml, 'levelquestion').text = str(group['level'])
            # Xử lý nội dung câu hỏi
            try:
                # Gọi protocol_of_q với danh sách lỗi
                self.protocol_of_q(question_dict['items'], each_question_xml, group['subject'], errors, idx + 1) # idx+1 là số thứ tự câu hỏi
            except Exception as e:
                # Nếu protocol_of_q ném lỗi không bắt được (nên ít xảy ra sau khi sửa)
                # thì vẫn ghi vào danh sách lỗi và tiếp tục
                error_msg = f"Lỗi không xử lý được khi phân tích câu hỏi {idx + 1}: {str(e)}"
                errors.append(error_msg)
                print(f"[ERROR] format_questions: {error_msg}")
                traceback.print_exc()
                continue # Bỏ qua câu hỏi lỗi, tiếp tục với câu tiếp theo

            self.index_question += 1
            questions_xml.append(each_question_xml)


    def _get_image_tags_from_run(self, run):
            """
            Tìm image references trong run._r (blip / v:imagedata),
            trả về list tag <img src="data:..."/> (base64).
            """
            imgs = []
            try:
                # truy cập vào phần XML thô của run
                r = run._r
                # 1) DrawingML blip (thường thấy với images chèn hiện đại)
                blips = r.xpath('.//a:blip', namespaces={'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'})
                for blip in blips:
                    # attribute chứa relationship id
                    rId = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                    if rId:
                        img_tag = self._make_img_tag_from_rid(rId)
                        if img_tag:
                            imgs.append(img_tag)
                # 2) VML (cũ hơn) - v:imagedata với attribute r:id
                picts = r.xpath('.//v:imagedata', namespaces={'v': 'urn:schemas-microsoft-com:vml'})
                for pict in picts:
                    rId = pict.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                    if rId:
                        img_tag = self._make_img_tag_from_rid(rId)
                        if img_tag:
                            imgs.append(img_tag)
            except Exception:
                # im lặng nếu không tìm thấy hoặc lỗi, tránh crash
                pass
            return imgs

    # def _make_img_tag_from_rid(self, rId):
    #     # (phần code cũ không thay đổi)
    #     ...

    # def _make_img_tag_from_rid(self, rId):
    #     """
    #     Dùng rId để lấy image part từ self.doc.part.related_parts,
    #     trả về một thẻ <img src="data:..."> hoặc None.
    #     """
    #     try:
    #         part = self.doc.part.related_parts.get(rId)
    #         if not part:
    #             # fallback: tìm trong các rels
    #             for rel in self.doc.part.rels.values():
    #                 try:
    #                     target = getattr(rel, 'target_part', None)
    #                     if target and 'image' in getattr(target, 'content_type', ''):
    #                         if rel.rId == rId:
    #                             part = target
    #                             break
    #                 except Exception:
    #                     continue
    #         if not part:
    #             print(f"[DEBUG] Không tìm thấy part cho rId={rId}")
    #             return None
    #         img_bytes = part.blob
    #         content_type = getattr(part, 'content_type', 'image/png')
    #         # --- Đọc kích thước gốc ---
    #         try:
    #             img = Image.open(BytesIO(img_bytes))
    #             width, height = img.size
    #             print(f"[DEBUG] Ảnh rId={rId} size: {width}x{height}")
    #         except Exception as e:
    #             print(f"[WARN] Không đọc được kích thước ảnh: {e}")
    #             width, height = 300, 200  # fallback
    #         # encode base64
    #         b64 = base64.b64encode(img_bytes).decode('ascii')
    #         # --- Sinh tag HTML ---
    #         style = f'style="max-width:{width}px; height:{height};"'
    #         # hoặc nếu muốn cố định tỉ lệ: style = f'style="width:{width}px;height:{height}px;"'
    #         return f'<center><img src="data:{content_type};base64,{b64}" {style} /></center>'
    #     except Exception as e:
    #         print(f"[ERROR] _make_img_tag_from_rid lỗi: {e}")
    #         import traceback; traceback.print_exc()
    #         return None

  

    def _make_img_tag_from_rid(self, rId):
        """
        Dùng rId để lấy image part từ self.doc.part.related_parts,
        trả về một thẻ <img src="data:..."> hoặc None.
        """
        try:
            part = self.doc.part.related_parts.get(rId)
            if not part:
                # fallback: tìm trong các rels
                for rel in self.doc.part.rels.values():
                    try:
                        target = getattr(rel, 'target_part', None)
                        if target and 'image' in getattr(target, 'content_type', ''):
                            if rel.rId == rId:
                                part = target
                                break
                    except Exception:
                        continue
            if not part:
                print(f"[DEBUG] Không tìm thấy part cho rId={rId}")
                return None

            img_bytes = part.blob
            content_type = getattr(part, 'content_type', 'image/png')

            # --- ĐỌC KÍCH THƯỚC GỐC ---
            try:
                img = Image.open(BytesIO(img_bytes))
                original_width, original_height = img.size
                print(f"[DEBUG] Ảnh rId={rId} kích thước gốc: {original_width}x{original_height}")

                # --- GIỚI HẠN KÍCH THƯỚC ẢNH ---
                max_width = 800
                max_height = 800

                # Tính tỷ lệ giữ nguyên tỷ lệ khung hình (aspect ratio)
                ratio = min(max_width / original_width, max_height / original_height)

                if ratio < 1:
                    new_width = int(original_width * ratio)
                    new_height = int(original_height * ratio)
                    print(f"[DEBUG] Ảnh rId={rId} sau khi resize: {new_width}x{new_height}")

                    # Resize ảnh
                    img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
                    # hoặc Image.Resampling.LANCZOS nếu dùng Pillow >= 10
                else:
                    new_width = original_width
                    new_height = original_height

            except Exception as e:
                print(f"[WARN] Không đọc được kích thước ảnh: {e}")
                return None  # Không xử lý được ảnh, trả về None

            # --- GHIẢM KÍCH THƯỚC FILE (nếu cần) ---
            output = BytesIO()
            img_format = img.format or 'PNG'

            # Nếu là ảnh JPEG, có thể nén với chất lượng
            if img_format in ('JPEG', 'JPG'):
                img.save(output, format=img_format, quality=80, optimize=True)
            else:
                # Với PNG, có thể nén nhẹ hơn, hoặc chuyển sang JPEG nếu cần giảm size
                img.save(output, format=img_format, optimize=True)

            resized_img_bytes = output.getvalue()
            output.close()

            # --- KIỂM TRA KÍCH THƯỚC FILE ---
            max_file_size = 1 * 1024 * 1024  # 1MB
            if len(resized_img_bytes) > max_file_size:
                print(f"[DEBUG] Ảnh rId={rId} quá lớn ({len(resized_img_bytes)} bytes), chuyển sang JPEG để nén hơn...")
                # Chuyển sang JPEG để giảm dung lượng
                output_jpg = BytesIO()
                if img.mode in ('RGBA', 'P') and img_format == 'PNG':
                    # Chuyển RGBA sang RGB để lưu JPEG
                    if img.mode == 'P':
                        img = img.convert("RGBA")
                    if img.mode == 'RGBA':
                        background = Image.new('RGB', img.size, (255, 255, 255))
                        background.paste(img, mask=img.split()[-1])
                        img = background
                img.save(output_jpg, format='JPEG', quality=75, optimize=True)
                resized_img_bytes = output_jpg.getvalue()
                output_jpg.close()

            # --- ENCODE BASE64 ---
            b64 = base64.b64encode(resized_img_bytes).decode('ascii')

            # --- SINH TAG HTML ---
            style = f'style="width:{new_width}px; height:{new_height}px;"'
            return f'<center><img src="data:{content_type};base64,{b64}" {style} /></center>'

        except Exception as e:
            print(f"[ERROR] _make_img_tag_from_rid lỗi: {e}")
            import traceback; traceback.print_exc()
            return None

    def protocol_of_q(self, question, each_question_xml, subject, errors, question_index):
        """Phân tích cấu trúc câu hỏi, nhận danh sách errors và số thứ tự câu hỏi question_index"""
        # Chia thành phần: nội dung câu hỏi và lời giải
        thanh_phan_1q = []
        for idx, para in enumerate(question):
            if idx == 0:
                thanh_phan_1q.append([para])
                continue
            if isinstance(para, Paragraph):
                text = para.text.strip().lower()
                if re.match(r'^l[ờờ]i gi[ảả]i', text):
                    thanh_phan_1q.append([])
                    continue
            if thanh_phan_1q:
                thanh_phan_1q[-1].append(para)

        if len(thanh_phan_1q) < 2:
            # raise ValueError(f"Thiếu 'Lời giải' trong câu: {question[0].text[:50]}")
            error_msg = f"Thiếu 'Lời giải' trong câu hỏi {question_index}"
            errors.append(error_msg)
            print(f"[ERROR] protocol_of_q: {error_msg}")
            # Trả về hoặc tiếp tục để xử lý các phần khác nếu có thể, mặc dù thiếu lời giải
            # Có thể thêm phần tử giả hoặc bỏ qua câu hỏi này
            # Trong trường hợp này, ta tiếp tục để tạo XML rỗng hoặc với thông tin cơ bản
            # Tuy nhiên, để đảm bảo XML hợp lệ, ta nên bỏ qua phần xử lý sâu hơn
            # hoặc tạo các phần tử cần thiết với giá trị mặc định.
            # Ví dụ: Tạo phần tử trống cho contentquestion và explainquestion
            SubElement(each_question_xml, 'contentquestion').text = ''
            SubElement(each_question_xml, 'explainquestion').text = f'--- LỖI: Thiếu lời giải ---'
            SubElement(each_question_xml, 'typeAnswer').text = '0' # Mặc định
            listanswers = SubElement(each_question_xml, 'listanswers') # Danh sách rỗng
            return # Kết thúc xử lý câu hỏi này

        # Phân tích nội dung câu hỏi và lời giải
        thanh_phan_cau_hoi = []
        link_cau_hoi = []
        # Xử lý links và nội dung
        for para in thanh_phan_1q[0]:
            if isinstance(para, Paragraph):
                text = para.text.strip()
                # Phát hiện Audio
                if text.startswith('Audio:'):
                    link_cau_hoi.append(text)
                    continue
                # Phát hiện URLs
                urls = re.findall(r'https?://[^\s]+', text)
                for url in urls:
                    if url not in link_cau_hoi:
                        link_cau_hoi.append(url)
                if urls and not text.replace(urls[0], '').strip():
                    continue
            thanh_phan_cau_hoi.append(para)

        # Xử lý links
        self.xu_ly_link_cau_hoi(link_cau_hoi, each_question_xml)

        # Phân tích lời giải
        thanh_phan_hdg = []
        link_speech_explain = []
        for idx, para in enumerate(thanh_phan_1q[1]):
            if idx == 0:
                thanh_phan_hdg.append([para])
                continue
            if isinstance(para, Paragraph):
                text = para.text.strip()
                if text.startswith('###'):
                    thanh_phan_hdg.append([])
                    continue
                # URLs trong HDG
                urls = re.findall(r'https?://[^\s]+', text)
                for url in urls:
                    link_speech_explain.append(url)
                    continue
            if thanh_phan_hdg:
                thanh_phan_hdg[-1].append(para)

        # Xử lý urlSpeechExplain
        if link_speech_explain:
            if len(link_speech_explain) > 1:
                # raise ValueError(f"HDG chỉ được có 1 link TTS: {link_speech_explain}")
                error_msg = f"HDG có nhiều hơn 1 link TTS ở câu hỏi {question_index}: {link_speech_explain}"
                errors.append(error_msg)
                print(f"[ERROR] protocol_of_q: {error_msg}")
                # Có thể chọn 1 link hoặc bỏ qua, ở đây ta chọn link đầu tiên
                if link_speech_explain[0].endswith(('.mp3', '.mp4')):
                    SubElement(each_question_xml, 'urlSpeechExplain').text = link_speech_explain[0]

        # Xác định dạng câu hỏi
        answer = thanh_phan_hdg[0][0].text.strip() if thanh_phan_hdg[0] else ''
        cau_sau_xu_ly = [thanh_phan_cau_hoi, thanh_phan_hdg]
        audio = [link for link in link_cau_hoi if 'Audio:' in link]
        # Routing theo subject
        if self.is_tinhoc_subject(subject):
            self.route_to_tinhoc_module(cau_sau_xu_ly, each_question_xml, audio, answer, subject, errors, question_index)
        else:
            self.route_to_default_module(cau_sau_xu_ly, each_question_xml, audio, answer, subject, errors, question_index)


    def is_tinhoc_subject(self, subject):
        """Kiểm tra có phải môn tin học không"""
        return any(subject.startswith(tinhoc) for tinhoc in self.tinhoc_subjects)

    def route_to_tinhoc_module(self, cau_sau_xu_ly, xml, audio, answer, subject, errors, question_index):
        """Xử lý cho môn Tin học, nhận danh sách lỗi và số câu hỏi"""
        # ✅ Gọi từ instance tinhoc_processor
        if re.match(r'^\d+', answer):
            if len(answer) > 1 and re.match(r'^[01]+', answer):
                self.tinhoc_processor.dang_ds_tinhoc(cau_sau_xu_ly, xml, audio, self.doc)
            else:
                self.tinhoc_processor.dang_tn_tinhoc(cau_sau_xu_ly, xml, audio, self.doc)
        elif answer.startswith('##'):
            self.dang_dt(cau_sau_xu_ly, xml, subject)
        else:
            self.dang_tl(cau_sau_xu_ly, xml, audio)

    def route_to_default_module(self, cau_sau_xu_ly, xml, audio, answer, subject, errors, question_index):
        """Xử lý cho môn thông thường, nhận danh sách lỗi và số câu hỏi"""
        if re.match(r'^\d+', answer):
            if len(answer) > 1 and re.match(r'^[01]+', answer):
                self.dang_ds(cau_sau_xu_ly, xml, audio)
            else:
                self.dang_tn(cau_sau_xu_ly, xml, audio)
        elif answer.startswith('##'):
            self.dang_dt(cau_sau_xu_ly, xml, subject)
        else:
            self.dang_tl(cau_sau_xu_ly, xml, audio)

    def xu_ly_link_cau_hoi(self, links, xml):
        """Xử lý links trong câu hỏi"""
        one_tts = False
        one_media = False
        for link in links:
            if link.startswith('Audio:'):
                continue
            if link.endswith(('.mp3', '.mp4')):
                if one_tts:
                    # raise ValueError(f"Chỉ được 1 link TTS: {link}")
                    print(f"[WARN] Có nhiều hơn 1 link TTS trong câu hỏi, bỏ qua: {link}")
                    continue
                SubElement(xml, 'urlSpeechContent').text = link
                one_tts = True
            else:
                if one_media:
                    # raise ValueError(f"Chỉ được 1 link Video: {link}")
                    print(f"[WARN] Có nhiều hơn 1 link Video trong câu hỏi, bỏ qua: {link}")
                    continue
                if 'vimeo.com' in link:
                    code = link.split('vimeo.com/')[1]
                    parts = code.split('/')
                    if len(parts) > 1:
                        code = f"{parts[0]}?h={parts[1].split('?share')[0]}"
                    else:
                        code = parts[0]
                    SubElement(xml, 'contentMedia').text = code
                    SubElement(xml, 'typeContentMedia').text = 'CodeVimeo'
                    one_media = True
                elif 'youtu' in link:
                    if 'watch?v=' in link:
                        code = link.split('watch?v=')[1]
                    elif 'youtu.be/' in link:
                        code = link.split('youtu.be/')[1].split('?')[0]
                    else:
                        continue
                    SubElement(xml, 'contentMedia').text = code
                    SubElement(xml, 'typeContentMedia').text = 'CodeYouTuBe'
                    one_media = True

    # ... (các hàm convert_content_to_html, dang_tn, list_answers_tn, strip_html, hdg_tn, dang_ds, dang_dt, dang_tl, convert_b4_add, convert_normal_paras, escape_html, prettify_xml) ...
    # Các hàm này không cần thay đổi để phù hợp với cơ chế mới, trừ khi chúng có thể ném lỗi và cần được xử lý riêng.
    # Tuy nhiên, để an toàn, ta có thể bao bọc các hàm chính được gọi từ format_questions trong try-except.

    def convert_content_to_html(self, paragraphs):
        """
        Chuyển list Paragraph / Table sang HTML hoàn chỉnh, giữ table, ảnh, math-latex.
        KHÔNG tự bọc <div class='content'> để tránh lặp.
        Hỗ trợ flatten đệ quy: chấp nhận paragraphs là Paragraph, Table,
        list/tuple lồng nhau ở bất kỳ mức độ nào.
        """
        from docx.table import Table
        # Đệ quy flatten: trả về list các phần tử không phải list/tuple nữa
        def _flatten(items):
            for it in items:
                if isinstance(it, (list, tuple)):
                    yield from _flatten(it)
                else:
                    yield it
        # Nếu người gọi chuyền 1 object không phải iterable (ví dụ một Paragraph),
        # ta chuẩn hóa thành list để xử lý thống nhất.
        if paragraphs is None:
            flat = []
        elif isinstance(paragraphs, (list, tuple)):
            flat = list(_flatten(paragraphs))
        else:
            # Một phần tử đơn lẻ (có thể là Paragraph hoặc Table)
            flat = [paragraphs]
        string_content = ""
        for para in flat:
            # Bảo vệ: nếu para là None thì bỏ qua
            if para is None:
                continue
            # Nếu là Table (obj từ python-docx), xử lý riêng
            if isinstance(para, Table):
                string_content += self.convert_table_to_html(para)
                string_content += "<br>"
                continue
            # Nếu là string (đã chuyển trước đó), thêm trực tiếp
            if isinstance(para, str):
                string_content += para + "<br>"
                continue
            # Một số đối tượng paragraph-like có thể không đến từ python-docx
            # nhưng có attribute 'runs' — kiểm tra trước khi gọi convert_normal_paras
            new_children = []
            try:
                # Nếu paragraph không phải object paragraph hợp lệ, convert_normal_paras có thể ném
                self.convert_normal_paras(para, 0, new_children)
                string_content += "".join(new_children)
            except TypeError:
                # Thử gọi convert_normal_paras theo kiểu cũ (nếu hàm được thiết kế trả về string/list)
                try:
                    res = self.convert_normal_paras(para)
                except Exception as e:
                    # Nếu vẫn lỗi, chuyển sang fallback: str(para)
                    string_content += str(para)
                else:
                    if isinstance(res, str):
                        string_content += res
                    elif isinstance(res, list):
                        string_content += "".join(res)
                    else:
                        string_content += str(res)
            except AttributeError:
                # Thường xảy ra khi para là 1 list lồng mà chưa flatten đúng mức
                # Fallback robust: chuyển thành str(para)
                string_content += str(para)
            string_content += "<br>"
        # Xử lý math-latex
        import re
        math_latex = re.compile(r"\$[^$]*\$")
        string_content = math_latex.sub(lambda m: f'<span class="math-tex">{m.group()}</span>', string_content)
        return string_content.strip()

    def dang_tn(self, cau_sau_xu_ly, xml, audio):
        """
        Xử lý dạng Trắc nghiệm (typeAnswer=0, template=0)
        - Đáp án đúng được xác định bằng số 1,2,3,4 trong phần Lời giải (1=A, 2=B, 3=C, 4=D)
        """
        import re
        from xml.etree.ElementTree import SubElement
        from docx.text.paragraph import Paragraph
        from docx.table import Table
        SubElement(xml, 'typeAnswer').text = '0'
        SubElement(xml, 'typeViewContent').text = '0'
        SubElement(xml, 'template').text = '0'
        # ===== 1️⃣ Xử lý phần nội dung câu hỏi =====
        content_part = []
        answers_part = []
        for para in cau_sau_xu_ly[0]:
            if isinstance(para, Paragraph):
                text = para.text.strip()
                # Nhận diện các dòng A. B. C. D.
                if re.match(r'^[A-D]\.', text):
                    answers_part.append(para)
                else:
                    content_part.append(para)
            elif isinstance(para, Table):
                content_part.append(para)
        # HTML câu hỏi
        content_html = self.convert_content_to_html(content_part)
        if audio and len(audio[0]) > 8:
            link = audio[0].replace('Audio:', '').strip()
            content_html += f'<audio controls=""><source src="{link}" type="audio/mpeg"></audio>'
        SubElement(xml, 'contentquestion').text = content_html.strip()
        # ===== 2️⃣ Tìm đáp án đúng từ phần Lời giải =====
        correct_index = None  # chỉ số 0-based của đáp án đúng
        if len(cau_sau_xu_ly) > 1 and cau_sau_xu_ly[1]:
            # Lấy đoạn đầu tiên của phần lời giải
            first = cau_sau_xu_ly[1][0]
            if isinstance(first, list):
                # Nếu là danh sách Paragraph
                for p in first:
                    if hasattr(p, 'text'):
                        m = re.search(r'\b([1-4])\b', p.text.strip())
                        if m:
                            correct_index = int(m.group(1)) - 1
                            break
            elif hasattr(first, 'text'):
                m = re.search(r'\b([1-4])\b', first.text.strip())
                if m:
                    correct_index = int(m.group(1)) - 1
        # ===== 3️⃣ Sinh danh sách đáp án =====
        listanswers = SubElement(xml, 'listanswers')
        for i, para in enumerate(answers_part):
            # Bỏ prefix A./B./C./D.
            text = re.sub(r'^[A-D]\.\s*', '', para.text.strip())
            content_html = f'<p>{text}</p>'
            answer_el = SubElement(listanswers, 'answer')
            SubElement(answer_el, 'index').text = str(i)
            SubElement(answer_el, 'content').text = content_html
            SubElement(answer_el, 'isanswer').text = 'TRUE' if i == correct_index else 'FALSE'
        # ===== 4️⃣ Gọi hdg_tn() để xử lý phần giải thích chi tiết =====
        self.hdg_tn(cau_sau_xu_ly[1] if len(cau_sau_xu_ly) > 1 else None, xml)

    def list_answers_tn(self, content, answer_para, xml):
            """Tạo danh sách đáp án TN, bỏ prefix A./B./C./D. và KHÔNG bọc <div class='content'>."""
            import re
            multiple_choices = []
            for array_para in content:
                choice_html = self.convert_content_to_html(array_para if isinstance(array_para, list) else [array_para])
                # Bỏ prefix A. B. C. D. nếu có (đầu câu)
                choice_html = re.sub(r"^(<[^>]+>)*\s*[A-Za-z][\.\)]\s*", "", choice_html)
                multiple_choices.append(choice_html.strip())
            # Lấy đáp án đúng
            if isinstance(answer_para, list) and len(answer_para) > 0:
                answer_text = answer_para[0].text.strip()
            else:
                answer_text = answer_para.text.strip()
            number_of_answer = [c for c in answer_text if c.isdigit()]
            listanswers = SubElement(xml, 'listanswers')
            for i, choice in enumerate(multiple_choices):
                answer = SubElement(listanswers, 'answer')
                SubElement(answer, 'index').text = str(i)
                content_elem = SubElement(answer, 'content')
                # Không bọc <div> nữa, chỉ giữ nội dung HTML thuần
                content_elem.text = choice
                is_correct = 'TRUE' if str(i + 1) in number_of_answer else 'FALSE'
                SubElement(answer, 'isanswer').text = is_correct

    # Hàm tiện ích loại bỏ thẻ HTML
    import re
    def strip_html(self, html_text):
        # Loại bỏ tất cả thẻ <...>
        text = re.sub(r'<[^>]+>', '', html_text)
        # Loại bỏ các khoảng trắng thừa
        text = text.strip()
        return text

    def hdg_tn(self, array_hdg, xml: Element):
        """
        Hướng dẫn giải TN, giữ HTML (ảnh/table)
        - Nếu có hướng dẫn chi tiết thì thêm explainquestion
        - Nếu chỉ có đáp án đúng thì không thêm
        """
        import re
        from xml.etree.ElementTree import SubElement
        if not array_hdg:
            return
        # Xóa thẻ explainquestion cũ nếu có
        existing_explain = xml.find('explainquestion')
        if existing_explain is not None:
            xml.remove(existing_explain)
        hdg_raw = ''
        # Ghép nội dung thô từ array_hdg
        if isinstance(array_hdg, list):
            for part in array_hdg:
                if hasattr(part, "text"):
                    hdg_raw += part.text.strip() + " "
                elif isinstance(part, list):
                    for p in part:
                        if hasattr(p, "text"):
                            hdg_raw += p.text.strip() + " "
        # Chuyển sang HTML (giữ nguyên tag ảnh/table)
        hdg_html = self.convert_content_to_html(array_hdg)
        plain = re.sub(r'<[^>]+>', '', hdg_html).strip()
        explain_text = ""
        # Nếu có nội dung giải thích thực sự
        if len(plain) > 4:
            explain_text = hdg_html.strip()
            # --- 1) Bỏ số hoặc chữ đáp án đầu dòng, kể cả khi nó bị bọc trong thẻ HTML ---
            # Ví dụ: "<strong>1</strong><br>" hoặc "<strong>A</strong>:" hoặc "1. " ...
            explain_text = re.sub(
                r'^\s*(?:<[^>]+>\s*)*(?:\d+|[A-Da-d])(?:\s*</[^>]+>\s*)*(?:\s*(?:<br\s*/?>|:|\.|,))?\s*',
                '',
                explain_text,
                flags=re.IGNORECASE | re.UNICODE
            )
            # --- 2) Bỏ tiền tố "Giải thích:" kể cả khi bị bọc trong thẻ ---
            # Ví dụ: "<strong>Giải thích:</strong><br>" hoặc "Giải thích<br>"
            explain_text = re.sub(
                r'^\s*(?:<[^>]+>\s*)*Giải\s*thích\s*[:：]?(?:\s*</[^>]+>\s*)*(?:\s*(?:<br\s*/?>))?\s*',
                '',
                explain_text,
                flags=re.IGNORECASE | re.UNICODE
            ).strip()
            # Chỉ thêm thẻ nếu còn nội dung sau khi làm sạch
            if explain_text:
                SubElement(xml, 'explainquestion').text = explain_text.strip()

    def dang_ds(self, cau_sau_xu_ly, xml, audio):
        """Xử lý dạng Đúng/Sai, tách đúng phần phát biểu và HDG"""
        SubElement(xml, 'typeAnswer').text = '1'
        SubElement(xml, 'typeViewContent').text = '0'
        SubElement(xml, 'template').text = '0'
        import re
        paragraphs = cau_sau_xu_ly[0]
        statements = []
        intro_paras = []
        # ✅ Phân loại phần mở đầu và các phát biểu
        for para in paragraphs:
            if isinstance(para, Paragraph) and re.match(r'^[a-d]\s*[\.\)]', para.text.strip(), re.IGNORECASE):
                statements.append(para)
            else:
                intro_paras.append(para)
        # ✅ Phần mở đầu (mô tả tình huống)
        content_html = self.convert_content_to_html(intro_paras)
        if audio and len(audio[0]) > 8:
            link = audio[0].replace('Audio:', '').strip()
            content_html += f'<audio controls=""><source src="{link}" type="audio/mpeg"></audio>'
        SubElement(xml, 'contentquestion').text = content_html
        # ✅ Danh sách phát biểu a/b/c/d
        listanswers = SubElement(xml, 'listanswers')
        for i, para in enumerate(statements):
            ans_html = self.convert_content_to_html([para])
            # --- Bỏ prefix a) / b. / c) / d) (kể cả có tag HTML) ---
            ans_html = re.sub(
                r'^\s*(<[^>]+>)*\s*([A-Da-d])\s*[\.\)]\s*',
                '',
                ans_html
            )
            # cũng bỏ trường hợp prefix nằm trong thẻ <strong> hoặc <b>
            # ans_html = re.sub(
            #     r'^(<strong>|<b>)?\s*([A-Da-d])[\.\)]\s*(</strong>|</b>)?',
            #     '',
            #     ans_html
            # )
            ans_html = re.sub(
                r'^\s*(?:<[^>]*>)*\s*[A-Da-d]\s*(?:<[^>]*>)*\s*[\.\)]\s*(?:<[^>]*>)*\s*',
                '',
                ans_html,
                flags=re.IGNORECASE
            )
            answer = SubElement(listanswers, 'answer')
            SubElement(answer, 'index').text = str(i)
            SubElement(answer, 'content').text = ans_html
            SubElement(answer, 'isanswer').text = 'FALSE'  # tạm thời FALSE, sẽ cập nhật sau
        # ✅ Lấy chuỗi đáp án đúng/sai (ví dụ: 0111, 1010, ...)
        if len(cau_sau_xu_ly[1]) > 0:
            if isinstance(cau_sau_xu_ly[1][0], list):
                ans_text = cau_sau_xu_ly[1][0][0].text.strip()
            else:
                ans_text = cau_sau_xu_ly[1][0].text.strip()
            for i, ch in enumerate(ans_text):
                if i < len(listanswers):
                    listanswers[i].find('isanswer').text = 'TRUE' if ch == '1' else 'FALSE'
        # ✅ Hướng dẫn giải (HDG)
        if len(cau_sau_xu_ly[1]) > 1:
            flat_hdg = []
            for item in cau_sau_xu_ly[1][1:]:
                if isinstance(item, list):
                    flat_hdg.extend(item)
                else:
                    flat_hdg.append(item)
            hdg_html = self.convert_content_to_html(flat_hdg)
        else:
            hdg_html = ''
        SubElement(xml, 'explainquestion').text = hdg_html

    # def dang_dt(self, cau_sau_xu_ly, xml, subject):
    #     """
    #     Dạng điền đáp án (typeAnswer=5) - rút gọn, không dùng normalize/unescape.
    #     Tìm đáp án trực tiếp từ [[...]] rồi xây XML đúng format (contentquestion, listanswers, explainquestion).
    #     """
    #     # ===== 1. Meta =====
    #     SubElement(xml, 'typeAnswer').text = '5'
    #     SubElement(xml, 'typeViewContent').text = '0'
    #     SubElement(xml, 'template').text = '23'
    #     # ===== 2. Hint (nếu có) =====
    #     if len(cau_sau_xu_ly) > 1 and isinstance(cau_sau_xu_ly[1], list) and len(cau_sau_xu_ly[1]) > 1:
    #         hint_html = self.convert_b4_add(cau_sau_xu_ly[1][1])
    #         SubElement(xml, 'hintQuestion').text = hint_html
    #     # ===== 3. Lấy nội dung gốc và tìm đáp án [[...]] từ đó =====
    #     raw_html = self.convert_b4_add(cau_sau_xu_ly[0])  # nội dung gốc có thể chứa [[...]]
    #     # chuẩn hóa <br/>
    #     raw_html = re.sub(r'<br\s*/?>', '<br/>', raw_html)
    #     # tìm mọi biểu thức [[...]] trong raw_html (giữ nguyên nội dung giữa [[ ]])
    #     found_answers = re.findall(r'\[\[(.*?)\]\]', raw_html, flags=re.DOTALL)
    #     # trim từng answer
    #     dap_an_dt = [a.strip() for a in found_answers if a.strip()]
    #     # ===== 4. Loại bỏ các dòng tiêu đề / "Đáp án:" và loại bỏ [[...]] khỏi nội dung hiển thị =====
    #     # Tách theo <br/> để giữ cấu trúc giống trước
    #     lines = [ln.strip() for ln in raw_html.split('<br/>')]
    #     filtered = []
    #     for ln in lines:
    #         # if not ln:
    #         #     continue
    #         # # bỏ các dòng bắt đầu bằng tiêu đề hoặc "Đáp án" (các dạng có thể xuất hiện)
    #         # if ln.startswith("Điền đáp án") or ln.startswith("Đáp án") or ln.startswith("Đáp án:"):
    #         #     continue
    #         # # loại bỏ mọi [[...]] còn lại
    #         # ln_clean = re.sub(r'\[\[.*?\]\]', '', ln)
    #         # ln_clean = ln_clean.strip()
    #         # if ln_clean:
    #         #     filtered.append(ln_clean)
    #         if not ln:
    #             continue
    #             # Dùng BeautifulSoup để lấy nội dung văn bản thuần, bỏ thẻ HTML
    #         text_content = BeautifulSoup(ln, 'html.parser').get_text().strip()
    #             # bỏ các dòng bắt đầu bằng tiêu đề hoặc "Đáp án" (các dạng có thể xuất hiện)
    #         if text_content.startswith("Điền đáp án") or text_content.startswith("Đáp án") or text_content.startswith("Đáp án:"):
    #                 continue
    #             # loại bỏ mọi [[...]] còn lại
    #         ln_clean = re.sub(r'\[\[.*?\]\]', '', ln)
    #         ln_clean = ln_clean.strip()
    #         if ln_clean:
    #                 filtered.append(ln_clean)
    #     # ===== 5. Dựng phần contentquestion (title + content + answer-input) =====
    #     title_html = '<div class="title">Điền đáp án thích hợp vào ô trống (chỉ sử dụng chữ số, dấu \",\" và dấu \"-\")</div>'
    #     content_block = '<div class="content">' + '<br/>'.join(filtered) + '</div>'
    #     answer_input_html = (
    #         '<div class="answer-input">'
    #         '<div class="line">Đáp án: <span class="ans-span-second"></span>'
    #         '<input class="can-resize-second" type="text" id="mathplay-answer-1"/></div></div>'
    #     )
    #     full = title_html + content_block + answer_input_html
    #     SubElement(xml, 'contentquestion').text = full
    #     # ===== 6. Tạo listanswers đúng format (nếu có đáp án) =====
    #     if dap_an_dt:
    #         listanswers = SubElement(xml, 'listanswers')
    #         for i, ans in enumerate(dap_an_dt):
    #             # ans có thể là "56,3" hoặc "3" etc. giữ nguyên như người nhập
    #             answer = SubElement(listanswers, 'answer')
    #             SubElement(answer, 'index').text = str(i)
    #             SubElement(answer, 'content').text = ans    
    #             SubElement(answer, 'isanswer').text = 'TRUE'
    #         # ===== 7. explainquestion =====
    #         SubElement(xml, 'explainquestion').text = f"Đáp án đúng theo thứ tự là: {', '.join(dap_an_dt)}"
    #     else:
    #         # không có đáp án: không tạo listanswers và explainquestion
    #         pass
   
    def dang_dt(self, cau_sau_xu_ly, xml, subject):
        """
        Dạng điền đáp án (typeAnswer=5) - đúng theo mẫu chuẩn và logic GAS.
        - Contentquestion: title, content rỗng, answer-input chứa toàn bộ câu hỏi + ô input.
        - explainquestion: lấy từ lời giải thực tế (cau_sau_xu_ly[1][1:]), fallback nếu không có.
        """
        from xml.etree.ElementTree import SubElement
        import re
        from bs4 import BeautifulSoup

        # ===== 1. Meta =====
        SubElement(xml, 'typeAnswer').text = '5'
        SubElement(xml, 'typeViewContent').text = '0'
        SubElement(xml, 'template').text = '23'

        # ===== 2. Hint (nếu có) =====
        if len(cau_sau_xu_ly) > 1 and isinstance(cau_sau_xu_ly[1], list) and len(cau_sau_xu_ly[1]) > 1:
            hint_html = self.convert_b4_add(cau_sau_xu_ly[1][1])
            SubElement(xml, 'hintQuestion').text = hint_html

        # ===== 3. Xử lý nội dung câu hỏi =====
        raw_html = self.convert_b4_add(cau_sau_xu_ly[0])
        raw_html = re.sub(r'<br\s*/?>', '<br/>', raw_html)

        # Tách thành các dòng
        lines = [ln.strip() for ln in raw_html.split('<br/>')]

        # --- Xử lý tiêu đề ---
        current_title_txt = lines[0] if lines else ''
        text_only = BeautifulSoup(current_title_txt, 'html.parser').get_text().strip()

        # Loại bỏ "Câu X." nếu có
        title_clean = re.sub(r'^C[âa]u\s*\d+[\.:]\s*', '', current_title_txt, flags=re.IGNORECASE).strip()
        title_text = title_clean if len(BeautifulSoup(title_clean, 'html.parser').get_text().strip()) > 1 else ''

        # Nếu không có tiêu đề hợp lệ → dùng mặc định
        if not title_text and subject in self.subjects_with_default_titles:
            found_answers = re.findall(r'\[\[(.*?)\]\]', raw_html, flags=re.DOTALL)
            dap_an_dt = [a.strip() for a in found_answers if a.strip()]
            all_ans = ''.join(dap_an_dt)
            if any(c.isalpha() for c in all_ans):
                title_text = 'Điền đáp án thích hợp vào ô trống'
            else:
                title_text = 'Điền đáp án thích hợp vào ô trống (chỉ sử dụng chữ số, dấu ",", và dấu "-")'

        # --- Tạo div.title ---
        title_div = f'<div class="title">{title_text}</div>' if title_text else ''

        # ===== 4. Trích đáp án và thay bằng ô input =====
        input_index = 0

        def replace_with_input(match):
            nonlocal input_index
            input_index += 1
            return (
                f'<span class="ans-span-second"></span>'
                f'<input class="can-resize-second" type="text" id="mathplay-answer-{input_index}">'
            )

        # Thay toàn bộ [[...]] bằng ô input
        processed_content = re.sub(r'\[\[.*?\]\]', replace_with_input, raw_html, flags=re.DOTALL)

        # Loại bỏ dòng đầu (đã dùng làm title)
        content_lines = lines[1:] if len(lines) > 1 else []
        full_content_after_title = '<br/>'.join(content_lines)

        # Thay lại [[...]] trong phần này
        final_line_content = re.sub(r'\[\[.*?\]\]', replace_with_input, full_content_after_title, flags=re.DOTALL)

        # ===== 5. Xây dựng answer-input =====
        answer_input_block = f'<div class="answer-input"><div class="line">{final_line_content}</div></div>'

        # ===== 6. contentquestion =====
        content_block = '<div class="content"></div>'
        full_content = title_div + content_block + answer_input_block
        SubElement(xml, 'contentquestion').text = full_content

        # ===== 7. listanswers =====
        found_answers = re.findall(r'\[\[(.*?)\]\]', raw_html, flags=re.DOTALL)
        dap_an_dt = [a.strip() for a in found_answers if a.strip()]

        if dap_an_dt:
            listanswers = SubElement(xml, 'listanswers')
            for i, ans in enumerate(dap_an_dt):
                cleaned_answer = ans.replace('‘', "'").replace('’', "'").replace('|', '[-]')
                answer = SubElement(listanswers, 'answer')
                SubElement(answer, 'index').text = str(i)
                SubElement(answer, 'content').text = cleaned_answer
                SubElement(answer, 'isanswer').text = 'TRUE'

            # ===== 8. explainquestion: LẤY TRỰC TIẾP TỪ LỜI GIẢI TRONG DOCX =====
           # ===== 8. explainquestion =====
            huong_dan_giai_html = self.convert_b4_add(cau_sau_xu_ly[1][0])

            # Loại bỏ #### đầu dòng
            huong_dan_giai_html = re.sub(r'^\s*#+\s*', '', huong_dan_giai_html)

            # Thêm kiểm tra nội dung trống
            plain_hdg = BeautifulSoup(huong_dan_giai_html, 'html.parser').get_text().strip()

            if len(plain_hdg) > 4:
                SubElement(xml, 'explainquestion').text = huong_dan_giai_html
            else:
                SubElement(xml, 'explainquestion').text = f"Đáp án đúng theo thứ tự là: {', '.join(dap_an_dt)}"

        else:
            # ❌ Không có đáp án → explainquestion rỗng hoặc có thể để mặc định
            SubElement(xml, 'explainquestion').text = ''
            SubElement(xml, 'listanswers')  # tạo rỗng

    def dang_tl(self, cau_sau_xu_ly, xml, audio):
            """Xử lý dạng Tự luận, giữ table/ảnh trong content và HDG"""
            SubElement(xml, 'typeAnswer').text = '3'
            SubElement(xml, 'typeViewContent').text = '0'
            SubElement(xml, 'template').text = '0'
            # Content
            content_html = self.convert_content_to_html(cau_sau_xu_ly[0])
            if audio and len(audio[0]) > 8:
                link = audio[0].replace('Audio:', '').strip()
                content_html += f'<audio controls=""><source src="{link}" type="audio/mpeg"></audio>'
            SubElement(xml, 'contentquestion').text = content_html
            # List answers placeholder
            listanswers = SubElement(xml, 'listanswers')
            answer = SubElement(listanswers, 'answer')
            SubElement(answer, 'index').text = '0'
            SubElement(answer, 'content').text = 'REPLACELATER'
            SubElement(answer, 'isanswer').text = 'TRUE'
            # HDG
            hdg_html = self.convert_content_to_html(cau_sau_xu_ly[1]) if len(cau_sau_xu_ly) > 1 else ''
            SubElement(xml, 'explainquestion').text = hdg_html

    def convert_b4_add(self, paragraphs):
        """Xử lý danh sách paragraph thành HTML (giống GAS ConvertB4Add)"""
        string_content = '<p>'
        for index, paragraph in enumerate(paragraphs):
            new_children = []
            # if paragraph._element.tag.endswith('tbl'):
            #     html_table = self.convert_table_to_html(paragraph)
            #     new_children.append(html_table)
            if isinstance(paragraph, Table):
                html_table = self.convert_table_to_html(paragraph)
                new_children.append(html_table)
            else:
                self.convert_normal_paras(paragraph, index, new_children)
            new_content = "".join(new_children)
            string_content += f"{new_content}<br>"
        # string_content += "</div>"
        string_content += "</p>"
        # Xử lý math-latex: $...$
        import re
        math_latex = re.compile(r"\$[^$]*\$")
        string_content = math_latex.sub(lambda m: f' <span class="math-tex">{m.group()}</span>', string_content)
        return string_content

    def convert_normal_paras(self, paragraph: Paragraph, index, new_children: list):
        """Chuyển 1 paragraph sang HTML, bỏ phần đầu (Câu, HL, A/B/C/D) và giữ format,
        xử lý cả trường hợp các phần đó bị chia nhỏ qua nhiều run."""
        import re
        # ✅ Gom từng run để dò pattern, kể cả khi chia nhỏ
        progressive_text = ""
        content_start_pos = 0
        detected = False
        patterns = []
        if index == 0:
            patterns.append(r"^C[âa]u\s*\d+[\.:]\s*")  # Câu 1:
        patterns.append(r"^HL:\s*")
        patterns.append(r"^([A-D])\.\s*")
        # Dò dần theo run
        for run in paragraph.runs:
            if detected:
                break
            full_text = run.text or ""
            progressive_text += full_text
            for pat in patterns:
                m = re.match(pat, progressive_text, re.IGNORECASE)
                if m:
                    content_start_pos = m.end()
                    detected = True
                    break
        # ✅ Sau khi có content_start_pos, xử lý như cũ
        html_content = ""
        prev_style = None
        buffer = ""
        current_text_pos = 0
        for run in paragraph.runs:
            full_text = run.text or ""
            text_start = current_text_pos
            text_end = current_text_pos + len(full_text)
            if text_end <= content_start_pos:
                current_text_pos = text_end
                continue
            if text_start < content_start_pos:
                slice_start = content_start_pos - text_start
                segment_text = full_text[slice_start:]
            else:
                segment_text = full_text
            style = (
                bool(run.bold),
                bool(run.italic),
                bool(run.underline),
                bool(getattr(run.font, 'superscript', False)),
                bool(getattr(run.font, 'subscript', False)),
            )
            if prev_style is not None and style != prev_style:
                html_content += self.wrap_style(self.escape_html(buffer), prev_style)
                buffer = ""
            buffer += segment_text
            prev_style = style
            current_text_pos = text_end
        if buffer:
            html_content += self.wrap_style(self.escape_html(buffer), prev_style)
        # ✅ Giữ logic thêm ảnh cũ
        for run in paragraph.runs:
            blips = run._element.xpath(
                './/*[local-name()="blip" and namespace-uri()="http://schemas.openxmlformats.org/drawingml/2006/main"]'
            )
            if blips:
                try:
                    rId_nodes = run._element.xpath(
                        './/*[local-name()="blip"]/@*[local-name()="embed"]'
                    )
                    if rId_nodes:
                        rId = rId_nodes[0]
                        img_tag = self._make_img_tag_from_rid(rId)
                        if img_tag:
                            html_content += img_tag
                except Exception:
                    pass
        new_children.append(html_content.strip())

    def escape_html(self, text):
        """Escape HTML entities"""
        return (text
            .replace('&', '&amp;')
            .replace('<', '<')
            .replace('>', '>')
            .replace('"', '&quot;')
            .replace("'", '&#039;'))

    def prettify_xml(self, elem):
        """Tạo XML đẹp với indentation"""
        rough_string = tostring(elem, encoding='utf-8')
        reparsed = minidom.parseString(rough_string)
        return reparsed.toprettyxml(indent="  ", encoding='UTF-8').decode('utf-8')


    def post_process_xml(self, xml_str):
        """
        Sửa lại hàm post_process_xml:
        - Di chuyển second_correction ra khỏi vòng lặp đầu
        - Thay đổi cách xử lý math-tex để lấy nội dung bên trong span
        - Thêm các regex để unescape các thẻ có attribute như <table class='...'>
        - Một số sửa nhỏ khác để tránh phá hỏng XML quá sớm
        """
        import re
        from xml.dom import minidom
        import html

        # đảm bảo header
        xml_str = xml_str.replace('<?xml version="1.0" ?>', '<?xml version="1.0" encoding="UTF-8"?>')

        # các thay thế cố định (dùng re.escape khi cần)
        correction = {
            'REPLACELATER': '',
            '&lt;br&gt;': '<br>',
            '&lt;br/&gt;': '<br/>',
            '&lt;em&gt;': '<em>',
            '&lt;/em&gt;': '</em>',
            '&lt;u&gt;': '<u>',
            '&lt;/u&gt;': '</u>',
            '&lt;strong&gt;': '<strong>',
            '&lt;/strong&gt;': '</strong>',
            '&lt;/font&gt;': '</font>',
            '&lt;font': '<font',
            '&lt;span': '<span',
            '&lt;/span&gt;': '</span>',
            '&lt;input': '<input',
            '"&gt;': '">',
            '&lt;/div&gt;': '</div>',
            '&lt;div': '<div',
            '&#xD;': '',
            '&lt;label': '<label',
            '&lt;select': '<select',
            '&lt;option': '<option',
            'hidden&gt;': 'hidden>',
            '&lt;/option&gt;': '</option>',
            '&lt;/select&gt;': '</select>',
            '&lt;/label&gt;': '</label>',
            '&quot;': '"',
            '&lt;center&gt;': '<center>',
            '&lt;/center&gt;': '</center>',
            '&lt;p&gt;': '<p>',
            '&lt;/p&gt;': '</p>',
            '&lt;img': '<img',
            ' /&gt;': ' />',
            '/&gt;': '/>',
            '&lt;audio': '<audio',
            '&lt;/audio&gt;': '</audio>',
            '&lt;source': '<source',
            '&lt;blockquote&gt;': '<blockquote>',
            '&lt;/blockquote&gt;': '</blockquote>',
            '&lt;table&gt;': '<table>',
            '&lt;/table&gt;': '</table>',
            '&lt;tr&gt;': '<tr>',
            '&lt;/tr&gt;': '</tr>',
            '&lt;td&gt;': '<td>',
            '&lt;/td&gt;': '</td>',
            '&lt;li&gt;': '<li>',
            '&lt;/li&gt;': '</li>',
            '&lt;i&gt;': '<i>',
            '&lt;/i&gt;': '</i>',
            '&lt;sub&gt;': '<sub>',
            '&lt;/sub&gt;': '</sub>',
            '&lt;sup&gt;': '<sup>',
            '&lt;/sup&gt;': '</sup>',
        }

        # first pass of simple replacements
        for key, val in correction.items():
            xml_str = re.sub(re.escape(key), val, xml_str, flags=re.IGNORECASE)

        # second set of corrections (ensure it is NOT nested inside the previous loop)
        second_correction = {
            '&lt;i&gt;': '<i>',
            '&lt;/i&gt;': '</i>',
            '&lt;u&gt;': '<u>',
            '&lt;/u&gt;': '</u>',
            '&lt;strong&gt;': '<strong>',
            '&lt;/strong&gt;': '</strong>',
            '&lt;sub&gt;': '<sub>',
            '&lt;/sub&gt;': '</sub>',
            '&lt;sup&gt;': '<sup>',
            '&lt;/sup&gt;': '</sup>',
        }
        for key, val in second_correction.items():
            xml_str = re.sub(re.escape(key), val, xml_str, flags=re.IGNORECASE)

        # === XỬ LÝ CÁC THẺ CÓ ATTR (ví dụ: &lt;table class='...'&gt;) ===
        tags_with_attrs = [
            'table', 'tr', 'td', 'th', 'tbody', 'thead', 'tfoot',
            'img', 'div', 'span', 'p', 'sup', 'sub', 'input', 'label',
            'select', 'option', 'audio', 'source', 'blockquote', 'li', 'center', 'font'
        ]
        for tag in tags_with_attrs:
            xml_str = re.sub(r'&lt;(' + tag + r'\b)', r'<\1', xml_str, flags=re.IGNORECASE)
            xml_str = re.sub(r'&lt;\/(' + tag + r')\s*&gt;', r'</\1>', xml_str, flags=re.IGNORECASE)

        # chuyển các thực thể HTML phổ biến sang ký tự thật (an toàn hơn là unescape toàn bộ)
        xml_str = html.unescape(xml_str)

        # === XỬ LÝ MATHLATEX ===
        def clean_mathlatex(match):
            inner = match.group(1)
            inner = (
                inner
                .replace('<strong>', '')
                .replace('</strong>', '')
                .replace('<i>', '')
                .replace('</i>', '')
                .replace('<u>', '')
                .replace('</u>', '')
                .replace('<br>', '')
                .replace('<br/>', '')
                .replace('%', '\\%')
                .replace('\\frac', '\\dfrac')
            )
            return inner

        xml_str = re.sub(
            r'<span\s+class=["\']math-tex["\']\s*>(.*?)</span>',
            clean_mathlatex,
            xml_str,
            flags=re.DOTALL | re.IGNORECASE
        )

        # === LÀM ĐẸP LẠI XML ===
        try:
            xml_str = minidom.parseString(xml_str.encode('utf-8')).toprettyxml(indent="  ", encoding="UTF-8").decode("utf-8")
        except Exception:
            pass

        # === LƯU FILE ===
        file_name = "docXML.xml"
        if "<itemDocuments>" in xml_str:
            file_name = "docHL.xml"
        try:
            with open(file_name, "w", encoding="utf-8") as f:
                f.write(xml_str)
        except Exception:
            pass

        return xml_str
