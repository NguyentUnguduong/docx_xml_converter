"""
Module xử lý chuyển đổi DOCX sang XML
Dựa trên logic từ Google Apps Script
"""

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
from tinhoc_processor import TinHocProcessor
from typing import List, Union, Any, Iterable, Optional
import traceback
from PIL import Image
from io import BytesIO

class DocxProcessor:
    """Class chính xử lý DOCX"""
    
    def __init__(self):
        self.subjects_with_default_titles = [
            "TOANTHPT", "VATLITHPT2", "HOATHPT2", "SINHTHPT2",
            "LICHSUTHPT", "DIALITHPT", "GDCDTHPT2", "NGUVANTHPT",
            "TOANTHCS2", "KHTN", "KHXHTHCS", "GDCDTHCS2", "NGUVANTHCS2"
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
        """Xử lý file DOCX và trả về XML string"""
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
                    raise ValueError(f"Sai format header: {text}")
                
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
                self.format_questions(group, root)
        
        # Convert sang string
        xml_str = self.prettify_xml(root)
        xml_str = self.post_process_xml(xml_str)
        
        return xml_str
    
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
            self.format_questions(group, list_question)
        
        return item_doc
    
    def xu_ly_hl(self, content):
        """
        Xử lý nội dung học liệu (HL) thành HTML hoàn chỉnh.
        - Hỗ trợ Paragraph (bold/italic/underline/sub/sup)
        - Hỗ trợ Ảnh (DrawingML / VML)
        - Hỗ trợ Bảng (bao gồm nested tables)
        - Chạy được với cả Document, _Body hoặc list phần tử
        """
        import traceback
        from docx.text.paragraph import Paragraph
        from docx.table import Table as DocxTable
        from docx.oxml.text.paragraph import CT_P
        from docx.oxml.table import CT_Tbl

        print("[DEBUG] === BẮT ĐẦU HÀM xu_ly_hl ===")

        # =================== HELPER: EXTRACT ELEMENTS ===================
        def extract_elements(container):
            """
            Lấy tất cả phần tử (Paragraph + Table) từ Document hoặc Body.
            """
            elements = []
            
            # Thử dùng thuộc tính tables và paragraphs trước
            try:
                if hasattr(container, 'paragraphs') and hasattr(container, 'tables'):
                    # Lấy tất cả paragraphs và tables
                    paragraphs = list(container.paragraphs)
                    tables = list(container.tables)
                    
                    # Tạo dictionary để sắp xếp theo thứ tự xuất hiện
                    elements_dict = {}
                    
                    # Thêm paragraphs
                    for p in paragraphs:
                        try:
                            # Lấy vị trí của paragraph trong document
                            p_elem = p._element
                            parent = p_elem.getparent()
                            if parent is not None:
                                index = list(parent).index(p_elem)
                                elements_dict[(parent, index)] = p
                        except:
                            elements.append(p)
                    
                    # Thêm tables
                    for t in tables:
                        try:
                            # Lấy vị trí của table trong document
                            t_elem = t._element
                            parent = t_elem.getparent()
                            if parent is not None:
                                index = list(parent).index(t_elem)
                                elements_dict[(parent, index)] = t
                        except:
                            elements.append(t)
                    
                    # Sắp xếp theo thứ tự xuất hiện
                    sorted_items = sorted(elements_dict.items(), key=lambda x: x[0][1])
                    elements.extend([item[1] for item in sorted_items])
                    
                    print(f"[DEBUG] Trích xuất được {len(paragraphs)} paragraphs và {len(tables)} tables")
                    return elements
            except Exception as e:
                print(f"[WARN] Không thể dùng phương pháp tables/paragraphs: {e}")
            
            # Fallback: dùng phương pháp cũ
            for child in container._element:
                if isinstance(child, CT_P):
                    elements.append(Paragraph(child, container))
                elif isinstance(child, CT_Tbl):
                    elements.append(DocxTable(child, container))
            
            return elements

        # =================== HELPER: CONVERT PARAGRAPH ===================
        def convert_paragraph_to_html(p: Paragraph) -> str:
            html = ""
            try:
                runs = p.runs
                print(f"[DEBUG] → Paragraph có {len(runs)} runs")
                for i, run in enumerate(runs):
                    text = run.text or ""
                    # 1) Luôn kiểm tra ảnh trong run
                    try:
                        imgs = self._get_image_tags_from_run(run)
                        if imgs:
                            print(f"[DEBUG]   Run {i}: tìm thấy {len(imgs)} ảnh trong run")
                            for it in imgs:
                                html += it
                    except Exception as e:
                        print(f"[WARN] Lỗi khi lấy ảnh từ run {i}: {e}")

                    # 2) Nếu text rỗng — không xử lý style, bỏ qua
                    if text == "":
                        print(f"[DEBUG]   Run {i}: text rỗng, bỏ qua format")
                        continue

                    print(f"[DEBUG]   Run {i}: {text!r}")

                    # Apply format chỉ khi có text
                    pieces = text
                    if run.bold:
                        pieces = f"<b>{pieces}</b>"
                        print(f"[DEBUG]     bold")
                    if run.italic:
                        pieces = f"<i>{pieces}</i>"
                        print(f"[DEBUG]     italic")
                    if run.underline:
                        pieces = f"<u>{pieces}</u>"
                        print(f"[DEBUG]     underline")
                    if getattr(run.font, "subscript", False):
                        pieces = f"<sub>{pieces}</sub>"
                        print(f"[DEBUG]     subscript")
                    if getattr(run.font, "superscript", False):
                        pieces = f"<sup>{pieces}</sup>"
                        print(f"[DEBUG]     superscript")

                    # Escape HTML entities
                    try:
                        pieces = self.escape_html(pieces)
                    except Exception:
                        pieces = (pieces
                                .replace('&', '&amp;')
                                .replace('<', '&lt;')
                                .replace('>', '&gt;'))

                    html += pieces

                # Ảnh inline / outside runs
                try:
                    inline_imgs = p._element.xpath(".//a:blip/@r:embed")
                    print(f"[DEBUG]   Phát hiện {len(inline_imgs)} ảnh inline trong paragraph")
                    for rId in inline_imgs:
                        tag = self._make_img_tag_from_rid(rId)
                        if tag:
                            html += tag
                            print(f"[DEBUG]   Ảnh rId={rId} đã được xử lý (inline)")
                        else:
                            print(f"[DEBUG]   Không tìm thấy part cho rId={rId} (inline)")
                except Exception as e:
                    print(f"[WARN] Lỗi khi xử lý ảnh inline: {e}")

            except Exception as e:
                print(f"[ERROR] Lỗi convert_paragraph_to_html: {e}")
                traceback.print_exc()

            return f"<p>{html}</p>"

    
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
        Convert table sang HTML (hỗ trợ nested table, ảnh trong ô, colspan).
        NOTE: Rowspan chính xác yêu cầu build matrix toàn bộ bảng; hiện tại chỉ xử lý colspan.
        """
        html = "<table class='table-material-question'>"

        # Duyệt từng row theo python-docx
        for row in table.rows:
            html += "<tr>"
            for cell in row.cells:
                # --- Tính colspan (gridSpan) bằng XPath (an toàn hơn)
                try:
                    # tìm giá trị w:gridSpan/@w:val trong tcPr nếu có
                    grid_span_vals = cell._tc.xpath(".//w:gridSpan/@w:val", namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
                    colspan = int(grid_span_vals[0]) if grid_span_vals else 1
                except Exception:
                    colspan = 1

                # --- Tạo nội dung ô giữ đúng thứ tự paragraph + nested table ---
                parts: List[str] = []
                for child in cell._element:
                    if isinstance(child, CT_P):
                        # Paragraph native -> wrap bằng Paragraph object
                        try:
                            p = DocxParagraph(child, cell)
                            parts.append(self.convert_content_to_html(p))
                        except Exception:
                            # fallback: text raw
                            try:
                                parts.append("".join(run.text for run in cell.paragraphs[0].runs))
                            except Exception:
                                parts.append("")
                    elif isinstance(child, CT_Tbl):
                        try:
                            nested = DocxTable(child, cell)  # tạo object Table từ oxml
                            parts.append(self.convert_table_to_html(nested))
                        except Exception:
                            parts.append("")

                cell_html = "".join(parts).strip()
                if not cell_html:
                    cell_html = "&nbsp;"

                td_attrs = ""
                if colspan > 1:
                    td_attrs = f' colspan="{colspan}"'

                html += f"<td{td_attrs}>{cell_html}</td>"
            html += "</tr>"

        html += "</table><br>"
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
    
   
    
    def format_questions(self, group, questions_xml):
        """Format các câu hỏi"""
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
        for question_dict in group_of_q:
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
            self.protocol_of_q(question_dict['items'], each_question_xml, group['subject'])
            
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
    #     """
    #     Dùng rId để lấy image part từ self.doc.part.related_parts,
    #     trả về một thẻ <img src="data:..."> hoặc None.
    #     """
    #     try:
    #         # related_parts: mapping rId -> Part (chứa .blob và .content_type)
    #         part = self.doc.part.related_parts.get(rId)
    #         if not part:
    #             # có thể relationship nằm trong phụ part (ví dụ trong headers/footers),
    #             # thử tìm mọi part trong document (an toàn hơn)
    #             for rel in self.doc.part.rels.values():
    #                 try:
    #                     target = getattr(rel, 'target_part', None)
    #                     if target and getattr(target, 'reltype', None) and 'image' in getattr(target, 'content_type', ''):
    #                         if rel.rId == rId:
    #                             part = target
    #                             break
    #                 except Exception:
    #                     continue

    #         if not part:
    #             # không tìm thấy image part
    #             return None

    #         img_bytes = part.blob
    #         content_type = getattr(part, 'content_type', 'image/png')
    #         # encode base64
    #         b64 = base64.b64encode(img_bytes).decode('ascii')
    #         # hardcode width và height
    #         style = 'style="width:321px;height:214px;"'
    #         return f'<center><img src="data:{content_type};base64,{b64}" {style} /></center>'
    #     except Exception:
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

            # --- Đọc kích thước gốc ---
            try:
                img = Image.open(BytesIO(img_bytes))
                width, height = img.size
                print(f"[DEBUG] Ảnh rId={rId} size: {width}x{height}")
            except Exception as e:
                print(f"[WARN] Không đọc được kích thước ảnh: {e}")
                width, height = 300, 200  # fallback

            # encode base64
            b64 = base64.b64encode(img_bytes).decode('ascii')

            # --- Sinh tag HTML ---
            style = f'style="max-width:{width}px; height:auto;"'
            # hoặc nếu muốn cố định tỉ lệ: style = f'style="width:{width}px;height:{height}px;"'

            return f'<center><img src="data:{content_type};base64,{b64}" {style} /></center>'

        except Exception as e:
            print(f"[ERROR] _make_img_tag_from_rid lỗi: {e}")
            import traceback; traceback.print_exc()
            return None
        
    def protocol_of_q(self, question, each_question_xml, subject):
        """Phân tích cấu trúc câu hỏi"""
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
            raise ValueError(f"Thiếu 'Lời giải' trong câu: {question[0].text[:50]}")
        
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
                raise ValueError(f"HDG chỉ được có 1 link TTS: {link_speech_explain}")
            
            if link_speech_explain[0].endswith(('.mp3', '.mp4')):
                SubElement(each_question_xml, 'urlSpeechExplain').text = link_speech_explain[0]
        
        # Xác định dạng câu hỏi
        answer = thanh_phan_hdg[0][0].text.strip() if thanh_phan_hdg[0] else ''
        
        cau_sau_xu_ly = [thanh_phan_cau_hoi, thanh_phan_hdg]
        audio = [link for link in link_cau_hoi if 'Audio:' in link]
        
        # Routing theo subject
        if self.is_tinhoc_subject(subject):
            self.route_to_tinhoc_module(cau_sau_xu_ly, each_question_xml, audio, answer, subject)
        else:
            self.route_to_default_module(cau_sau_xu_ly, each_question_xml, audio, answer, subject)
    
    def is_tinhoc_subject(self, subject):
        """Kiểm tra có phải môn tin học không"""
        return any(subject.startswith(tinhoc) for tinhoc in self.tinhoc_subjects)
    
    def route_to_tinhoc_module(self, cau_sau_xu_ly, xml, audio, answer, subject):
        """Xử lý cho môn Tin học"""
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
    
    def route_to_default_module(self, cau_sau_xu_ly, xml, audio, answer, subject):
        """Xử lý cho môn thông thường"""
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
                    raise ValueError(f"Chỉ được 1 link TTS: {link}")
                SubElement(xml, 'urlSpeechContent').text = link
                one_tts = True
            else:
                if one_media:
                    raise ValueError(f"Chỉ được 1 link Video: {link}")
                
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
    
    def hdg_tn(self, array_hdg, xml):
        """
        Hướng dẫn giải TN, giữ HTML (ảnh/table)
        - Giữ logic cũ phát hiện đáp án
        - Tự động bỏ dòng "Đáp án đúng là...", số đáp án đầu dòng (1,2,3,...) hoặc chữ (A,B,C,D)
        - Bỏ luôn tiền tố "Giải thích:" nếu có
        """
        import re
        from xml.etree.ElementTree import SubElement

        if not array_hdg:
            return

        # Xóa thẻ explainquestion cũ nếu có
        existing_explain = xml.find('explainquestion')
        if existing_explain is not None:
            xml.remove(existing_explain)

        explain_text = ''
        answer_letters = ['A', 'B', 'C', 'D']

        # ===== 1️⃣ Tìm đáp án đúng từ phần hướng dẫn =====
        index_answer = []
        hdg_raw = ''

        if isinstance(array_hdg, list):
            for part in array_hdg:
                if hasattr(part, "text"):
                    hdg_raw += part.text.strip() + " "
                elif isinstance(part, list):
                    for p in part:
                        if hasattr(p, "text"):
                            hdg_raw += p.text.strip() + " "

        # Tìm đáp án
        index_answer = [int(ch) for ch in re.findall(r'\d+', hdg_raw)]
        if index_answer:
            dap_an = ' '.join(answer_letters[i - 1] for i in index_answer if 1 <= i <= len(answer_letters))
            explain_text = f"Đáp án đúng là: {dap_an}"
        else:
            match = re.search(r"([A-D])", hdg_raw, re.IGNORECASE)
            if match:
                explain_text = f"Đáp án đúng là: {match.group(1).upper()}"

        # ===== 2️⃣ Nếu có nội dung hướng dẫn thực sự (giải thích chi tiết)
        hdg_html = self.convert_content_to_html(array_hdg)
        plain = re.sub(r'<[^>]+>', '', hdg_html).strip()

        if len(plain) > 4:
            # Nếu có giải thích thật → bỏ phần đáp án và tiền tố "Giải thích:"
            explain_text = hdg_html.strip()

            # --- Xóa phần đáp án đầu đoạn (dạng "1", "2", "A", "B"...) ---
            explain_text = re.sub(
                r'^\s*(\d+|[A-Da-d])\s*(<br\s*/?>|:|\.|,)?\s*',
                '',
                explain_text,
                flags=re.IGNORECASE
            )

            # --- Bỏ tiền tố "Giải thích:" hoặc "Giải thích<br>" ---
            explain_text = re.sub(
                r'^\s*Giải\s*thích\s*[:：]?\s*(<br\s*/?>)?',
                '',
                explain_text,
                flags=re.IGNORECASE
            ).strip()

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
            ans_html = re.sub(
                r'^(<strong>|<b>)?\s*([A-Da-d])[\.\)]\s*(</strong>|</b>)?',
                '',
                ans_html
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
    

    def dang_dt(self, cau_sau_xu_ly, xml, subject):
        """
        Dạng điền đáp án (typeAnswer=5) - rút gọn, không dùng normalize/unescape.
        Tìm đáp án trực tiếp từ [[...]] rồi xây XML đúng format (contentquestion, listanswers, explainquestion).
        """
        # ===== 1. Meta =====
        SubElement(xml, 'typeAnswer').text = '5'
        SubElement(xml, 'typeViewContent').text = '0'
        SubElement(xml, 'template').text = '23'

        # ===== 2. Hint (nếu có) =====
        if len(cau_sau_xu_ly) > 1 and isinstance(cau_sau_xu_ly[1], list) and len(cau_sau_xu_ly[1]) > 1:
            hint_html = self.convert_b4_add(cau_sau_xu_ly[1][1])
            SubElement(xml, 'hintQuestion').text = hint_html

        # ===== 3. Lấy nội dung gốc và tìm đáp án [[...]] từ đó =====
        raw_html = self.convert_b4_add(cau_sau_xu_ly[0])  # nội dung gốc có thể chứa [[...]]
        # chuẩn hóa <br/>
        raw_html = re.sub(r'<br\s*/?>', '<br/>', raw_html)

        # tìm mọi biểu thức [[...]] trong raw_html (giữ nguyên nội dung giữa [[ ]])
        found_answers = re.findall(r'\[\[(.*?)\]\]', raw_html, flags=re.DOTALL)
        # trim từng answer
        dap_an_dt = [a.strip() for a in found_answers if a.strip()]

        # ===== 4. Loại bỏ các dòng tiêu đề / "Đáp án:" và loại bỏ [[...]] khỏi nội dung hiển thị =====
        # Tách theo <br/> để giữ cấu trúc giống trước
        lines = [ln.strip() for ln in raw_html.split('<br/>')]

        filtered = []
        for ln in lines:
            if not ln:
                continue
            # bỏ các dòng bắt đầu bằng tiêu đề hoặc "Đáp án" (các dạng có thể xuất hiện)
            if ln.startswith("Điền đáp án") or ln.startswith("Đáp án") or ln.startswith("Đáp án:"):
                continue
            # loại bỏ mọi [[...]] còn lại
            ln_clean = re.sub(r'\[\[.*?\]\]', '', ln)
            ln_clean = ln_clean.strip()
            if ln_clean:
                filtered.append(ln_clean)

        # ===== 5. Dựng phần contentquestion (title + content + answer-input) =====
        title_html = '<div class="title">Điền đáp án thích hợp vào ô trống (chỉ sử dụng chữ số, dấu \",\" và dấu \"-\")</div>'
        content_block = '<div class="content">' + '<br/>'.join(filtered) + '</div>'
        answer_input_html = (
            '<div class="answer-input">'
            '<div class="line">Đáp án: <span class="ans-span-second"></span>'
            '<input class="can-resize-second" type="text" id="mathplay-answer-1"/></div></div>'
        )

        full = title_html + content_block + answer_input_html
        SubElement(xml, 'contentquestion').text = full

        # ===== 6. Tạo listanswers đúng format (nếu có đáp án) =====
        if dap_an_dt:
            listanswers = SubElement(xml, 'listanswers')
            for i, ans in enumerate(dap_an_dt):
                # ans có thể là "56,3" hoặc "3" etc. giữ nguyên như người nhập
                answer = SubElement(listanswers, 'answer')
                SubElement(answer, 'index').text = str(i)
                SubElement(answer, 'content').text = ans
                SubElement(answer, 'isanswer').text = 'TRUE'

            # ===== 7. explainquestion =====
            SubElement(xml, 'explainquestion').text = f"Đáp án đúng theo thứ tự là: {', '.join(dap_an_dt)}"
        else:
            # không có đáp án: không tạo listanswers và explainquestion
            pass



                
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
        # string_content = ""
        # for index, paragraph in enumerate(paragraphs):
        #     new_children = []

        #     if paragraph._element.tag.endswith('tbl'):
        #         html_table = self.convert_table_to_html(paragraph)
        #         new_children.append(html_table)
        #     else:
        #         self.convert_normal_paras(paragraph, index, new_children)

        #     new_content = "".join(new_children)
        #     if len(paragraphs) > 1:
        #         string_content += f"{new_content}<br>"
        #     else:
        #         string_content += new_content
        # string_content = '<div class="content">'
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
    
    def convert_normal_paras(self, paragraph, index, new_children):
        """Chuyển 1 paragraph sang HTML, bỏ phần đầu (Câu, HL, A/B/C/D) và giữ format,
        bây giờ hỗ trợ cả sup và sub giống convert_paragraph_to_html()"""
        raw_text = "".join(run.text for run in paragraph.runs or [])

        # ✅ Bước 1: Xác định vị trí bắt đầu thực sự của nội dung
        import re
        content_start_pos = 0
        if index == 0:
            cau_match = re.match(r"^C[âa]u\s*\d+[\.:]\s*", raw_text, re.IGNORECASE)
            if cau_match:
                content_start_pos = cau_match.end()

        hl_match = re.match(r"^HL:\s*", raw_text[content_start_pos:], re.IGNORECASE)
        if hl_match:
            content_start_pos += hl_match.end()

        answer_match = re.match(r"^([A-D])\.\s*", raw_text[content_start_pos:], re.IGNORECASE)
        if answer_match:
            content_start_pos += answer_match.end()

        html_content = ""
        prev_style = None
        buffer = ""
        current_text_pos = 0

        # Duyệt qua từng run, xử lý cắt theo content_start_pos và gom theo style (bao gồm sup/sub)
        for run in paragraph.runs:
            full_text = run.text or ""
            text_start = current_text_pos
            text_end = current_text_pos + len(full_text)

            # Nếu toàn bộ phần này nằm trước content start thì bỏ qua
            if text_end <= content_start_pos:
                current_text_pos = text_end
                continue

            # Nếu phần bắt đầu nằm trước content_start_pos thì cắt phần phía trước
            if text_start < content_start_pos:
                slice_start = content_start_pos - text_start
                segment_text = full_text[slice_start:]
            else:
                segment_text = full_text

            # Build style tuple giống convert_paragraph_to_html
            style = (
                bool(run.bold),
                bool(run.italic),
                bool(run.underline),
                bool(getattr(run.font, 'superscript', False)),
                bool(getattr(run.font, 'subscript', False)),
            )

            # Nếu khác style hiện tại -> flush buffer
            if prev_style is not None and style != prev_style:
                # dùng wrap_style để đóng/gói buffer theo prev_style
                html_content += self.wrap_style(self.escape_html(buffer), prev_style)
                buffer = ""

            buffer += segment_text
            prev_style = style
            current_text_pos = text_end

        # flush buffer cuối cùng
        if buffer:
            html_content += self.wrap_style(self.escape_html(buffer), prev_style)

        # Thêm ảnh nếu có (giữ nguyên logic cũ)
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

        # trim và append vào new_children
        new_children.append(html_content.strip())
    
    def escape_html(self, text):
        """Escape HTML entities"""
        return (text
            .replace('&', '&amp;')
            .replace('<', '&lt;')
            .replace('>', '&gt;')
            .replace('"', '&quot;')
            .replace("'", '&#039;'))
    
    def prettify_xml(self, elem):
        """Tạo XML đẹp với indentation"""

        rough_string = tostring(elem, encoding='utf-8')
        reparsed = minidom.parseString(rough_string)
        return reparsed.toprettyxml(indent="  ", encoding='UTF-8').decode('utf-8')
    
    def post_process_xml(self, xml_str):
        """
        Xử lý XML tương tự logic của hàm TaoFile(root) bên Google Apps Script
        - Chuyển đổi các ký tự HTML encode về thẻ thật
        - Làm sạch nội dung trong <span class="math-tex">
        - Giữ nguyên cấu trúc XML và format đẹp
        """

        import re
        from xml.dom import minidom

        # Đảm bảo header XML đúng
        xml_str = xml_str.replace('<?xml version="1.0" ?>', '<?xml version="1.0" encoding="UTF-8"?>')

        # === CORRECTION LẦN 1 ===
        correction = {
            'REPLACELATER': '',
            '&lt;br&gt;': '<br>',
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
            '&amp;': '&',
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
            '&lt;/li&gt;': '</li>',
            '&lt;li&gt;': '<li>',
        }

        for key, val in correction.items():
            xml_str = re.sub(re.escape(key), val, xml_str)

        # === CORRECTION LẦN 2 ===
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
            xml_str = re.sub(re.escape(key), val, xml_str)

        # === XỬ LÝ MATHLATEX ===
        def clean_mathlatex(match):
            mathlatex = match.group(0)
            # Bỏ các style trong math
            mathlatex = (
                mathlatex
                .replace('<strong>', '')
                .replace('</strong>', '')
                .replace('<i>', '')
                .replace('</i>', '')
                .replace('<u>', '')
                .replace('</u>', '')
                .replace('<br>', '')
                .replace('%', '\\%')
                .replace('\\frac', '\\dfrac')
            )
            return mathlatex

        xml_str = re.sub(
            r'<span class="math-tex">(.*?)</span>',
            clean_mathlatex,
            xml_str,
            flags=re.DOTALL
        )

        # === GỢI Ý: KHÔNG ĐỤNG ĐẾN P-TAG Ở ĐÂY ===
        # (Google Apps Script code không chỉnh <p>, nên không thêm xử lý tự động này)

        # === LÀM ĐẸP LẠI XML ===
        try:
            xml_str = minidom.parseString(xml_str).toprettyxml(indent="  ", encoding="UTF-8").decode("utf-8")
        except Exception:
            # Nếu XML lỗi cú pháp (do chứa < hoặc & không hợp lệ)
            # thì vẫn trả bản gốc, tránh crash
            pass
        xml_str = xml_str.replace('&lt;', '<').replace('&gt;', '>').replace('&amp;', '&')
        # === LƯU FILE ===
        file_name = "docXML.xml"
        if "<itemDocuments>" in xml_str:
            file_name = "docHL.xml"

        # Tùy vào hệ thống của bạn — ví dụ ghi ra thư mục output
        with open(file_name, "w", encoding="utf-8") as f:
            f.write(xml_str)

        return xml_str