import re
import base64
from typing import List, Dict, Any, Tuple, Optional
from xml.etree import ElementTree as ET
from document_element import (
    get_blob, get_bytes, get_width, get_height, get_element_type, 
    get_text, get_num_children, get_child, get_attributes, 
    get_text_attribute_indices, get_num_rows, get_row, get_num_cells, get_cell
)
from docx import Document
from docx.text.paragraph import Paragraph
from docx.text.run import Run
class TinHocProcessor:
    
    def __init__(self):
        pass
    
    def create_safe_text_node(self, tag_name: str, content: str) -> ET.Element:
        """
        Create XML element with safe HTML content.
        Preserves allowed HTML tags while escaping others.
        Special case: preserves entire <table class='table-material-question'>...</table> blocks untouched.
        """
        element = ET.Element(tag_name)
        
        if not content:
            element.text = ''
            return element

        # ================================================
        # STEP 0: Extract and protect full table blocks
        # ================================================
        table_blocks = []
        table_placeholder = '__TABLE_BLOCK_{}__'

        def protect_table_block(match):
            table_html = match.group(0)
            placeholder = table_placeholder.format(len(table_blocks))
            table_blocks.append(table_html)
            return placeholder

        # Regex để match toàn bộ <table class='table-material-question'> ... </table>
        # Dùng re.DOTALL để . match newline
        table_pattern = re.compile(
            r'<table\s+class\s*=\s*[\'"]table-material-question[\'"][^>]*>.*?</table>',
            re.IGNORECASE | re.DOTALL
        )
        content = table_pattern.sub(protect_table_block, content)

        # ================================================
        # STEP 1–5: Run original escaping logic on the rest
        # ================================================
        allowed_tags = [
            'b', 'i', 'u', 'strong', 'em', 'br', 'center', 'img',
            'sub', 'sup', 'small', 'big', 'mark'
        ]

        # STEP 1: Detect already escaped tags
        escaped_tag_pattern = re.compile(r'&lt;(\/?[a-zA-Z][a-zA-Z0-9]*)\b[^&]*?&gt;')
        escaped_tags = []
        escaped_index = 0

        def replace_escaped(match):
            nonlocal escaped_index
            placeholder = f'__ESCAPED_TAG_{escaped_index}__'
            escaped_tags.append({'placeholder': placeholder, 'original': match.group(0)})
            escaped_index += 1
            return placeholder

        processed_content = escaped_tag_pattern.sub(replace_escaped, content)

        # STEP 2: Detect tags in quotes/code
        quoted_tag_pattern = re.compile(r'"([^"]*?<[^>]+>[^"]*)"')
        quoted_tags = []
        quoted_index = 0

        def replace_quoted(match):
            nonlocal quoted_index
            inner = match.group(1)
            placeholder = f'__QUOTED_{quoted_index}__'
            escaped = inner.replace('<', '＜').replace('>', '＞')
            quoted_tags.append({'placeholder': placeholder, 'original': f'"{escaped}"'})
            quoted_index += 1
            return placeholder

        processed_content = quoted_tag_pattern.sub(replace_quoted, processed_content)

        # STEP 3: Process actual HTML tags
        html_tag_pattern = re.compile(r'<\/?([a-zA-Z][a-zA-Z0-9]*)\b[^<>]*\/?>')
        tags_to_restore = []
        tag_index = 0

        def replace_html_tag(match):
            nonlocal tag_index
            tag_name = match.group(1)
            lower_tag = tag_name.lower()
            is_allowed = lower_tag in allowed_tags
            placeholder = f'__TAG_{tag_index}__'
            tags_to_restore.append({
                'placeholder': placeholder,
                'original': match.group(0),
                'is_allowed': is_allowed
            })
            tag_index += 1
            return placeholder

        processed_content = html_tag_pattern.sub(replace_html_tag, processed_content)

        # STEP 4: Escape remaining < > characters
        safe_content = processed_content.replace('<', '＜').replace('>', '＞')

        # STEP 5: Restore in reverse order
        for tag_info in tags_to_restore:
            placeholder = tag_info['placeholder']
            original = tag_info['original']
            is_allowed = tag_info['is_allowed']
            restored = original if is_allowed else original.replace('<', '＜').replace('>', '＞')
            safe_content = safe_content.replace(placeholder, restored)

        for quoted_info in quoted_tags:
            safe_content = safe_content.replace(quoted_info['placeholder'], quoted_info['original'])

        for escaped_info in escaped_tags:
            safe_content = safe_content.replace(escaped_info['placeholder'], escaped_info['original'])

        # STEP 6: Restore table blocks (they were never escaped!)
        for i, table_html in enumerate(table_blocks):
            placeholder = table_placeholder.format(i)
            safe_content = safe_content.replace(placeholder, table_html)

        # STEP 7: Convert fullwidth to actual HTML (ONLY for allowed tags)
        fullwidth_tag_pattern = re.compile(
            r'＜(\/?(?:b|i|u|strong|em|br|center|img|sub|sup|small|big|mark))\b([^＜＞]*?)＞',
            re.IGNORECASE
        )
        safe_content = fullwidth_tag_pattern.sub(r'<\1\2>', safe_content)

        element.text = safe_content
        return element

    # ============================================
    # TN (Multiple Choice) Processing Functions
    # ============================================

    def dang_tn_tinhoc(self, cau_sau_xu_ly: List, each_question_xml: ET.Element, audio: List, doc: Document = None):
        """Process multiple choice question - Tin học"""
        # XML structure similar to normal TN
        type_ans = ET.SubElement(each_question_xml, 'typeAnswer')
        type_ans.text = '0'
        
        type_view = ET.SubElement(each_question_xml, 'typeViewContent')
        type_view.text = '0'
        
        temp = ET.SubElement(each_question_xml, 'template')
        temp.text = '0'
        
        # Process hint question if exists
        if len(cau_sau_xu_ly[1]) > 2:
            hint = self.convert_b4_add_tinhoc(cau_sau_xu_ly[1][2], doc)
            hint_question = self.create_safe_text_node('hintQuestion', hint)
            each_question_xml.append(hint_question)
        
        # Parse question content
        content_q = []
        for index, para in enumerate(cau_sau_xu_ly[0]):
            if index == 0:
                content_q.append([para])
                continue
            
            # Check for answers A. B. C. D.
            for i in range(get_num_children(para)):
                child = get_child(para, i)
                
                if get_element_type(child) == 'TEXT':
                    text = get_text(child).strip()
                    
                    # Detect answer (A. B. C. D. or a. b. c. d.)
                    if re.match(r'^[A-Da-d]\.', text):
                        if len(content_q) == 1:
                            content_q.append([[para]])
                        else:
                            content_q[1].append([para])
                        break
            else:
                # Add paragraph to content
                if len(content_q) == 1:
                    content_q[0].append(para)
                else:
                    content_q[1][-1].append(para)
        
        # Create question content
        noidung = self.convert_b4_add_tinhoc(content_q[0], doc)
        
        # Add audio if exists
        if audio and len(audio[0]) > 8:
            link = audio[0].split('Audio:')[1].strip()
            noidung += f'''
        <audio controls="">
            <source src="{link}" type="audio/mpeg">
            Your browser does not support the audio element.
        </audio>
        '''
        
        content_question = self.create_safe_text_node('contentquestion', noidung)
        each_question_xml.append(content_question)
        
        # Process choices A B C D
        # index_answer = self.list_answers_tn_tinhoc(content_q[1], cau_sau_xu_ly[1][0], each_question_xml, doc)

        if not cau_sau_xu_ly[1]:
            # Không có lời giải → xử lý mặc định hoặc báo lỗi
            # Ví dụ: coi như không có đáp án đúng
            index_answer = []
            # Tạo listanswers rỗng nếu cần
            listanswers = ET.SubElement(each_question_xml, 'listanswers')
        else:
            index_answer = self.list_answers_tn_tinhoc(content_q[1], cau_sau_xu_ly[1][0], each_question_xml, doc)
        
        # Process explanation
        array_hdg = []
        if len(cau_sau_xu_ly[1]) > 1:
            array_hdg = cau_sau_xu_ly[1][1]
        
        self.hdg_tn_tinhoc(array_hdg, index_answer, each_question_xml, doc)

    def list_answers_tn_tinhoc(self, content: List, answer_para: Any, 
                                each_question_xml: ET.Element, doc: Document = None) -> List[str]:
        """Process answer list - Tin học"""
        multiple_choices = []
        
        # Process multiple answers or simple answers
        if len(content) > 2 or (len(content) == 2 and len(content[0]) == 1 and len(content[1]) == 1):
            for array_para in content:
                choice = self.convert_b4_add_tinhoc(array_para, doc)
                multiple_choices.append(choice)
        else:
            # Process answers with images or complex content
            tmp_raw = []
            
            for array_para in content:
                for para in array_para:
                    for i in range(get_num_children(para)):
                        child = get_child(para, i)
                        
                        if get_element_type(child) == 'TEXT':
                            text = get_text(child).strip()
                            if re.match(r'^[A-Da-d]\.', text):
                                tmp_raw.append([child])
                            elif tmp_raw:
                                tmp_raw[-1].append(child)
                        elif tmp_raw:
                            tmp_raw[-1].append(child)
            
            for choice in tmp_raw:
                after_convert = self.convert_dang_tn_tinhoc(choice, doc)
                multiple_choices.append(after_convert)
        
        # Get correct answer
        answer_tn = get_text(get_child(answer_para[0], 0)).strip()
        number_of_answer = [s for s in re.split(r'(\S)', answer_tn) if s]
        
        listanswers = ET.SubElement(each_question_xml, 'listanswers')
        
        # Create XML for answers
        for i, choice in enumerate(multiple_choices):
            answer = ET.SubElement(listanswers, 'answer')
            
            index = ET.SubElement(answer, 'index')
            index.text = str(i)
            
            cont = self.create_safe_text_node('content', choice)
            answer.append(cont)
            
            dung_sai = 'TRUE' if str(i + 1) in number_of_answer else 'FALSE'
            isans = ET.SubElement(answer, 'isanswer')
            isans.text = dung_sai
        
        return number_of_answer

    def convert_dang_tn_tinhoc(self, children: List, doc: Document = None) -> str:
        """Convert answer for TN format - Tin học"""
        result = ''
        
        for child in children:
            if hasattr(child, 'getType') and get_element_type(child) == 'TEXT':
                text = self.process_style_tinhoc(child)
                result += text
            elif get_element_type(child) == 'INLINE_IMAGE':
                blob = get_blob(child)
                width = get_width(child)
                height = get_height(child)
                img_base64 = base64.b64encode(get_bytes(blob)).decode('utf-8')
                
                result += f'<img style="width:{width}px;height:{height}px;" src="data:image/png;base64,{img_base64}" />'
        
        return result

    def process_style_tinhoc(self, text_element: Any) -> str:
        """Process text with formatting (bold, italic, underline)"""
        full_text = get_text(text_element) or ''
        indices = get_text_attribute_indices(text_element)
        html_content = ''
        
        prev_format = {'bold': False, 'italic': False, 'underline': False}
        
        for j, start_pos in enumerate(indices):
            end_pos = indices[j + 1] if j + 1 < len(indices) else len(full_text)
            segment_text = full_text[start_pos:end_pos]
            
            attrs = get_attributes(text_element, start_pos)
            current_format = {
                'bold': attrs.get('BOLD') is True,
                'italic': attrs.get('ITALIC') is True,
                'underline': attrs.get('UNDERLINE') is True
            }
            
            # Close previous tags if format changed
            if prev_format != current_format:
                if prev_format['underline']:
                    html_content += '</u>'
                if prev_format['italic']:
                    html_content += '</i>'
                if prev_format['bold']:
                    html_content += '</strong>'
                
                # Open new tags
                if current_format['bold']:
                    html_content += '<strong>'
                if current_format['italic']:
                    html_content += '<i>'
                if current_format['underline']:
                    html_content += '<u>'
                
                prev_format = current_format.copy()
            
            html_content += segment_text
        
        # Close remaining tags
        if prev_format['underline']:
            html_content += '</u>'
        if prev_format['italic']:
            html_content += '</i>'
        if prev_format['bold']:
            html_content += '</strong>'
        
        return html_content
    def hdg_tn_tinhoc(self, array_hdg: List, index_answer: List[str], 
                    each_question_xml: ET.Element, doc: Document = None):
        """Process explanation for TN - Tin học.
        Nếu có hướng dẫn (hdg) thực sự thì thêm vào, 
        còn nếu trống thì KHÔNG thêm 'Đáp án đúng là ...'."""
        import re

        answer = ['A', 'B', 'C', 'D']

        # Xử lý đáp án đúng (vẫn giữ nguyên)
        dap_an = ' '.join(answer[int(idx) - 1] for idx in index_answer)

        # Convert phần hướng dẫn
        huong_dan_giai = self.convert_b4_add_tinhoc(array_hdg, doc).strip()

        # Loại bỏ thẻ HTML để kiểm tra text thực
        text_check = re.sub(r'<.*?>', '', huong_dan_giai).strip()

        # ✅ Nếu có nội dung thật (dài hơn 0 ký tự) thì dùng luôn
        if len(text_check) > 0:
            loi_giai = huong_dan_giai
        else:
            # Nếu không có hdg, KHÔNG thêm "Đáp án đúng là ..."
            loi_giai = ""

        # ✅ Chỉ tạo node nếu thực sự có nội dung
        if loi_giai:
            explainq = self.create_safe_text_node('explainquestion', loi_giai)
            each_question_xml.append(explainq)

    # ============================================
    # DS (True/False) Processing Functions
    # ============================================

    def dang_ds_tinhoc(self, cau_sau_xu_ly: List, each_question_xml: ET.Element, audio: List, doc: Document = None):
        """Process true/false question - Tin học"""
        type_ans = ET.SubElement(each_question_xml, 'typeAnswer')
        type_ans.text = '1'
        
        type_view = ET.SubElement(each_question_xml, 'typeViewContent')
        type_view.text = '0'
        
        temp = ET.SubElement(each_question_xml, 'template')
        temp.text = '0'
        
        # Process hint question
        if len(cau_sau_xu_ly[1]) > 2:
            hint = self.convert_b4_add_tinhoc(cau_sau_xu_ly[1][2], doc)
            hint_question = self.create_safe_text_node('hintQuestion', hint)
            each_question_xml.append(hint_question)
        
        # Parse question content
        content_q = []
        for index, para in enumerate(cau_sau_xu_ly[0]):
            child = get_child(para, 0)
            
            if index == 0:
                content_q.append([para])
                continue
            
            if get_element_type(child) == 'TEXT':
                text = get_text(child).strip()
                
                # Detect true/false items: a) b) c) d) or 1) 2) 3) 4)
                if (re.match(r'^[a-d]\)', text) or 
                    re.match(r'^\d+\)', text) or 
                    re.match(r'^-\s*[a-d]\)', text)):
                    if len(content_q) == 1:
                        content_q.append([[para]])
                    else:
                        content_q[1].append([para])
                    continue
            
            # Add to content
            if len(content_q) == 1:
                content_q[0].append(para)
            else:
                content_q[1][-1].append(para)
        
        # Validate structure
        if len(content_q) != 2:
            cau = get_text(get_child(content_q[0][0], 0))
            raise ValueError(f'Không đúng dạng ĐS kiểm tra lại: {cau}')
        
        # Validate answer count
        answers = get_text(get_child(cau_sau_xu_ly[1][0][0], 0)).strip()
        if len(answers) != len(content_q[1]):
            cau = get_text(get_child(content_q[0][0], 0))
            raise ValueError(f'Không đúng số lượng đáp án dạng ĐS: {cau}')
        
        self.question_ds_tinhoc(content_q, each_question_xml, audio, doc)
        self.dap_an_ds_tinhoc(answers, each_question_xml, content_q[1], doc)
        
        array_hdg = []
        if len(cau_sau_xu_ly[1]) > 1:
            array_hdg = cau_sau_xu_ly[1][1:]
        
        self.hdg_ds_tinhoc(content_q[1], array_hdg, answers, each_question_xml, doc)

    def question_ds_tinhoc(self, content_q: List, each_question_xml: ET.Element, audio: List, doc: Document = None):
        """Process question content for DS format - Tin học"""
        noidung = self.convert_b4_add_tinhoc(content_q[0], doc)
        
        if audio and len(audio[0]) > 8:
            link = audio[0].split('Audio:')[1].strip()
            noidung += f'''
        <audio controls="">
            <source src="{link}" type="audio/mpeg">
            Your browser does not support the audio element.
        </audio>
        '''
        
        cont_xml = self.create_safe_text_node('contentquestion', noidung)
        each_question_xml.append(cont_xml)

    def dap_an_ds_tinhoc(self, answers: str, each_question_xml: ET.Element, content: List, doc: Document = None):
        """Process answers for DS format - Tin học"""
        listanswers = ET.SubElement(each_question_xml, 'listanswers')
        
        for so in range(len(answers)):
            value_text = self.convert_b4_add_tinhoc(content[so], doc).strip()
            value_text = re.sub(r'^.*?\)', '', value_text).strip()
            
            answer = ET.SubElement(listanswers, 'answer')
            
            index = ET.SubElement(answer, 'index')
            index.text = str(so)
            
            cont = self.create_safe_text_node('content', value_text)
            answer.append(cont)
            
            isans = ET.SubElement(answer, 'isanswer')
            isans.text = 'TRUE' if answers[so] == '1' else 'FALSE'

    def hdg_ds_tinhoc(self, content_q: List, array_hdg: List, answers: str, 
                      each_question_xml: ET.Element, doc: Document = None):
        """Process explanation for DS - Tin học"""
        loi_giai = ''
        co_hdg = False
        
        if array_hdg and array_hdg[0]:
            hdg = self.convert_b4_add_tinhoc(array_hdg[0], doc)
            test = re.sub(r'<.*?>', '', hdg)
            test = re.sub(r'</.*?>', '', test)
            if len(test.strip()) > 4:
                co_hdg = True
                loi_giai = hdg
        
        explainq = self.create_safe_text_node('explainquestion', loi_giai)
        each_question_xml.append(explainq)

    # ============================================
    # Content Conversion Functions
    # ============================================

   
    def convert_table_tinhoc(self, table: Any) -> str:
        """Convert table to HTML - Tin học"""
        html = "<table class='table-material-question'>"
        num_rows = get_num_rows(table)
        
        for r in range(num_rows):
            html += '<tr>'
            row = get_row(table, r)
            num_cells = get_num_cells(row)
            
            for c in range(num_cells):
                cell = get_cell(row, c)
                num_paras = get_num_children(cell)
                cell_html = ''
                
                for p in range(num_paras):
                    child = get_child(cell, p)
                    if get_element_type(child) == 'PARAGRAPH':
                        text = get_text(child).strip()
                        if text:
                            cell_html += f'<p>{text}</p>'
                
                html += f'<td>{cell_html}</td>'
            
            html += '</tr>'
        
        html += '</table><br>'
        return html

    def convert_b4_add_tinhoc(self, paragraphs: List, doc: Document = None) -> str:
        """Convert paragraphs to HTML content - Tin học"""
        string_content = ''

        for index, paragraph in enumerate(paragraphs):
            new_children = []

            # Nếu là bảng
            if paragraph._element.tag.endswith('tbl'):
                html_table = self.convert_table_tinhoc(paragraph)
                new_children.append(html_table)
            else:
                # Xử lý text
                self.convert_normal_paras_tinhoc(paragraph, index, new_children, doc)

                # ✅ Xử lý ảnh (giống hoàn toàn convert_b4_add)
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
                                img_tag = self._make_img_tag_from_rid(rId, doc)
                                if img_tag:
                                    new_children.append(img_tag)
                        except Exception:
                            pass

            new_content = ''.join(new_children)
            if len(paragraphs) > 1:
                string_content += f'{new_content}<br>'
            else:
                string_content += new_content

        # ✅ Xử lý MathJax/LaTeX
        import re
        math_latex = re.compile(r'\$[^$]*\$')
        string_content = math_latex.sub(
            lambda m: f' <span class="math-tex">{m.group()}</span>',
            string_content
        )

        return string_content


    def _make_img_tag_from_rid(self, rId: str, doc: Document) -> str:
        """Dùng rId để lấy image part từ doc.part.related_parts, trả về thẻ <img src="data:...">"""
        try:
            part = doc.part.related_parts.get(rId)
            if not part:
                for rel in doc.part.rels.values():
                    target = getattr(rel, 'target_part', None)
                    if target and getattr(target, 'content_type', '').startswith('image') and rel.rId == rId:
                        part = target
                        break

            if not part:
                return ''

            img_bytes = part.blob
            content_type = getattr(part, 'content_type', 'image/png')
            import base64
            b64 = base64.b64encode(img_bytes).decode('ascii')
            return f'<center><img src="data:{content_type};base64,{b64}" /></center>'
        except Exception:
            return ''


    def convert_normal_paras_tinhoc(self, paragraph: Any, index: int, new_children: List, doc: Document = None):
        """Convert normal paragraphs to HTML - Tin học"""
        prev_format = {'bold': False, 'italic': False, 'underline': False}
        html_content = ''
        num_children = get_num_children(paragraph)

        for i in range(num_children):
            child = get_child(paragraph, i)
            if get_element_type(child) == 'TEXT':
                text_element = child
                full_text = get_text(text_element) or ''
                indices = get_text_attribute_indices(text_element)

                for j, start_pos in enumerate(indices):
                    end_pos = indices[j + 1] if j + 1 < len(indices) else len(full_text)
                    segment_text = full_text[start_pos:end_pos]
                    attrs = get_attributes(text_element, start_pos)
                    current_format = {
                        'bold': attrs.get('BOLD') is True,
                        'italic': attrs.get('ITALIC') is True,
                        'underline': attrs.get('UNDERLINE') is True
                    }

                    format_changed = (
                        prev_format['bold'] != current_format['bold']
                        or prev_format['italic'] != current_format['italic']
                        or prev_format['underline'] != current_format['underline']
                    )

                    if format_changed:
                        if prev_format['underline']:
                            html_content += '</u>'
                        if prev_format['italic']:
                            html_content += '</i>'
                        if prev_format['bold']:
                            html_content += '</strong>'

                        if current_format['bold']:
                            html_content += '<strong>'
                        if current_format['italic']:
                            html_content += '<i>'
                        if current_format['underline']:
                            html_content += '<u>'

                        prev_format = current_format.copy()

                    html_content += segment_text

        if prev_format['underline']:
            html_content += '</u>'
        if prev_format['italic']:
            html_content += '</i>'
        if prev_format['bold']:
            html_content += '</strong>'

        import re
        if index == 0:
            html_content = re.sub(r'^(<[^>]*>)*C[ââ]u\s*\d+[\.:]\s*', '', html_content, flags=re.IGNORECASE)
        html_content = re.sub(r'^HL:\s*', '', html_content, flags=re.IGNORECASE)
        html_content = re.sub(
            r'^(<[^>]*>)*([A-D])\.\s*',
            lambda m: m.group(1) or '',
            html_content,
            flags=re.IGNORECASE
        )

        new_children.append(html_content.strip())