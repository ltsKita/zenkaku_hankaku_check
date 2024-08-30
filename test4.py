# ハイライトなし

import re
from lxml import etree as ET
import copy

namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
ET.register_namespace('w', namespaces['w'])

# 変換ルールと対応する変換関数の定義
conversion_rules = [
    {
        'name': '全角英字を半角英字に変換',
        'pattern': r'[Ａ-Ｚ]',
        'replace': lambda match: chr(ord(match.group()) - 0xFEE0),
        'color': 'blue'
    },
    {
        'name': '全角数字を半角数字に変換',
        'pattern': r'[０-９]',
        'replace': lambda match: chr(ord(match.group()) - 0xFEE0),
        'color': 'green'
    }
]

def create_plain_run(original_rpr, text):
    plain_run = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
    if original_rpr is not None:
        plain_rpr = copy.deepcopy(original_rpr)
        plain_run.append(plain_rpr)
    plain_text_element = ET.SubElement(plain_run, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
    plain_text_element.text = text
    return plain_run

def create_highlighted_run(original_rpr, text, color):
    new_run = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
    if original_rpr is not None:
        new_rpr = copy.deepcopy(original_rpr)
        new_run.append(new_rpr)
        highlight_elem = ET.SubElement(new_rpr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}highlight')
        highlight_elem.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', color)
    else:
        new_rpr = ET.SubElement(new_run, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
        highlight_elem = ET.SubElement(new_rpr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}highlight')
        highlight_elem.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', color)
    new_text_element = ET.SubElement(new_run, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
    new_text_element.text = text
    return new_run

def apply_conversion_rules(text, rules):
    modified_text = text
    conversion_info = []
    conversion_occurred = False  # 変換が発生したかどうかを追跡
    
    for rule in rules:
        pattern = rule['pattern']
        replace_function = rule['replace']
        
        new_text, num_subs = re.subn(pattern, replace_function, modified_text)
        
        if num_subs > 0:
            conversion_occurred = True
            conversion_info.append((rule['name'], pattern, rule['color'], new_text))
        modified_text = new_text
    
    if not conversion_occurred:
        conversion_info.append(('No conversion', 'No Rule Matched', '', modified_text))
    
    return modified_text, conversion_info

def process_paragraph(paragraph, log_file, rules):
    full_text = "".join(text_elem.text for text_elem in paragraph.findall('.//w:t', namespaces) if text_elem.text)
    log_file.write(f"Full Paragraph Text: {full_text}\n")
    
    modified_text, conversion_info = apply_conversion_rules(full_text, rules)
    if conversion_info:
        log_file.write(f"Processing Paragraph: {full_text}\n")
        for info in conversion_info:
            if info[0] != 'No conversion':
                log_file.write(f"Matched Rule: {info[0]}\n")
                log_file.write(f"Pattern: {info[1]}\n")
                log_file.write(f"Modified Text: {info[3]}\n")

        # 元の<w:t>要素をすべて削除
        for text_elem in paragraph.findall('.//w:t', namespaces):
            parent_run = text_elem.getparent()
            parent = parent_run.getparent()
            if parent is not None:
                parent.remove(parent_run)

        # 新しいテキストを再分割して挿入
        new_elements = []
        current_position = 0
        
        for match in re.finditer('|'.join(rule['pattern'] for rule in rules), modified_text):
            start, end = match.span()
            
            if current_position < start:
                unhighlighted_text = modified_text[current_position:start]
                if unhighlighted_text:
                    new_elements.append(create_plain_run(None, unhighlighted_text))
            
            matched_text = modified_text[start:end]
            for rule in rules:
                if re.match(rule['pattern'], matched_text):
                    replaced_text = re.sub(rule['pattern'], rule['replace'], matched_text)
                    new_elements.append(create_highlighted_run(None, replaced_text, rule['color']))
                    break
            
            current_position = end

        if current_position < len(modified_text):
            remaining_text = modified_text[current_position:]
            new_elements.append(create_plain_run(None, remaining_text))
        
        # 新しい要素を親に挿入
        for new_element in new_elements:
            paragraph.append(new_element)

def process_xml_with_conversion_rules(xml_file, log_filename):
    with open(log_filename, 'w', encoding='utf-8') as log_file:
        tree = ET.parse(xml_file)
        root = tree.getroot()

        for paragraph in root.findall('.//w:p', namespaces):
            process_paragraph(paragraph, log_file, conversion_rules)

        tree.write(xml_file, encoding='utf-8', xml_declaration=True, pretty_print=True)

# 使用例
process_xml_with_conversion_rules('xml_new/word/document.xml', 'conversion_rules_log.txt')
