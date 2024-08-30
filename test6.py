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

def create_run_with_text_and_style(original_rpr, text, color=None):
    new_run = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
    
    if original_rpr is not None:
        new_rpr = copy.deepcopy(original_rpr)
        new_run.append(new_rpr)
        
        if color is not None:
            highlight_elem = ET.SubElement(new_rpr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}highlight')
            highlight_elem.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', color)
    
    new_text_element = ET.SubElement(new_run, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
    new_text_element.text = text
    
    return new_run

def apply_conversion_rules(text, rules):
    modified_text = text
    conversion_info = []
    match_positions = []
    
    for rule in rules:
        pattern = rule['pattern']
        replace_function = rule['replace']
        
        # 見つかった箇所をリストに追加
        for match in re.finditer(pattern, modified_text):
            start, end = match.span()
            replaced_text = replace_function(match)
            match_positions.append((start, end, replaced_text, rule['color'], rule['name']))
    
    match_positions.sort(key=lambda x: x[0])  # 開始位置でソート
    return match_positions

def process_paragraph(paragraph, log_file, rules):
    full_text = "".join(text_elem.text for text_elem in paragraph.findall('.//w:t', namespaces) if text_elem.text)
    log_file.write(f"Full Paragraph Text: {full_text}\n")
    
    match_positions = apply_conversion_rules(full_text, rules)
    
    # 元の<w:r>要素をすべて削除
    for run in paragraph.findall('.//w:r', namespaces):
        parent = run.getparent()
        if parent is not None:
            parent.remove(run)

    # 新しい要素を作成して挿入
    current_position = 0
    original_rpr = paragraph.find('.//w:rPr', namespaces)  # 元の<w:rPr>を取得
    
    for start, end, replaced_text, color, rule_name in match_positions:
        if current_position < start:
            unhighlighted_text = full_text[current_position:start]
            paragraph.append(create_run_with_text_and_style(original_rpr, unhighlighted_text))

        paragraph.append(create_run_with_text_and_style(original_rpr, replaced_text, color))
        current_position = end

    if current_position < len(full_text):
        remaining_text = full_text[current_position:]
        paragraph.append(create_run_with_text_and_style(original_rpr, remaining_text))

def process_xml_with_conversion_rules(xml_file, log_filename):
    with open(log_filename, 'w', encoding='utf-8') as log_file:
        tree = ET.parse(xml_file)
        root = tree.getroot()

        for paragraph in root.findall('.//w:p', namespaces):
            process_paragraph(paragraph, log_file, conversion_rules)

        tree.write(xml_file, encoding='utf-8', xml_declaration=True, pretty_print=True)

# 使用例
process_xml_with_conversion_rules('xml_new/word/document.xml', 'conversion_rules_log.txt')