import re
from lxml import etree as ET
import copy

namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
ET.register_namespace('w', namespaces['w'])

# Conversion rules
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
    },
    {
        'name': '全角記号を半角記号に変換（特定の条件を除く）',
        'pattern': r'[！＂＃＄＆＇＊＋－＜＝＞＠［＼］＾＿｀｛｜｝](?!\s*[a-zA-Z]*．)|(?<![ぁ-んァ-ヶ一-龠])[，。](?![a-zA-Z])',
        'replace': lambda match: {
            '！': '!', '＂': '"', '＃': '#', '＄': '$', '＆': '&', '＇': "'", '＊': '*', 
            '＋': '+', '－': '-', '＜': '<', '＝': '=', '＞': '>', '＠': '@', 
            '［': '[', '＼': '\\', '］': ']', '＾': '^', '＿': '_', '｀': '`', 
            '｛': '{', '｜': '|', '｝': '}', '～': '～'
        }.get(match.group(), match.group()),  # 変換対象外
        'color': 'red'
    },
    {
        'name': '全角カッコを半角カッコに変換（項目番号）',
        'pattern': r'（(1|[1-9][0-9]?)）|（(i|ii|iii|iv|v|vi|vii|viii|ix|x)）|（[a-zA-Z](-[1-9][0-9]{0,2})*(-[1-9][0-9]{0,2})?）',
        'replace': lambda match: match.group().replace('（', '(').replace('）', ')'),
        'color': 'orange'
    },
    {
        'name': '半角カッコを全角カッコに変換（項目番号以外）',
        'pattern': r'\((?!1|[1-9][0-9]?\)|[a-zA-Z](-[1-9][0-9]{0,2})*(-[1-9][0-9]{0,2})?)',
        'replace': lambda match: match.group().replace('(', '（').replace(')', '）'),
        'color': 'purple'
    }
]

def create_run_with_text_and_style(original_run, text, color=None):
    new_run = copy.deepcopy(original_run)
    t_elements = new_run.findall('.//w:t', namespaces)
    
    if len(t_elements) > 0:
        t_elements[0].text = text
        for t_element in t_elements[1:]:
            t_element.getparent().remove(t_element)
    else:
        new_text_element = ET.SubElement(new_run, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
        new_text_element.text = text

    if color:
        rpr = new_run.find('.//w:rPr', namespaces)
        if rpr is None:
            rpr = ET.SubElement(new_run, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
        highlight_elem = ET.SubElement(rpr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}highlight')
        highlight_elem.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', color)

    return new_run

def apply_conversion_rules(text, rules):
    match_positions = []
    for rule in rules:
        for match in re.finditer(rule['pattern'], text):
            start, end = match.span()
            replaced_text = rule['replace'](match)
            match_positions.append((start, end, replaced_text, rule['color'], rule['name'], match.group()))

    match_positions.sort(key=lambda x: x[0])
    return match_positions

def process_run(run, log_file, rules):
    t_elements = run.findall('.//w:t', namespaces)
    
    if not t_elements:
        return
    
    nested_runs = run.findall('.//w:r', namespaces)
    if nested_runs and nested_runs != [run]:
        for nested_run in nested_runs:
            process_run(nested_run, log_file, rules)
        return

    original_text = "".join([t_elem.text for t_elem in t_elements if t_elem.text])
    match_positions = apply_conversion_rules(original_text, rules)

    if not match_positions:
        return

    log_file.write(f"Original Run Text: {original_text}\n")
    parent = run.getparent()
    
    # 親要素が存在しない場合、処理をスキップ
    if parent is None:
        return
    
    index = parent.index(run)
    current_position = 0
    new_runs = []

    for start, end, replaced_text, color, rule_name, original_text_part in match_positions:
        if current_position < start:
            unhighlighted_text = original_text[current_position:start]
            if unhighlighted_text:
                new_run = create_run_with_text_and_style(run, unhighlighted_text)
                new_runs.append(new_run)

        highlighted_run = create_run_with_text_and_style(run, replaced_text, color)
        new_runs.append(highlighted_run)

        log_file.write(f"Matched Rule: {rule_name}\n")
        log_file.write(f"Original Text: {original_text_part}\n")
        log_file.write(f"Replaced Text: {replaced_text}\n")
        log_file.write("-" * 40 + "\n")

        current_position = end

    if current_position < len(original_text):
        remaining_text = original_text[current_position:]
        new_run = create_run_with_text_and_style(run, remaining_text)
        new_runs.append(new_run)

    parent.remove(run)
    for new_run in new_runs:
        parent.insert(index, new_run)
        index += 1

def process_paragraph(paragraph, log_file, rules):
    runs = paragraph.findall('.//w:r', namespaces)
    for run in runs:
        if run.find('.//w:t', namespaces) is not None:
            process_run(run, log_file, rules)

def process_xml_with_conversion_rules(xml_file, log_filename):
    with open(log_filename, 'w', encoding='utf-8') as log_file:
        tree = ET.parse(xml_file)
        root = tree.getroot()

        for paragraph in root.findall('.//w:p', namespaces):
            process_paragraph(paragraph, log_file, conversion_rules)

        tree.write(xml_file, encoding='utf-8', xml_declaration=True, pretty_print=True)

process_xml_with_conversion_rules('xml_new/word/document.xml', 'conversion_rules_log.txt')