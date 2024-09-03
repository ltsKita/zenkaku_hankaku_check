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
        'pattern': r'([！＂＃＄＆＇＊＜＞＠［＼］＾＿｀｛｜｝])',
        'replace': lambda match: {
            '！': '!', '＂': '"', '＃': '#', '＄': '$', '＆': '&', '＇': "'", '＊': '*', '＜': '<', '＞': '>', '＠': '@', 
            '［': '[', '＼': '\\', '］': ']', '＾': '^', '＿': '_', '｀': '`', '｛': '{', '｜': '|', '｝': '}'
        }.get(match.group(), match.group()),
        'color': 'red'
    },
    {
        'name': '半角記号を全角記号に変換（数式文字および「〜、：、％」）',
        'pattern': r'([~:%\+\*÷\=])',
        'replace': lambda match: {
            '~': '〜', ':': '：', '%': '％',
            '+': '＋', '*': '×', '÷': '÷', '=': '＝'
        }.get(match.group(), match.group()),
        'color': 'yellow'
    },
    {
        'name': '括弧の校閲ルールを適用',
        'pattern': r'（(\d{1,2})）|（([a-z])）|（([a-z]-\d{1,2})）|（([a-z]-\d{1,2}-\d{1,2})）|\(([^0-9a-zA-Z]+)\)',
        'replace': lambda match: (
            f'({match.group(1)})' if match.group(1) else
            f'({match.group(2)})' if match.group(2) else
            f'({match.group(3)})' if match.group(3) else
            f'({match.group(4)})' if match.group(4) else
            f'（{match.group(5)}）'
        ),
        'color': 'purple'
    }
]

def get_all_text_elements(element):
    """再帰的にすべての<w:t>要素を取得し、親要素も一緒に返す"""
    text_elements = []
    if element.tag.endswith('t'):
        text_elements.append(element)
    for child in element:
        text_elements.extend(get_all_text_elements(child))
    return text_elements

def split_and_process_runs(paragraph, log_file, new_runs_log_file, rules):
    runs = paragraph.findall('.//w:r', namespaces)
    if not runs:
        return

    for i, run in enumerate(runs):
        # mc:AlternateContent などのタグを変更せずに処理を行う
        if any(child.tag.endswith('AlternateContent') for child in run):
            continue  # 変更せずスキップ

        t_elements = get_all_text_elements(run)
        if not t_elements:
            continue  # <w:t>要素がない場合はスキップ

        original_text = "".join([t.text for t in t_elements if t.text])
        if original_text.strip() == "":
            continue

        combined_text = original_text
        if original_text in ['（', '）']:
            combined_text = get_surrounding_text(paragraph, i, t_elements[0])

        match_positions = apply_conversion_rules(combined_text, rules)
        if not match_positions:
            continue

        # ハイライトを適用するための新しいラン要素を保持
        new_elements = []
        current_position = 0
        for start, end, replaced_text, color, rule_name, original_text_part in match_positions:
            if current_position < start:
                unhighlighted_text = combined_text[current_position:start]
                if unhighlighted_text:
                    new_run = create_run_with_text_and_style(run, unhighlighted_text)
                    new_elements.append(new_run)

            highlighted_run = create_run_with_text_and_style(run, replaced_text, color)
            new_elements.append(highlighted_run)

            log_file.write(f"Matched Rule: {rule_name}\n")
            log_file.write(f"Original Text: {original_text}\n")
            log_file.write(f"Combined Text: {combined_text}\n")
            log_file.write(f"Original Text (Matched Part): {original_text_part}\n")
            log_file.write(f"Replaced Text: {replaced_text}\n")
            log_file.write("-" * 40 + "\n")

            current_position = end

        if current_position < len(combined_text):
            remaining_text = combined_text[current_position:]
            new_run = create_run_with_text_and_style(run, remaining_text)
            new_elements.append(new_run)

        parent = run.getparent()
        if parent is not None:
            for j, new_element in enumerate(new_elements):
                parent.insert(parent.index(run) + j, new_element)
            parent.remove(run)

def get_surrounding_text(paragraph, run_index, current_t):
    """指定された<w:t>要素の前後のテキストを取得"""
    prev_text = ''
    next_text = ''

    runs = paragraph.findall('.//w:r', namespaces)

    # 前のテキストを取得
    if run_index > 0:
        prev_run = runs[run_index - 1]
        prev_text = "".join([t.text for t in get_all_text_elements(prev_run) if t.text])

    # 次のテキストを取得
    if run_index < len(runs) - 1:
        next_run = runs[run_index + 1]
        next_text = "".join([t.text for t in get_all_text_elements(next_run) if t.text])

    return prev_text + current_t.text + next_text

def create_run_with_text_and_style(original_run, text, color=None):
    new_run = copy.deepcopy(original_run)
    t_elements = get_all_text_elements(new_run)

    if t_elements:
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

def process_paragraph(paragraph, log_file, new_runs_log_file, rules):
    split_and_process_runs(paragraph, log_file, new_runs_log_file, rules)

def process_xml_with_conversion_rules(xml_file, log_filename, new_runs_log_filename):
    with open(log_filename, 'w', encoding='utf-8') as log_file, open(new_runs_log_filename, 'w', encoding='utf-8') as new_runs_log_file:
        tree = ET.parse(xml_file)
        root = tree.getroot()

        for paragraph in root.findall('.//w:p', namespaces):
            process_paragraph(paragraph, log_file, new_runs_log_file, conversion_rules)

        tree.write(xml_file, encoding='utf-8', xml_declaration=True, pretty_print=True)

# 呼び出し部分
process_xml_with_conversion_rules('xml_new/word/document.xml', 'conversion_rules_log.txt', 'new_runs_log.txt')