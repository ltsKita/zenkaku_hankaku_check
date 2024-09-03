# いったんこれ
import re
from lxml import etree as ET
import copy

namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
ET.register_namespace('w', namespaces['w'])

# Conversion rules including brackets
conversion_rules = [
    {
        'name': 'カッコの校閲（半角）',
        'pattern': r'([\(（][0-9iIvVxX]+[\-0-9iIvVxX]*[a-zA-Z]*[\-0-9iIvVxX]*[\)）])',
        'replace': lambda match: match.group().replace('（', '(').replace('）', ')'),
        'color': 'cyan'
    },
    {
        'name': 'カッコの校閲（全角）',
        'pattern': r'([\(（])(?=\d{3,}|[ぁ-んァ-ヶ一-龠]+[\)）])|(?<=[ぁ-んァ-ヶ一-龠]\))',
        'replace': lambda match: match.group().replace('(', '（').replace(')', '）'),
        'color': 'magenta'
    },
    # Add any other conversion rules if necessary
    # 全角英字を半角英字に変換
    {
        'name': '全角英字を半角英字に変換',
        'pattern': r'[Ａ-Ｚａ-ｚ](?!．)',
        'replace': lambda match: chr(ord(match.group()) - 0xFEE0),
        'color': 'blue'
    },
    # 全角数字を半角数字に変換
    {
        'name': '全角数字を半角数字に変換',
        'pattern': r'[０-９]',
        'replace': lambda match: chr(ord(match.group()) - 0xFEE0),
        'color': 'green'
    },
    # 全角記号を半角記号に変換
    {
        'name': '全角記号を半角記号に変換（特定の条件を除く）',
        'pattern': r'([！＂＃＄＆＇＊＜＞＠［＼］＾＿｀｛｜｝／])',
        'replace': lambda match: {
            '！': '!', '＂': '"', '＃': '#', '＄': '$', '＆': '&', '＇': "'", '＊': '*', '＜': '<', '＞': '>', '＠': '@', 
            '［': '[', '＼': '\\', '］': ']', '＾': '^', '＿': '_', '｀': '`', '｛': '{', '｜': '|', '｝': '}', '／': '/'
        }.get(match.group(), match.group()),
        'color': 'red'
    },
    # 半角記号を全角記号に変換
    {
        'name': '半角記号を全角記号に変換（数式文字および「〜、：、％」）',
        'pattern': r'([~:%\+\*÷\=])',
        'replace': lambda match: {
            '~': '〜', ':': '：', '%': '％',
            '+': '＋', '*': '×', '÷': '÷', '=': '＝'
        }.get(match.group(), match.group()),
        'color': 'yellow'
    },
]

def split_and_process_runs(paragraph, log_file, new_runs_log_file, rules):
    runs = paragraph.findall('.//w:r', namespaces)
    if not runs:
        return

    for run in runs:
        # 図や画像を含む場合はスキップ
        if run.find('.//w:drawing', namespaces) is not None or run.find('.//w:pict', namespaces) is not None:
            continue

        t_elements = run.findall('.//w:t', namespaces)
        if not t_elements:
            continue  # <w:t>要素がない場合はスキップ

        for t_element in t_elements:
            original_text = t_element.text
            if not original_text or original_text.strip() == "":
                continue

            match_positions = apply_conversion_rules(original_text, rules)
            if not match_positions:
                continue

            new_elements = []
            current_position = 0
            for start, end, replaced_text, color, rule_name, original_text_part in match_positions:
                if current_position < start:
                    unhighlighted_text = original_text[current_position:start]
                    if unhighlighted_text:
                        new_run = create_run_with_text_and_style(run, unhighlighted_text)
                        new_elements.append(new_run)

                highlighted_run = create_run_with_text_and_style(run, replaced_text, color)
                new_elements.append(highlighted_run)

                log_file.write(f"Matched Rule: {rule_name}\n")
                log_file.write(f"Original Text: {original_text_part}\n")
                log_file.write(f"Replaced Text: {replaced_text}\n")
                log_file.write("-" * 40 + "\n")

                current_position = end

            if current_position < len(original_text):
                remaining_text = original_text[current_position:]
                new_run = create_run_with_text_and_style(run, remaining_text)
                new_elements.append(new_run)

            # 元の<w:t>要素を新しい要素で置き換える
            parent = t_element.getparent()
            if parent is not None:
                for i, new_element in enumerate(new_elements):
                    parent.insert(parent.index(t_element) + i, new_element)
                parent.remove(t_element)

def create_run_with_text_and_style(original_run, text, color=None):
    new_run = copy.deepcopy(original_run)
    t_elements = new_run.findall('.//w:t', namespaces)

    if t_elements:
        t_elements[0].text = text
        for t_element in t_elements[1:]:
            t_element.getparent().remove(t_element)  # 余分なテキスト要素を削除
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