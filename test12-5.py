import re
from lxml import etree as ET
import copy

namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
ET.register_namespace('w', namespaces['w'])

# Conversion rules including brackets and other symbols
conversion_rules = [
    # 半角カッコに変換するルール
    {
        'name': 'カッコの校閲（半角: 数字・ローマ数字・アルファベット）',
        'pattern': r'（[0-9ivxlcIVXLCa-zA-Z\-]+）',
        'replace': lambda match: match.group().replace('（', '(').replace('）', ')'),
        'color': 'cyan',
        'check_japanese': False
    },
    # 全角カッコに変換するルール
    {
        'name': 'カッコの校閲（全角: 日本語を含む）',
        'pattern': r'\([^\divxlcIVXLCa-zA-Z\-]*[ぁ-んァ-ヶ一-龠]+[^\divxlcIVXLCa-zA-Z\-]*\)',
        'replace': lambda match: match.group().replace('(', '（').replace(')', '）'),
        'color': 'magenta',
        'check_japanese': True
    },
    # 全角英字を半角英字に変換
    {
        'name': '全角英字を半角英字に変換',
        'pattern': r'[Ａ-Ｚａ-ｚ](?!．)',
        'replace': lambda match: chr(ord(match.group()) - 0xFEE0),
        'color': 'blue',
        'check_japanese': False
    },
    # 全角数字を半角数字に変換
    {
        'name': '全角数字を半角数字に変換',
        'pattern': r'[０-９]',
        'replace': lambda match: chr(ord(match.group()) - 0xFEE0),
        'color': 'green',
        'check_japanese': False
    },
    # 全角記号を半角記号に変換
    {
        'name': '全角記号を半角記号に変換（特定の条件を除く）',
        'pattern': r'([！＂＃＄＆＇＊＜＞＠［＼］＾＿｀｛｜｝／])',
        'replace': lambda match: {
            '！': '!', '＂': '"', '＃': '#', '＄': '$', '＆': '&', '＇': "'", '＊': '*', '＜': '<', '＞': '>', '＠': '@', 
            '［': '[', '＼': '\\', '］': ']', '＾': '^', '＿': '_', '｀': '`', '｛': '{', '｜': '|', '｝': '}', '／': '/'
        }.get(match.group(), match.group()),
        'color': 'red',
        'check_japanese': False
    },
    # 半角記号を全角記号に変換
    {
        'name': '半角記号を全角記号に変換（数式文字および「〜、：、％」）',
        'pattern': r'([~:%\+\*÷\=])',
        'replace': lambda match: {
            '~': '〜', ':': '：', '%': '％',
            '+': '＋', '*': '×', '÷': '÷', '=': '＝'
        }.get(match.group(), match.group()),
        'color': 'yellow',
        'check_japanese': False
    },
]

def split_and_process_runs(paragraph, log_file, new_runs_log_file, rules):
    runs = paragraph.findall('.//w:r', namespaces)
    if not runs:
        return

    combined_text = ""
    for run in runs:
        t_elements = run.findall('.//w:t', namespaces)
        if not t_elements:
            continue

        for t_element in t_elements:
            text = t_element.text
            if not text:
                continue

            combined_text += text  # テキストを結合

    # 結合されたテキストに対してカッコの判定と変換を行う
    for rule in rules:
        combined_text = apply_conversion_rule(combined_text, rule)

    # 変換後のテキストを元の<t>要素に分割して適用
    split_and_apply_converted_text(runs, combined_text)

    # ログの記録
    log_file.write(f"Processed Text: {combined_text}\n")
    log_file.write("-" * 40 + "\n")

def apply_conversion_rule(text, rule):
    if rule['check_japanese']:
        if re.search('[ぁ-んァ-ヶ一-龠]', text):
            return re.sub(rule['pattern'], rule['replace'], text)
    else:
        return re.sub(rule['pattern'], rule['replace'], text)
    return text

def split_and_apply_converted_text(runs, combined_text):
    current_position = 0
    for run in runs:
        t_elements = run.findall('.//w:t', namespaces)
        if not t_elements:
            continue

        for t_element in t_elements:
            text_length = len(t_element.text)
            new_text = combined_text[current_position:current_position + text_length]
            t_element.text = new_text
            current_position += text_length


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