import re
from lxml import etree as ET
import copy

namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
ET.register_namespace('w', namespaces['w'])

# Conversion rules, including parentheses rules
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

def split_and_process_runs(paragraph, log_file, new_runs_log_file, rules):
    runs = paragraph.findall('.//w:r', namespaces)
    if not runs:
        return

    # 結合して段落全体のテキストを取得
    original_text = "".join([t.text for run in runs for t in run.findall('.//w:t', namespaces) if t.text])
    if original_text.strip() == "":
        return

    match_positions = apply_conversion_rules(original_text, rules)
    if not match_positions:
        return

    # 元の<w:t>要素にマッピングして修正を反映
    modified_text = list(original_text)
    for start, end, replaced_text, color, rule_name, original_text_part in match_positions:
        modified_text[start:end] = list(replaced_text)  # 置換を実行

    # 修正されたテキストを元の<w:t>要素に反映
    current_pos = 0
    for run in runs:
        for t in run.findall('.//w:t', namespaces):
            if t.text:
                text_length = len(t.text)
                t.text = ''.join(modified_text[current_pos:current_pos + text_length])
                current_pos += text_length

    # 変更があった部分にハイライトを追加
    for start, end, replaced_text, color, rule_name, original_text_part in match_positions:
        apply_highlight_to_runs(runs, start, end, color)

    # ログに追加情報を書き込む
    log_file.write(f"Matched Rule: {rule_name}\n")
    log_file.write(f"Original Full Text: {original_text}\n")
    log_file.write(f"Replaced Text: {''.join(modified_text)}\n")
    log_file.write("-" * 40 + "\n")

def apply_highlight_to_runs(runs, start, end, color):
    """修正箇所にハイライトを適用"""
    current_pos = 0
    for run in runs:
        for t in run.findall('.//w:t', namespaces):
            if t.text:
                text_length = len(t.text)
                if current_pos + text_length >= start and current_pos <= end:
                    # ハイライトが必要な場合にハイライトを追加
                    add_highlight_to_run(run, color)
                current_pos += text_length

def add_highlight_to_run(run, color):
    """指定されたランにハイライトを追加"""
    rpr = run.find('.//w:rPr', namespaces)
    if rpr is None:
        rpr = ET.SubElement(run, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
    highlight_elem = ET.SubElement(rpr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}highlight')
    highlight_elem.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', color)

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