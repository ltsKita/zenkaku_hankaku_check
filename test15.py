import re
from lxml import etree as ET
import copy

# 名前空間を定義
namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

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

def split_and_process_runs(paragraph, log_file, new_runs_log_file, rules):
    runs = paragraph.findall('.//w:r', namespaces)
    if not runs:
        return

    for run in runs:
        t_elements = run.findall('.//w:t', namespaces)
        if not t_elements:
            continue  # <w:t>要素がない場合はスキップ

        original_text = "".join([t.text for t in t_elements if t.text])
        if original_text.strip() == "":
            continue

        combined_text = original_text
        if original_text in ['（', '）', '(', ')']:
            combined_text = get_surrounding_text(paragraph, run, t_elements[0])

        match_positions = apply_conversion_rules(combined_text, rules)
        if not match_positions:
            continue

        new_runs = []
        current_position = 0

        for start, end, replaced_text, color, rule_name, original_text_part in match_positions:
            if start > current_position:
                unmodified_text = combined_text[current_position:start]
                new_runs.append(create_new_run(unmodified_text, run))

            modified_run = create_new_run(replaced_text, run, color)
            new_runs.append(modified_run)

            current_position = end

        # 残りのテキストを適切に処理
        if current_position < len(combined_text):
            remaining_text = combined_text[current_position:]
            new_runs.append(create_new_run(remaining_text, run))

        # 元の要素を削除せず、未処理のテキストを残す
        update_original_run(run, combined_text, current_position)
        insert_new_runs(paragraph, run, new_runs)

        # ログの出力
        log_file.write(f"Matched Rule: {rule_name}\n")
        log_file.write(f"Original Text: {original_text}\n")
        log_file.write(f"Combined Text: {combined_text}\n")
        log_file.write(f"Original Text (Matched Part): {original_text_part}\n")
        log_file.write(f"Replaced Text: {replaced_text}\n")
        log_file.write("-" * 40 + "\n")

def create_new_run(text, original_run, color=None):
    """新しい<w:r>要素を作成"""
    new_run = copy.deepcopy(original_run)
    t_element = new_run.find('.//w:t', namespaces)
    if t_element is not None:
        t_element.text = text
    else:
        t_element = ET.SubElement(new_run, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
        t_element.text = text

    if color:
        apply_highlight_to_run(new_run, color)

    return new_run

def update_original_run(original_run, combined_text, current_position):
    """元の<w:r>要素を削除せず、未処理のテキストを残す"""
    t_element = original_run.find('.//w:t', namespaces)
    if t_element is not None:
        t_element.text = combined_text[current_position:]

def insert_new_runs(paragraph, original_run, new_runs):
    """元の<w:r>要素を削除せずに新しい要素を挿入する"""
    parent = original_run.getparent()
    original_index = parent.index(original_run)
    
    # 新しい要素を元の要素の直前に追加
    for i, new_run in enumerate(new_runs):
        parent.insert(original_index + i, new_run)

def apply_highlight_to_run(run, color):
    """指定されたランにハイライトを追加"""
    rpr = run.find('.//w:rPr', namespaces)
    if rpr is None:
        rpr = ET.SubElement(run, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
    highlight_elem = ET.SubElement(rpr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}highlight')
    highlight_elem.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', color)

def get_surrounding_text(paragraph, current_run, current_t):
    """指定された<w:t>要素の前後のテキストを取得"""
    prev_text = ''
    next_text = ''

    runs = paragraph.findall('.//w:r', namespaces)

    # 前のテキストを取得
    run_index = runs.index(current_run)
    if run_index > 0:
        prev_run = runs[run_index - 1]
        prev_text = "".join([t.text for t in prev_run.findall('.//w:t', namespaces) if t.text])

    # 次のテキストを取得
    if run_index < len(runs) - 1:
        next_run = runs[run_index + 1]
        next_text = "".join([t.text for t in next_run.findall('.//w:t', namespaces) if t.text])

    return prev_text + current_t.text + next_text

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

        tree.write(xml_file, encoding='utf-8', xml_declaration=True)

# 呼び出し部分
process_xml_with_conversion_rules('xml_new/word/document.xml', 'conversion_rules_log.txt', 'new_runs_log.txt')