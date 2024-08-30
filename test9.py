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
        'pattern': r'(?<![ぁ-んァ-ヶ一-龠a-zA-Z])([！＂＃＄＆＇＊＋－＜＝＞＠［＼］＾＿｀｛｜｝])',
        'replace': lambda match: {
            '！': '!', '＂': '"', '＃': '#', '＄': '$', '＆': '&', '＇': "'", '＊': '*', 
            '＋': '+', '－': '-', '＜': '<', '＝': '=', '＞': '>', '＠': '@', 
            '［': '[', '＼': '\\', '］': ']', '＾': '^', '＿': '_', '｀': '`', 
            '｛': '{', '｜': '|', '｝': '}'
        }.get(match.group(), match.group()),  # 変換対象外の記号をそのまま返す
        'color': 'red'
    },
    {
        'name': '全角カッコを半角カッコに変換（項目番号）',
        'pattern': r'（([1-9][0-9]?|i{1,3}|iv|v|vi{0,3}|[a-zA-Z](-[1-9][0-9]{0,2}){0,2})）(?=\s)',
        'replace': lambda match: match.group().replace('（', '(').replace('）', ')'),
        'color': 'orange'
    },
    {
        'name': '半角カッコを全角カッコに変換（項目番号以外）',
        'pattern': r'\((?!([1-9][0-9]?|i{1,3}|iv|v|vi{0,3}|[a-zA-Z](-[1-9][0-9]{0,2}){0,2})\s*\))',
        'replace': lambda match: match.group().replace('(', '（').replace(')', '）'),
        'color': 'purple'
    },
    {
        'name': '半角閉じカッコを全角閉じカッコに変換（項目番号以外）',
        'pattern': r'(?<!\()\)(?!([1-9][0-9]?|i{1,3}|iv|v|vi{0,3}|[a-zA-Z](-[1-9][0-9]{0,2}){0,2}))',
        'replace': lambda match: match.group().replace(')', '）'),
        'color': 'purple'
    },
    {
        'name': '全角閉じカッコを半角閉じカッコに変換（項目番号）',
        'pattern': r'（([1-9][0-9]?|i{1,3}|iv|v|vi{0,3}|[a-zA-Z](-[1-9][0-9]{0,2}){0,2})）',
        'replace': lambda match: match.group().replace('（', '(').replace('）', ')'),
        'color': 'orange'
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

def process_paragraph(paragraph, log_file, new_runs_log_file, rules):
    runs = paragraph.findall('.//w:r', namespaces)
    if not runs:
        return  # runsが空の場合は処理をスキップ

    parent = runs[0].getparent()

    for run in runs:
        if run.getparent() is None:
            continue  # runが既に削除されている場合はスキップ

        t_elements = run.findall('.//w:t', namespaces)
        if not t_elements:
            continue

        original_text = "".join([t.text for t in t_elements if t.text])

        if original_text.strip() == "":
            continue

        match_positions = apply_conversion_rules(original_text, rules)
        if not match_positions:
            continue

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

        # 元の要素の位置を取得
        if run in parent:
            index = parent.index(run)

            # 元の要素を削除
            parent.remove(run)

            # 新しい要素を元の位置に順番に挿入
            for i, new_run in enumerate(new_runs):
                parent.insert(index + i, new_run)  # 挿入するたびにインデックスがずれるので、順番を維持
                # ログに追加された新しい要素を記録
                new_runs_log_file.write(ET.tostring(new_run, encoding='unicode'))
                new_runs_log_file.write("\n" + "-"*40 + "\n")

def process_xml_with_conversion_rules(xml_file, log_filename, new_runs_log_filename):
    with open(log_filename, 'w', encoding='utf-8') as log_file, open(new_runs_log_filename, 'w', encoding='utf-8') as new_runs_log_file:
        tree = ET.parse(xml_file)
        root = tree.getroot()

        for paragraph in root.findall('.//w:p', namespaces):
            process_paragraph(paragraph, log_file, new_runs_log_file, conversion_rules)

        tree.write(xml_file, encoding='utf-8', xml_declaration=True, pretty_print=True)

# 呼び出し部分
process_xml_with_conversion_rules('xml_new/word/document.xml', 'conversion_rules_log.txt', 'new_runs_log.txt')