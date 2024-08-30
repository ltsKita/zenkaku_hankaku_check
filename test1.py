import re
from lxml import etree as ET
import copy

namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
ET.register_namespace('w', namespaces['w'])

# 変換ルールと対応する変換関数の定義
conversion_rules = [
    {
        'name': 'Full-width to Half-width Symbols',
        'pattern': {'（': '(', '）': ')', '。': '.', '、': ',', '：': ':', '；': ';', '！': '!', '？': '?'},
        'color': 'yellow'
    },
    {
        'name': 'Full-width to Half-width Alphabets',
        'pattern': {chr(0xFF21 + i): chr(0x41 + i) for i in range(26)},  # A-Z
        'color': 'blue'
    },
    {
        'name': 'Full-width to Half-width Numbers',
        'pattern': {chr(0xFF10 + i): chr(0x30 + i) for i in range(10)},  # 0-9
        'color': 'green'
    }
]

def create_plain_run(original_rpr, text):
    """
    元の<w:rPr>を保持しつつ、プレーンテキスト用の<w:r>要素を作成
    """
    plain_run = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
    if original_rpr is not None:
        plain_rpr = copy.deepcopy(original_rpr)
        plain_run.append(plain_rpr)
    plain_text_element = ET.SubElement(plain_run, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
    plain_text_element.text = text
    return plain_run

def create_highlighted_run(original_rpr, text, color):
    """
    新しい <w:r> 要素を作成し、指定された色でハイライトを適用した <w:t> を含む。
    元の <w:rPr> 要素をそのままコピーして適用し、ハイライトを追加する。
    """
    new_run = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')

    if original_rpr is not None:
        # 元の <w:rPr> 要素を深くコピー
        new_rpr = copy.deepcopy(original_rpr)
        new_run.append(new_rpr)

        # ハイライトの要素を追加
        highlight_elem = ET.SubElement(new_rpr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}highlight')
        highlight_elem.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', color)
    else:
        # 元の <w:rPr> がない場合でもハイライトを追加
        new_rpr = ET.SubElement(new_run, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
        highlight_elem = ET.SubElement(new_rpr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}highlight')
        highlight_elem.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', color)

    # 新しい <w:t> 要素を追加
    new_text_element = ET.SubElement(new_run, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
    new_text_element.text = text

    return new_run

def apply_conversion_rules(text, rules):
    """
    すべての変換ルールを適用してテキストを変換し、変換情報を取得
    """
    modified_text = text
    conversion_info = []
    
    for rule in rules:
        pattern = rule['pattern']
        for full_width, half_width in pattern.items():
            if re.search(re.escape(full_width), text):
                modified_text = re.sub(re.escape(full_width), half_width, modified_text)
                # 変更された箇所の情報を保存
                conversion_info.append((full_width, half_width, rule['color'], rule['name']))
    
    return modified_text, conversion_info

def split_and_highlight_with_rules(text_element, log_file, rules):
    """
    変換ルールに基づいてテキストを変換し、変更箇所にハイライトを適用
    """
    parent_run = text_element.getparent()
    original_rpr = parent_run.find('.//w:rPr', namespaces)
    
    # <w:t>要素を結合して1つのテキストにする
    combined_text = "".join([t.text for t in parent_run.findall('.//w:t', namespaces) if t.text])
    
    # 変換ルールを適用
    modified_text, conversion_info = apply_conversion_rules(combined_text, rules)

    # ログファイルに変更前後の情報を書き出し
    if conversion_info:
        log_file.write("Original Text: {}\n".format(combined_text))
        for info in conversion_info:
            log_file.write("Condition: {}\n".format(info[3]))
            log_file.write("Before: {}\n".format(info[0]))
            log_file.write("After: {}\n".format(info[1]))
        log_file.write("Modified Text: {}\n\n".format(modified_text))

    # ハイライトとテキストの置き換え処理
    new_elements = []
    current_position = 0

    # 変更された箇所を検出し、ハイライトを適用
    for full_width, half_width, color, _ in conversion_info:
        for match in re.finditer(re.escape(half_width), modified_text):
            start, end = match.span()
            
            # ハイライトされない部分を追加
            if current_position < start:
                unhighlighted_text = modified_text[current_position:start]
                if unhighlighted_text:
                    new_elements.append(create_plain_run(original_rpr, unhighlighted_text))
            
            # ハイライトされた部分を追加
            new_elements.append(create_highlighted_run(original_rpr, half_width, color))
            current_position = end

    # 残りの部分を追加
    if current_position < len(modified_text):
        remaining_text = modified_text[current_position:]
        new_elements.append(create_plain_run(original_rpr, remaining_text))

    # 元の要素を削除して新しい要素を追加
    parent = parent_run.getparent()
    for new_element in new_elements:
        parent.insert(parent.index(parent_run), new_element)
    parent.remove(parent_run)

def process_xml_with_conversion_rules(xml_file, log_filename):
    # ログファイルを開く
    with open(log_filename, 'w', encoding='utf-8') as log_file:
        tree = ET.parse(xml_file)
        root = tree.getroot()

        for paragraph in root.findall('.//w:p', namespaces):
            # <w:p>内部のすべての<w:t>要素を結合
            full_text = "".join(text_elem.text for text_elem in paragraph.findall('.//w:t', namespaces) if text_elem.text)

            # 変換条件に該当するテキストが含まれているかチェック
            if any(re.search(re.escape(key), full_text) for rule in conversion_rules for key in rule['pattern'].keys()):
                # 該当する<w:t>要素を切り分け、ハイライトを追加
                for text_elem in paragraph.findall('.//w:t', namespaces):
                    if text_elem.text:
                        split_and_highlight_with_rules(text_elem, log_file, conversion_rules)

    tree.write(xml_file, encoding='utf-8', xml_declaration=True, pretty_print=True)

# 使用例
process_xml_with_conversion_rules('xml_new/word/document.xml', 'conversion_rules_log.txt')
