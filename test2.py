import re
from lxml import etree as ET
import copy

namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
ET.register_namespace('w', namespaces['w'])

# 変換ルールと対応する変換関数の定義
conversion_rules = [
    # {
    #     'name': 'Full-width to Half-width Symbols',
    #     'pattern': r'[（）。、：；！？]',
    #     'replace': lambda match: {
    #         '（': '(', '）': ')', '。': '.', '、': ',', '：': ':', '；': ';', '！': '!', '？': '?'
    #     }[match.group()],
    #     'color': 'yellow'
    # },
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

def apply_conversion_rules(text, rules, log_file):
    modified_text = text
    conversion_info = []
    conversion_occurred = False  # 変換が発生したかどうかを追跡
    
    log_file.write("Original Text: {}\n".format(text))  # テキストの初期状態をログに記録

    for rule in rules:
        pattern = rule['pattern']
        replace_function = rule['replace']
        
        # デバッグ情報: 置換前にログを記録
        log_file.write(f"Applying Rule: {rule['name']} with pattern {pattern}\n")
        
        new_text, num_subs = re.subn(pattern, replace_function, modified_text)
        
        # デバッグ情報: 置換結果をログに記録
        log_file.write(f"Text after applying rule: {new_text}\n")
        
        if num_subs > 0:
            conversion_occurred = True
            conversion_info.append((rule['name'], pattern, rule['color'], new_text))
        modified_text = new_text
    
    # 変換が発生しなかった場合のログ出力用情報
    if not conversion_occurred:
        log_file.write("No conversion occurred.\n")  # 変換が発生しなかったことをログに記録
        conversion_info.append(('No conversion', 'No Rule Matched', '', modified_text))
    
    log_file.write("Modified Text after all rules: {}\n".format(modified_text))  # 変換後のテキストをログに記録
    return modified_text, conversion_info

def split_and_highlight_with_rules(text_element, log_file, rules):
    parent_run = text_element.getparent()
    original_rpr = parent_run.find('.//w:rPr', namespaces)
    
    # <w:t>要素を結合してフルテキストを作成
    combined_text = "".join([t.text for t in parent_run.findall('.//w:t', namespaces) if t.text])
    log_file.write("Combined Text: {}\n".format(combined_text))  # 結合されたテキストをログに記録
    
    # 変換ルールを適用
    modified_text, conversion_info = apply_conversion_rules(combined_text, rules, log_file)
    
    # ログファイルに変更前後の情報を書き出し
    for info in conversion_info:
        log_file.write("Rule Applied: {}\n".format(info[0]))
        log_file.write("Pattern: {}\n".format(info[1]))
        log_file.write("Modified Text: {}\n\n".format(info[3]))

    # 元の要素を削除する前に、インデックスを取得
    parent = parent_run.getparent()
    index = parent.index(parent_run)
    
    # 元の要素を削除
    for child in list(parent_run):
        parent_run.remove(child)
    
    # 新しい要素を挿入
    current_position = 0
    new_elements = []

    for match in re.finditer('|'.join(re.escape(rule['replace'](match)) for rule in conversion_rules for match in re.finditer(rule['pattern'], combined_text)), modified_text):
        start, end = match.span()
        
        # ハイライトされない部分を追加
        if current_position < start:
            unhighlighted_text = modified_text[current_position:start]
            if unhighlighted_text:
                new_elements.append(create_plain_run(original_rpr, unhighlighted_text))
        
        # ハイライトされた部分を追加
        for rule in conversion_rules:
            if re.match(rule['pattern'], match.group()):
                new_elements.append(create_highlighted_run(original_rpr, match.group(), rule['color']))
                break
        current_position = end

    # 残りの部分を追加
    if current_position < len(modified_text):
        remaining_text = modified_text[current_position:]
        new_elements.append(create_plain_run(original_rpr, remaining_text))
    
    # 新しい要素を親に挿入
    for new_element in new_elements:
        parent.insert(index, new_element)
        index += 1  # インデックスを更新

    # 元の要素を親から削除
    parent.remove(parent_run)


def process_xml_with_conversion_rules(xml_file, log_filename):
    with open(log_filename, 'w', encoding='utf-8') as log_file:
        log_file.write("Processing XML File: {}\n".format(xml_file))  # ファイル処理の開始をログに記録
        tree = ET.parse(xml_file)
        root = tree.getroot()

        for paragraph in root.findall('.//w:p', namespaces):
            full_text = "".join(text_elem.text for text_elem in paragraph.findall('.//w:t', namespaces) if text_elem.text)
            log_file.write("Full Paragraph Text: {}\n".format(full_text))  # 段落全体のテキストをログに記録

            if full_text:
                log_file.write(f"Processing Paragraph: {full_text}\n")
                
            # 変換条件に該当するテキストが含まれているかチェック
            applied_rules = []
            for rule in conversion_rules:
                match = re.search(rule['pattern'], full_text)
                if match:
                    log_file.write(f"Matched Rule: {rule['name']}\n")
                    log_file.write(f"Pattern: {rule['pattern']}\n")
                    log_file.write(f"Matched Text: {full_text}\n\n")
                    applied_rules.append(rule['name'])

            if applied_rules:
                # <w:t>要素を結合して処理
                for text_elem in paragraph.findall('.//w:t', namespaces):
                    if text_elem.text:
                        split_and_highlight_with_rules(text_elem, log_file, conversion_rules)
            else:
                log_file.write("No rules applied for this paragraph.\n")
        
        log_file.write("Finished processing XML file.\n")  # ファイル処理の終了をログに記録
        tree.write(xml_file, encoding='utf-8', xml_declaration=True, pretty_print=True)


# 使用例
process_xml_with_conversion_rules('xml_new/word/document.xml', 'conversion_rules_log.txt')