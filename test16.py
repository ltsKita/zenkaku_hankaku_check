import regex as re  # regexモジュールを使用
from lxml import etree as ET
import glob  # globモジュールのインポート

namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
ET.register_namespace('w', namespaces['w'])

conversion_rules = [
    # 始まりカッコを全角に変換するルール（後に大文字アルファベットが続く場合、または全角英字が含まれる場合）
    {
        'name': '始まりカッコを全角に変換（後に大文字アルファベットが続く場合、または全角英字が含まれる場合）',
        'pattern': r'\((?=[Ａ-ＺA-Z])',
        'replace': lambda match: '（',  # 半角カッコを全角に変換
        'check_japanese': False
    },
    # 閉じカッコを全角に変換するルール（前に大文字アルファベット、または全角英字が含まれる場合）
    {
        'name': '閉じカッコを全角に変換（前に大文字アルファベット、または全角英字が含まれる場合）',
        'pattern': r'(?<=[Ａ-ＺA-Z])\)',
        'replace': lambda match: '）',  # 半角カッコを全角に変換
        'check_japanese': False
    },
    # 半角カッコに変換するルール（カッコ内がアルファベット・数字のみの場合）
    {
        'name': 'カッコの校閲（半角: カッコ内がアルファベット・数字のみの場合）',
        'pattern': r'（([0-9A-Za-zＡ-Ｚａ-ｚ０-９])）',
        'replace': lambda match: '(' + match.group(1).translate(str.maketrans(
            'ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ０１２３４５６７８９', 
            'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789')) + ')',
        'check_japanese': False
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
        'check_japanese': False
    },
]

def apply_conversion_rule(text, rule):
    """テキストに指定された変換ルールを適用する関数"""
    if rule['check_japanese']:
        # 日本語が含まれるか確認してから変換
        if re.search(r'[ぁ-んァ-ヶ一-龠]', text):
            return re.sub(rule['pattern'], rule['replace'], text)
    else:
        return re.sub(rule['pattern'], rule['replace'], text)
    return text

def process_runs_in_paragraph(paragraph, log_file, rules):
    """段落内のテキストに対して変換を行う関数"""
    runs = paragraph.findall('.//w:r', namespaces)
    for run in runs:
        t_elements = run.findall('.//w:t', namespaces)
        for t_element in t_elements:
            original_text = t_element.text
            new_text = original_text
            if original_text:
                for rule in rules:
                    new_text = apply_conversion_rule(new_text, rule)
                if new_text != original_text:
                    t_element.text = new_text
                    log_file.write(f"対象テキスト: '{original_text}', 適用ルール: '全ルール', 適用後テキスト: '{new_text}'\n")
                    apply_color_to_run(run, 'green')

def apply_color_to_run(run, color):
    """指定されたラン（run）にハイライトの色を適用する関数"""
    rpr = run.find('.//w:rPr', namespaces)
    if rpr is None:
        rpr = ET.SubElement(run, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
    highlight_elem = ET.SubElement(rpr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}highlight')
    highlight_elem.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', color)

def process_footer_file(file_path, log_file, rules):
    """footer名称が含まれるファイルに対して変換を行う関数"""
    tree = ET.parse(file_path)
    root = tree.getroot()
    
    # フッター内の各段落を処理
    for paragraph in root.findall('.//w:p', namespaces):
        process_runs_in_paragraph(paragraph, log_file, rules)

    tree.write(file_path, encoding='utf-8', xml_declaration=True, pretty_print=True)

def process_document_file(document_file, log_file, rules):
    """document.xmlに対して変換を行う関数"""
    tree = ET.parse(document_file)
    root = tree.getroot()
    
    # 文書内の各段落を処理
    for paragraph in root.findall('.//w:p', namespaces):
        process_runs_in_paragraph(paragraph, log_file, rules)

    tree.write(document_file, encoding='utf-8', xml_declaration=True, pretty_print=True)

def process_all_files(log_filename):
    """footer名称が含まれる全てのファイルとdocument.xmlを処理する関数"""
    with open(log_filename, 'w', encoding='utf-8') as log_file:
        # footer名称が含まれる全てのXMLファイルを取得
        footer_files = glob.glob('**/*footer*.xml', recursive=True)
        for file_path in footer_files:
            process_footer_file(file_path, log_file, conversion_rules)
        
        # document.xmlの処理
        document_file = 'xml_new/word/document.xml'
        process_document_file(document_file, log_file, conversion_rules)

# 実行部分
process_all_files('conversion_rules_log.txt')