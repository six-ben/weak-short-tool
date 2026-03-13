import re
import os
import docx


def read_file_text(filepath):
    """读取 txt 或 docx 文件的纯文本内容"""
    ext = os.path.splitext(filepath)[1].lower()

    if ext == '.docx':
        doc = docx.Document(filepath)
        return '\n'.join(p.text for p in doc.paragraphs)

    encodings = ['utf-8', 'gbk', 'gb2312', 'latin-1']
    for enc in encodings:
        try:
            with open(filepath, 'r', encoding=enc) as f:
                return f.read()
        except (UnicodeDecodeError, UnicodeError):
            continue

    with open(filepath, 'r', encoding='utf-8', errors='replace') as f:
        return f.read()


def extract_weak_short_section(text):
    """提取 Weak Short-Circuit Test 到 INT Pin Test 之间的内容"""
    pattern = (
        r'={5,}Test Item:[\s\-]*Weak Short-Circuit Test\s*'
        r'(.*?)'
        r'={5,}Test Item:[\s\-]*INT Pin Test'
    )
    match = re.search(pattern, text, re.DOTALL | re.IGNORECASE)
    if match:
        return match.group(1)
    return None


def judge_result(section_text):
    """判定 NG 或 OK，返回 'NG' / 'OK' / None"""
    if re.search(r'Weak\s*Short.*?NG', section_text, re.IGNORECASE):
        return 'NG'
    if re.search(r'Weak\s*Short.*?OK', section_text, re.IGNORECASE):
        return 'OK'
    return None


def extract_mul_short(section_text):
    """提取 Mul Short: 下的所有数据行"""
    pattern = r'Mul\s*Short:\s*\n(.*?)(?=Error Ground|Mutual\s*Short:|GND\s*Short:|//|$)'
    match = re.search(pattern, section_text, re.DOTALL | re.IGNORECASE)
    if not match:
        return ''

    raw = match.group(1).strip()
    lines = [line.strip() for line in raw.split('\n') if line.strip()]
    return '\n'.join(lines)


def extract_mutual_short(section_text):
    """提取 Mutual Short: 下的所有数据行"""
    pattern = r'Mutual\s*Short:\s*\n(.*?)(?=//|$)'
    match = re.search(pattern, section_text, re.DOTALL | re.IGNORECASE)
    if not match:
        return ''

    raw = match.group(1).strip()
    lines = [line.strip() for line in raw.split('\n') if line.strip()]
    return '\n'.join(lines)


def extract_gnd_short(section_text):
    """提取 GND Short: 下的所有数据行"""
    pattern = r'GND\s*Short:\s*\n(.*?)(?=//|$)'
    match = re.search(pattern, section_text, re.DOTALL | re.IGNORECASE)
    if not match:
        return ''

    raw = match.group(1).strip()
    lines = [line.strip() for line in raw.split('\n') if line.strip()]
    return '\n'.join(lines)


class ParseResult:
    def __init__(self, filepath, filename, status, mul_short='', mutual_short='',
                 result2_type='', error=''):
        self.filepath = filepath
        self.filename = filename
        self.status = status        # 'NG' / 'OK' / 'ERROR'
        self.mul_short = mul_short
        self.mutual_short = mutual_short
        self.result2_type = result2_type  # 'Mutual Short' / 'GND Short' / ''
        self.error = error

    def __repr__(self):
        return f'<ParseResult {self.filename} -> {self.status}>'


def parse_file(filepath):
    """解析单个文件，返回 ParseResult"""
    filename = os.path.basename(filepath)

    try:
        text = read_file_text(filepath)
    except Exception as e:
        return ParseResult(filepath, filename, 'ERROR', error=f'读取失败: {e}')

    section = extract_weak_short_section(text)
    if section is None:
        return ParseResult(filepath, filename, 'ERROR',
                           error='未找到 Weak Short-Circuit Test 区段')

    status = judge_result(section)
    if status is None:
        return ParseResult(filepath, filename, 'ERROR',
                           error='无法判定 NG/OK 结果')

    mul_short = ''
    mutual_short = ''
    result2_type = ''
    if status == 'NG':
        mul_short = extract_mul_short(section)
        mutual_short = extract_mutual_short(section)
        if mutual_short:
            result2_type = 'Mutual Short'
        else:
            mutual_short = extract_gnd_short(section)
            if mutual_short:
                result2_type = 'GND Short'

    return ParseResult(filepath, filename, status, mul_short, mutual_short, result2_type)


def parse_files(file_list):
    """批量解析文件列表，返回 ParseResult 列表"""
    results = []
    for fp in file_list:
        results.append(parse_file(fp))
    return results
