import os
import shutil
from datetime import datetime


def create_output_folder(base_dir=None):
    """
    在指定目录（默认桌面）创建带日期的输出文件夹
    返回 (root_dir, ng_dir, ok_dir, xlsx_path)
    """
    if base_dir is None:
        base_dir = os.path.join(os.path.expanduser('~'), 'Desktop')

    date_str = datetime.now().strftime('%Y-%m-%d_%H%M%S')
    root_dir = os.path.join(base_dir, f'审核结果_{date_str}')
    ng_dir = os.path.join(root_dir, 'NG')
    ok_dir = os.path.join(root_dir, 'OK')

    os.makedirs(ng_dir, exist_ok=True)
    os.makedirs(ok_dir, exist_ok=True)

    xlsx_path = os.path.join(root_dir, 'output.xlsx')
    return root_dir, ng_dir, ok_dir, xlsx_path


def classify_files(results, ng_dir, ok_dir):
    """
    将文件按 NG / OK 复制到对应目录
    results: list[ParseResult]
    """
    for r in results:
        if r.status == 'NG':
            dst = os.path.join(ng_dir, r.filename)
            shutil.copy2(r.filepath, dst)
        elif r.status == 'OK':
            dst = os.path.join(ok_dir, r.filename)
            shutil.copy2(r.filepath, dst)
