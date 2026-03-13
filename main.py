import os
import platform
import subprocess
import threading
import json
import webview

from core.parser import parse_files
from core.exporter import export_xlsx
from core.classifier import create_output_folder, classify_files

SUPPORTED_EXT = {'.txt', '.docx', '.doc'}

HTML = '''
<!DOCTYPE html>
<html lang="zh">
<head>
<meta charset="UTF-8">
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body {
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
  background: #F6F8FA; color: #24292F; height: 100vh; display: flex; flex-direction: column;
  user-select: none; -webkit-user-select: none;
}

/* Header */
.header {
  background: #fff; border-bottom: 1px solid #D0D7DE; padding: 16px 24px;
  display: flex; align-items: center; justify-content: space-between; flex-shrink: 0;
}
.header h1 { font-size: 18px; font-weight: 600; }
.header h1 span { color: #0969DA; }
.header p { font-size: 13px; color: #656D76; }

/* Toolbar */
.toolbar {
  padding: 12px 24px; display: flex; align-items: center; gap: 8px; flex-shrink: 0;
}
.btn {
  display: inline-flex; align-items: center; gap: 6px;
  padding: 7px 16px; border-radius: 6px; font-size: 13px; font-weight: 500;
  border: 1px solid transparent; cursor: pointer; transition: all .15s;
}
.btn-primary { background: #0969DA; color: #fff; border-color: #0969DA; }
.btn-primary:hover { background: #0757B5; }
.btn-outline { background: #fff; color: #24292F; border-color: #D0D7DE; }
.btn-outline:hover { background: #F3F4F6; }
.file-count { margin-left: auto; font-size: 13px; color: #656D76; }

/* File List */
.list-wrap {
  flex: 1; padding: 0 24px 12px; overflow: hidden; display: flex; flex-direction: column;
}
.list-card {
  flex: 1; background: #fff; border: 1px solid #D0D7DE; border-radius: 8px;
  overflow: hidden; display: flex; flex-direction: column;
}
.list-header {
  display: grid; grid-template-columns: 48px 1fr 1fr; gap: 0;
  background: #F6F8FA; border-bottom: 1px solid #D0D7DE;
  font-size: 12px; font-weight: 600; color: #656D76; text-transform: uppercase;
  letter-spacing: .3px;
}
.list-header > div { padding: 8px 12px; }
.list-body { flex: 1; overflow-y: auto; }
.list-row {
  display: grid; grid-template-columns: 48px 1fr 1fr; gap: 0;
  border-bottom: 1px solid #F0F0F0; font-size: 13px; transition: background .1s;
}
.list-row:hover { background: #F6F8FA; }
.list-row > div {
  padding: 8px 12px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
}
.list-row .idx { color: #656D76; text-align: center; }
.empty-state {
  flex: 1; display: flex; align-items: center; justify-content: center;
  flex-direction: column; color: #8C959F; gap: 8px;
}
.empty-state svg { opacity: .4; }
.empty-state p { font-size: 14px; }

/* Footer */
.footer {
  background: #fff; border-top: 1px solid #D0D7DE; padding: 12px 24px;
  display: flex; align-items: center; gap: 16px; flex-shrink: 0;
}
.btn-run {
  background: #1A7F37; color: #fff; border-color: #1A7F37; padding: 8px 24px; font-size: 14px;
}
.btn-run:hover { background: #15692E; }
.btn-run:disabled { background: #94D3A2; cursor: not-allowed; border-color: #94D3A2; }
.progress-wrap { flex: 1; }
.progress-bar {
  width: 100%; height: 6px; background: #E1E4E8; border-radius: 3px; overflow: hidden;
}
.progress-fill {
  height: 100%; background: #0969DA; border-radius: 3px;
  transition: width .3s; width: 0%;
}
.status { font-size: 13px; color: #656D76; min-width: 80px; text-align: right; }
.tags { display: flex; gap: 8px; }
.tag {
  font-size: 12px; font-weight: 600; padding: 2px 10px; border-radius: 12px;
}
.tag-ng { background: #FFEBE9; color: #CF222E; }
.tag-ok { background: #DAFBE1; color: #1A7F37; }
</style>
</head>
<body>

<div class="header">
  <h1><span>Weak Short</span> 批量审核工具</h1>
  <p>自动提取并归类 Weak Short-Circuit Test 结果</p>
</div>

<div class="toolbar">
  <button class="btn btn-primary" onclick="selectFiles()">
    <svg width="14" height="14" viewBox="0 0 16 16" fill="currentColor"><path d="M3.75 0A1.75 1.75 0 002 1.75v12.5c0 .966.784 1.75 1.75 1.75h8.5A1.75 1.75 0 0014 14.25V4.664a1.75 1.75 0 00-.513-1.237L10.573.513A1.75 1.75 0 009.336 0H3.75zM3.5 1.75a.25.25 0 01.25-.25h5.586a.25.25 0 01.177.073l2.914 2.914a.25.25 0 01.073.177V14.25a.25.25 0 01-.25.25h-8.5a.25.25 0 01-.25-.25V1.75z"/></svg>
    选择文件
  </button>
  <button class="btn btn-primary" onclick="selectFolder()">
    <svg width="14" height="14" viewBox="0 0 16 16" fill="currentColor"><path d="M.54 3.87L.5 3a2 2 0 012-2h3.672a2 2 0 011.414.586l.828.828A2 2 0 009.828 3H13.5a2 2 0 012 2v.054l-15-.184zM16 6.5H0v5.75A2.75 2.75 0 002.75 15h10.5A2.75 2.75 0 0016 12.25V6.5z"/></svg>
    选择文件夹
  </button>
  <button class="btn btn-outline" onclick="clearList()">清空</button>
  <span class="file-count" id="fileCount">尚未选择文件</span>
</div>

<div class="list-wrap">
  <div class="list-card">
    <div class="list-header">
      <div>#</div><div>文件名</div><div>路径</div>
    </div>
    <div class="list-body" id="listBody">
      <div class="empty-state" id="emptyState">
        <svg width="40" height="40" viewBox="0 0 16 16" fill="currentColor"><path d="M3.75 0A1.75 1.75 0 002 1.75v12.5c0 .966.784 1.75 1.75 1.75h8.5A1.75 1.75 0 0014 14.25V4.664a1.75 1.75 0 00-.513-1.237L10.573.513A1.75 1.75 0 009.336 0H3.75zM3.5 1.75a.25.25 0 01.25-.25h5.586a.25.25 0 01.177.073l2.914 2.914a.25.25 0 01.073.177V14.25a.25.25 0 01-.25.25h-8.5a.25.25 0 01-.25-.25V1.75z"/></svg>
        <p>点击上方按钮选择测试文件或文件夹</p>
      </div>
    </div>
  </div>
</div>

<div class="footer">
  <button class="btn btn-run" id="runBtn" onclick="runProcess()">开始处理</button>
  <div class="progress-wrap">
    <div class="progress-bar"><div class="progress-fill" id="progressFill"></div></div>
  </div>
  <div class="tags">
    <span class="tag tag-ng" id="ngTag">NG: -</span>
    <span class="tag tag-ok" id="okTag">OK: -</span>
  </div>
  <span class="status" id="statusText">就绪</span>
</div>

<script>
let files = [];

function renderList() {
  const body = document.getElementById('listBody');
  const empty = document.getElementById('emptyState');
  const count = document.getElementById('fileCount');

  if (files.length === 0) {
    body.innerHTML = '';
    body.appendChild(empty);
    empty.style.display = 'flex';
    count.textContent = '尚未选择文件';
    return;
  }

  empty.style.display = 'none';
  let html = '';
  files.forEach((f, i) => {
    const name = f.split(/[\\/]/).pop();
    html += '<div class="list-row"><div class="idx">' + (i+1) +
            '</div><div>' + name + '</div><div title="' + f + '">' + f + '</div></div>';
  });
  body.innerHTML = html;
  count.textContent = '已选 ' + files.length + ' 个文件';
}

async function selectFiles() {
  const result = await pywebview.api.select_files();
  if (result && result.length > 0) {
    const existing = new Set(files);
    result.forEach(f => { if (!existing.has(f)) files.push(f); });
    renderList();
  }
}

async function selectFolder() {
  const result = await pywebview.api.select_folder();
  if (result && result.length > 0) {
    const existing = new Set(files);
    result.forEach(f => { if (!existing.has(f)) files.push(f); });
    renderList();
  }
}

function clearList() {
  files = [];
  renderList();
  document.getElementById('ngTag').textContent = 'NG: -';
  document.getElementById('okTag').textContent = 'OK: -';
  document.getElementById('statusText').textContent = '就绪';
  document.getElementById('progressFill').style.width = '0%';
}

async function runProcess() {
  if (files.length === 0) {
    alert('请先选择要处理的文件');
    return;
  }
  const btn = document.getElementById('runBtn');
  const status = document.getElementById('statusText');
  const progress = document.getElementById('progressFill');

  btn.disabled = true;
  status.textContent = '处理中...';
  status.style.color = '#0969DA';
  progress.style.width = '20%';

  try {
    const result = await pywebview.api.process_files(files);
    progress.style.width = '100%';

    document.getElementById('ngTag').textContent = 'NG: ' + result.ng_count;
    document.getElementById('okTag').textContent = 'OK: ' + result.ok_count;

    if (result.success) {
      status.textContent = '完成';
      status.style.color = '#1A7F37';
      let msg = '处理完成！\\n\\nNG: ' + result.ng_count + ' 个\\nOK: ' + result.ok_count +
                ' 个\\n异常: ' + result.error_count + ' 个\\n\\n结果目录:\\n' + result.output_dir;
      if (result.errors && result.errors.length > 0) {
        msg += '\\n\\n异常详情:\\n' + result.errors.join('\\n');
      }
      alert(msg);
      await pywebview.api.open_folder(result.output_dir);
    } else {
      status.textContent = '出错';
      status.style.color = '#CF222E';
      alert('处理出错: ' + result.error);
    }
  } catch (e) {
    status.textContent = '出错';
    status.style.color = '#CF222E';
    alert('处理出错: ' + e);
  } finally {
    btn.disabled = false;
  }
}

renderList();
</script>
</body>
</html>
'''


class Api:
    def __init__(self, window):
        self._window = window

    def select_files(self):
        result = self._window.create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=True,
            file_types=('测试文件 (*.txt;*.docx)', '所有文件 (*.*)'),
        )
        if result:
            return [f for f in result if os.path.splitext(f)[1].lower() in SUPPORTED_EXT]
        return []

    def select_folder(self):
        result = self._window.create_file_dialog(webview.FOLDER_DIALOG)
        if not result:
            return []
        folder = result[0] if isinstance(result, (list, tuple)) else result
        paths = []
        for root, _, fnames in os.walk(folder):
            for f in fnames:
                if os.path.splitext(f)[1].lower() in SUPPORTED_EXT:
                    paths.append(os.path.join(root, f))
        return paths

    def process_files(self, file_list):
        try:
            results = parse_files(file_list)

            ng = [r for r in results if r.status == 'NG']
            ok = [r for r in results if r.status == 'OK']
            errors = [r for r in results if r.status == 'ERROR']

            root_dir, ng_dir, ok_dir, xlsx_path = create_output_folder()
            classify_files(results, ng_dir, ok_dir)
            export_xlsx(ng, ok, xlsx_path)

            return {
                'success': True,
                'ng_count': len(ng),
                'ok_count': len(ok),
                'error_count': len(errors),
                'errors': [f'{r.filename}: {r.error}' for r in errors],
                'output_dir': root_dir,
            }
        except Exception as e:
            return {'success': False, 'error': str(e)}

    def open_folder(self, path):
        system = platform.system()
        if system == 'Windows':
            os.startfile(path)
        elif system == 'Darwin':
            subprocess.Popen(['open', path])
        else:
            subprocess.Popen(['xdg-open', path])


def main():
    window = webview.create_window(
        'Weak Short 批量审核工具',
        html=HTML,
        width=820,
        height=580,
        min_size=(650, 450),
    )
    api = Api(window)
    window.expose(api.select_files)
    window.expose(api.select_folder)
    window.expose(api.process_files)
    window.expose(api.open_folder)
    webview.start()


if __name__ == '__main__':
    main()
