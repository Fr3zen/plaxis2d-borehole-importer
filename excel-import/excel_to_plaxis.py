# 【模块中文说明】
# Excel 钻孔导入的主入口，兼有命令行与文件选择窗口。
# 数据流程：读取 XLSX 原始值 → 验证统一层序和几何 → 输出预览 → 交给匹配版本的 PLAXIS Python 写入。
# 读表不依赖 Excel 安装；公式单元格被拒绝，避免把过时缓存当作真实标高。
# 原生导入只建立钻孔及地层几何，不自动赋予完整材料或地下水工况；零厚度层保留原始层位身份。
# 默认不划分网格、不计算；同名工程、配套数据目录和诊断目录都不会覆盖。
# 
"""Read borehole elevations from XLSX and save a native PLAXIS 2D project.

No Excel installation/reader dependency required. No binary .p2dx rewriting:
the official locally installed plxscripting client creates and saves the model.
"""
import argparse
import ctypes
from ctypes import wintypes
from datetime import datetime
import hashlib
import json
import math
from pathlib import Path
import posixpath
import re
import secrets
import socket
import subprocess
import sys
import time
import xml.etree.ElementTree as ET
import zipfile

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT.parent))

NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
RID = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'


# 按工作表名解析 XLSX 内部 ZIP/XML，返回 {(行号, 列号): (原始值, Excel地址)}。
# 支持共享字符串、内联字符串、布尔值和数值；数值统一解析为 float，后续整数约束必须考虑这一点。
# 限制解压体积和表格范围，并保留公式标记交给业务解析器拒绝，不使用公式缓存。
def read_xlsx(path, sheet_name='钻孔数据'):
    """Read literal XLSX values; reject formulas to avoid stale calculation caches."""
    path = Path(path)
    if path.suffix.lower() != '.xlsx':
        raise ValueError('请使用 .xlsx 文件；旧版 .xls 请先另存为 .xlsx。')
    with zipfile.ZipFile(path) as archive:
        if sum(info.file_size for info in archive.infolist()) > 64 * 1024 * 1024:
            raise ValueError('Excel 解压后超过 64 MB，请只保留需要导入的数据。')
        wb = ET.fromstring(archive.read('xl/workbook.xml'))
        sheets = wb.findall('s:sheets/s:sheet', NS)
        sheet = next((s for s in sheets if s.get('name') == sheet_name), None)
        if sheet is None:
            raise ValueError(f'找不到工作表“{sheet_name}”。现有工作表：{[s.get("name") for s in sheets]}')
        relationships = ET.fromstring(archive.read('xl/_rels/workbook.xml.rels'))
        relationship = next(r for r in relationships if r.get('Id') == sheet.get(RID))
        if relationship.get('TargetMode') == 'External':
            raise ValueError('不支持外部工作表。')
        target = relationship.get('Target')
        target = target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/' + target)
        strings = []
        if 'xl/sharedStrings.xml' in archive.namelist():
            strings = [''.join(item.itertext()) for item in ET.fromstring(archive.read('xl/sharedStrings.xml'))]
        xml = ET.fromstring(archive.read(target))
        cells = {}
        for row in xml.findall('s:sheetData/s:row', NS):
            for cell in row.findall('s:c', NS):
                ref = cell.get('r')
                # Template instructions in H onward are intentionally not imported.
                match = re.fullmatch(r'([A-Z]+)(\d+)', ref or '')
                if not match:
                    continue
                col = 0
                for letter in match.group(1):
                    col = col * 26 + ord(letter) - 64
                rownum = int(match.group(2))
                if col > 256 or rownum > 10001:
                    raise ValueError('数据范围超过 256 列或 10000 个钻孔。')
                # XLSX 可同时保存公式及上次计算结果；这里宁可拒绝公式，也不把缓存误当作最新钻孔标高。
                if cell.find('s:f', NS) is not None:
                    value = {'formula': True}
                elif cell.get('t') == 'inlineStr':
                    value = ''.join(t.text or '' for t in cell.findall('s:is//s:t', NS))
                else:
                    raw = cell.findtext('s:v', default='', namespaces=NS)
                    typ = cell.get('t')
                    if typ == 's':
                        value = strings[int(raw)] if raw else None
                    elif typ in ('str', 'e', 'd'):
                        value = raw or None
                    elif typ == 'b':
                        value = raw == '1'
                    else:
                        value = float(raw) if raw else None
                cells[rownum, col] = (value, ref)
    return cells


# 仅用于表头比较：移除空白并统一全角/半角括号。
# 不改写钻孔编号、层名数据、坐标或材料参数。
def normalize(text):
    return re.sub(r'\s+', '', str(text)).replace('（', '(').replace('）', ')')


# 把钻孔表单元格转换为 layers/boreholes 结构，并交给几何生成器再校验。
# 前三列为孔编号、距离、孔口标高；后续连续的层底列定义统一层序，层底相等允许零厚度。
# 空行跳过，缺值/公式/重复层名或表头错误抛异常，不把未知标高替换成 0。
def parse_cells(cells, name='boreholes'):
    from generate_strata import build_model
    # 读取指定单元格的字面值及 Excel 地址，缺项保留 None；若遇公式标记立即报错，避免后续当作有效输入。
    def value(row, col):
        v, ref = cells.get((row, col), (None, f'第{row}行第{col}列'))
        if isinstance(v, dict):
            raise ValueError(f'{ref} 是公式。请复制后粘贴为数值再导入。')
        return v, ref
    # 读取指定单元格并要求有限的实数，错误中保留 Excel 单元格位置。
    # 显式排除 bool，因为 Python 的 bool 继承 int，不能让勾选框代替标高。
    def number(row, col):
        v, ref = value(row, col)
        if isinstance(v, bool) or not isinstance(v, (int, float)) or not math.isfinite(v):
            raise ValueError(f'{ref} 必须填写数值标高/距离，当前值为 {v!r}。')
        return float(v)
    headers = [normalize(value(1, c)[0] or '') for c in range(1, 4)]
    if headers[0] != '钻孔编号' or headers[1] not in ('沿剖面距离x(m)', '剖面位置x(m)') or headers[2] not in ('孔口标高(m)', '地表标高(m)'):
        raise ValueError('A1:C1 应为：钻孔编号、沿剖面距离 x（m）、孔口标高（m）。请使用提供的模板。')
    layers = []
    colors = ['#D7B377', '#B78775', '#E8D69B', '#8CB8B3', '#B6A2C8', '#B4C990']
    col = 4
    while True:
        header = normalize(value(1, col)[0] or '')
        if not header:
            break
        match = re.fullmatch(r'(.+)层底标高\(m\)', header)
        if not match:
            raise ValueError(f'第 {col} 列列名应为“某土层底标高（m）”；数据表右侧请留一空列。')
        layer_name = match.group(1)
        if layer_name in [layer['name'] for layer in layers]:
            raise ValueError(f'地层名称重复：{layer_name}')
        layers.append({'id': f'layer_{len(layers)+1}', 'name': layer_name, 'color': colors[len(layers) % len(colors)]})
        col += 1
    if not layers:
        raise ValueError('至少需要一列层底标高。')
    # A blank column terminates the table, but must not silently hide another layer column.
    for (r, c), (v, _) in cells.items():
        if r == 1 and c > col and isinstance(v, str) and '层底标高' in v:
            raise ValueError('层底标高列中间有空列，请连续排列。')
    holes = []
    for row in range(2, max((r for r, _ in cells), default=1) + 1):
        values = [value(row, c)[0] for c in range(1, col)]
        if all(v is None or v == '' for v in values):
            continue
        ident, ref = value(row, 1)
        if ident is None or isinstance(ident, bool) or str(ident).strip() == '':
            raise ValueError(f'{ref} 缺少钻孔编号。')
        holes.append({'id': str(ident).strip(), 'x': number(row, 2), 'ground': number(row, 3),
                      'bottoms': [number(row, c) for c in range(4, col)]})
    data = {'name': name, 'units': 'm', 'layers': layers, 'boreholes': holes}
    # build_model checks unique IDs/x, descending levels, intersections and coverage.
    build_model(data)
    return data


# 组合 XLSX 读取和钻孔解析，以文件名作为该次输入数据集名称，返回已完成离线几何校验的数据。
def load_excel(path, sheet='钻孔数据'):
    return parse_cells(read_xlsx(path, sheet), Path(path).stem)


# 通过 Windows Toolhelp 进程快照检查是否存在 PLAXIS 2D Input。
# 此检查不依赖 WMI，也不连接/关闭任何工程；系统枚举失败时直接报错，不假定软件未运行。
def plaxis_running():
    # Toolhelp needs no WMI service and works when tasklist/WMI is unavailable.
    # Windows PROCESSENTRY32 结构声明；字段布局必须与系统 API 对齐，特别是指针宽度和宽字符进程名。
    class ProcessEntry(ctypes.Structure):
        _fields_ = [('dwSize', wintypes.DWORD), ('cntUsage', wintypes.DWORD),
                    ('th32ProcessID', wintypes.DWORD), ('th32DefaultHeapID', ctypes.c_size_t),
                    ('th32ModuleID', wintypes.DWORD), ('cntThreads', wintypes.DWORD),
                    ('th32ParentProcessID', wintypes.DWORD), ('pcPriClassBase', wintypes.LONG),
                    ('dwFlags', wintypes.DWORD), ('szExeFile', wintypes.WCHAR * 260)]
    kernel = ctypes.WinDLL('kernel32', use_last_error=True)
    kernel.CreateToolhelp32Snapshot.argtypes = [wintypes.DWORD, wintypes.DWORD]
    kernel.CreateToolhelp32Snapshot.restype = wintypes.HANDLE
    kernel.Process32FirstW.argtypes = [wintypes.HANDLE, ctypes.POINTER(ProcessEntry)]
    kernel.Process32NextW.argtypes = [wintypes.HANDLE, ctypes.POINTER(ProcessEntry)]
    kernel.CloseHandle.argtypes = [wintypes.HANDLE]
    snapshot = kernel.CreateToolhelp32Snapshot(2, 0)
    if snapshot == ctypes.c_void_p(-1).value:
        raise ctypes.WinError(ctypes.get_last_error())
    try:
        entry = ProcessEntry()
        entry.dwSize = ctypes.sizeof(entry)
        found = kernel.Process32FirstW(snapshot, ctypes.byref(entry))
        if not found:
            raise ctypes.WinError(ctypes.get_last_error())
        while found:
            if entry.szExeFile.lower() == 'plaxis2dxinput.exe':
                return True
            found = kernel.Process32NextW(snapshot, ctypes.byref(entry))
        return False
    finally:
        kernel.CloseHandle(snapshot)


# 验证程序存在且没有已有 Input，随后用随机本地端口和临时认证信息启动新进程。
# 轮询接口就绪并确认是 2D，返回 server、全局对象 g 及进程；退出或超时报错。
# 认证信息不加入模型报告，许可/登录交互由用户在软件中处理。
def launch_plaxis(executable, timeout=180):
    from plxscripting.easy import new_server
    if not Path(executable).is_file():
        raise ValueError(f'找不到 PLAXIS Input：{executable}')
    if plaxis_running():
        raise ValueError('检测到 PLAXIS Input 已运行。请先保存并关闭 Input，再运行导入，以免影响已有工程。')
    with socket.socket() as sock:
        sock.bind(('127.0.0.1', 0))
        port = sock.getsockname()[1]
    password = secrets.token_hex(16)
    process = subprocess.Popen([str(executable), f'--AppServerPort={port}', f'--AppServerPassword={password}'],
                               cwd=str(Path(executable).parent))
    print('正在启动 PLAXIS 并等待本机脚本接口；如出现许可或登录提示，请在软件中处理。', flush=True)
    start = time.monotonic()
    last_error = None
    while time.monotonic() - start < timeout:
        try:
            server, global_obj = new_server('127.0.0.1', port, password=password, timeout=2, request_timeout=30)
            if server.active:
                if not server.is_2d:
                    raise ValueError('连接的不是 PLAXIS 2D。')
                return server, global_obj, process
        except Exception as error:
            last_error = error
        if process.poll() is not None:
            raise RuntimeError('PLAXIS 在接口就绪前已退出。请确认程序启动和脚本功能可用。')
        time.sleep(1)
    raise TimeoutError(f'等待 PLAXIS 脚本接口超时。未写入模型，未关闭软件。原因：{type(last_error).__name__}')


# 用已验证钻孔数据创建新的原生地层工程，返回软件版本、层序和逐孔标高回读记录。
# 先分配逻辑层，再完整设置各孔实际标高；种子厚度不会保留成虚构土层。
# 只在显式 mesh=True 时划分网格；不设置完整本构/水条件，不计算，也不覆盖已有工程。
def import_native(data, project, executable, mesh=False):
    project = Path(project).resolve()
    if project.suffix.lower() != '.p2dx':
        raise ValueError('输出工程扩展名必须是 .p2dx。')
    if project.exists() or Path(str(project) + 'dat').exists():
        raise ValueError(f'输出已存在，拒绝覆盖：{project}。请选择新文件名。')
    holes = sorted(data['boreholes'], key=lambda h: h['x'])
    bounds = (holes[0]['x'], min(h['bottoms'][-1] for h in holes),
              holes[-1]['x'], max(h['ground'] for h in holes))
    server, g, process = launch_plaxis(executable)
    # Only reset the freshly launched dedicated instance, never attach to an existing app.
    server.new()
    g.setproperties('Title', data['name'], 'UnitLength', 'm', 'UnitForce', 'kN', 'UnitTime', 'day',
                    'ModelType', 'Plane strain', 'ElementType', '15-Noded')
    g.SoilContour.initializerectangular(*bounds)
    native_holes = [g.borehole(h['x']) for h in holes]
    # Allocate logical layers before assigning actual elevations at every hole.
    # Temporary seed thickness is replaced in full; never add thickness to source data.
    # 1 m 只用于调用 API 分配层位，下面会完整重写所有孔的层顶/层底。
    # 不能通过给零厚度层添加微小正厚度来绕过真实缺层。
    for _ in data['layers']:
        g.soillayer(1.0)
    for index, layer in enumerate(data['layers']):
        g.setproperties(g.Soillayers[index], 'Comments', layer['name'])
    for source, target in zip(holes, native_holes):
        values = []
        for i, elevation in enumerate([source['ground'], *source['bottoms']]):
            values.extend([i, elevation])
        g.setsoillayerlevel(target, *values)
        g.setproperties(target, 'Comments', source['id'])
    # Read back all borehole interface elevations before claiming success.
    # 写命令返回成功并不代表软件实际保存了相同标高。
    # 逐孔逐界面回读，绝对误差超 1e-7 m 立即停止，不生成虚假的导入成功记录。
    readback = []
    for source, target in zip(holes, native_holes):
        levels = [float(g.getsoillayerlevel(target, i)) for i in range(len(data['layers']) + 1)]
        expected = [source['ground'], *source['bottoms']]
        if any(abs(a-b) > 1e-7 for a, b in zip(levels, expected)):
            raise RuntimeError(f'{source["id"]} 标高回读不一致：{levels} != {expected}')
        readback.append({'source_id': source['id'], 'plaxis_name': str(target.Name), 'elevations': levels})
    if len(g.Boreholes) != len(holes) or len(g.Soillayers) != len(data['layers']):
        raise RuntimeError('PLAXIS 回读钻孔或地层数量不一致。')
    g.gotostructures()
    if mesh:
        g.gotomesh()
        mesh_feedback = str(g.mesh(0.06))
    else:
        mesh_feedback = None
    project.parent.mkdir(parents=True, exist_ok=True)
    g.save(str(project))
    if not project.is_file() or not Path(str(project) + 'dat').is_dir():
        raise RuntimeError('PLAXIS 未生成完整的 p2dx 工程及数据目录。')
    return {'status': 'saved', 'project': str(project), 'native_boreholes': len(holes),
            'native_layers': len(data['layers']), 'readback': readback,
            'layer_order': [layer['name'] for layer in data['layers']],
            'mesh_requested': mesh, 'mesh_feedback': mesh_feedback,
            'materials_assigned': False, 'water_configured': False,
            'application': server.server_full_name,
            'note': 'Native borehole geometry only. Layer labels are mapped here; materials must be assigned separately.'}


# 解析命令行/文件选择结果并组织用户入口流程。
# 先离线验证，再按选项调用匹配 Python 的软件工作进程；取消选择正常返回，异常返回非零状态。
# 最终成功状态必须来自工作进程结果，不能仅凭命令发出就报告模型已生成。
def main():
    parser = argparse.ArgumentParser(description='Excel 钻孔数据直接生成 PLAXIS 2D 工程（原生钻孔）。')
    parser.add_argument('--excel', type=Path, help='模板格式的 .xlsx 文件')
    parser.add_argument('--sheet', default='钻孔数据')
    parser.add_argument('--output', type=Path, help='新的 .p2dx 工程路径，不覆盖已有工程')
    parser.add_argument('--plaxis', type=Path, default=Path(r'E:\plaxis2024 2d\Plaxis2DXInput.exe'))
    bundled_plaxis_python = ROOT.parent / 'plaxis_python' / 'python.exe'
    default_plaxis_python = bundled_plaxis_python if bundled_plaxis_python.is_file() else Path(r'C:\ProgramData\Seequent\PLAXIS Python Distribution V2\python\python.exe')
    parser.add_argument('--plaxis-python', type=Path, default=default_plaxis_python,
                        help='与 PLAXIS 版本匹配的 Python 发行版')
    parser.add_argument('--validate-only', action='store_true', help='只检查 Excel 并生成预览，不启动 PLAXIS')
    parser.add_argument('--mesh', action='store_true', help='导入后划分中等网格')
    args = parser.parse_args()
    gui = args.excel is None
    root = None
    try:
        if gui:
            import tkinter as tk
            from tkinter import filedialog
            root = tk.Tk()
            root.withdraw()
            chosen = filedialog.askopenfilename(title='选择钻孔 Excel（请使用模板格式）', initialdir=ROOT,
                                               filetypes=[('Excel 工作簿', '*.xlsx')])
            if not chosen:
                return 0
            args.excel = Path(chosen)
        data = load_excel(args.excel, args.sheet)
        if args.output is None and gui and not args.validate_only:
            chosen = filedialog.asksaveasfilename(title='保存为新的 PLAXIS 工程', initialdir=ROOT / 'output',
                initialfile=f'{args.excel.stem}_{datetime.now():%Y%m%d_%H%M%S}.p2dx',
                defaultextension='.p2dx', filetypes=[('PLAXIS 2D 工程', '*.p2dx')])
            if not chosen:
                return 0
            args.output = Path(chosen)
        project = (args.output or ROOT / 'output' / f'{args.excel.stem}_{datetime.now():%Y%m%d_%H%M%S}.p2dx').resolve()
        if project.suffix.lower() != '.p2dx':
            raise ValueError('输出扩展名必须是 .p2dx。')
        if project.exists() or Path(str(project) + 'dat').exists():
            raise ValueError('目标工程已存在，请使用新文件名。')
        diagnostics = project.parent / (project.stem + '_import')
        if diagnostics.exists():
            raise ValueError('同名导入记录目录已存在，请使用新工程文件名。')
        diagnostics.mkdir(parents=True)
        config = diagnostics / 'excel_data.json'
        config.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
        from generate_strata import generate
        generate(config, diagnostics)
        report = {'status': 'excel_validated', 'excel': str(args.excel.resolve()),
                  'excel_sha256': hashlib.sha256(args.excel.read_bytes()).hexdigest(),
                  'sheet': args.sheet, 'diagnostics': str(diagnostics)}
        if not args.validate_only:
            if not args.plaxis_python.is_file():
                raise ValueError(f'找不到 PLAXIS Python：{args.plaxis_python}。请使用 --plaxis-python 指定匹配环境。')
            result_path = diagnostics / 'native_result.json'
            command = [str(args.plaxis_python), '-u', str(ROOT / 'plaxis_worker.py'),
                       '--config', str(config), '--output', str(project), '--result', str(result_path),
                       '--plaxis', str(args.plaxis)]
            if args.mesh:
                command.append('--mesh')
            completed = subprocess.run(command)
            if completed.returncode:
                raise RuntimeError(f'PLAXIS 导入未完成，退出码 {completed.returncode}。请查看上面的接口或软件提示。')
            report.update(json.loads(result_path.read_text(encoding='utf-8')))
        (diagnostics / 'import_report.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
        print(json.dumps(report, ensure_ascii=False, indent=2), flush=True)
        if root is not None:
            from tkinter import messagebox
            messagebox.showinfo('导入完成', f'已完成：{report["status"]}\n工程/记录位置：{project if not args.validate_only else diagnostics}\n材料参数和水条件需另行设置。')
        return 0
    except Exception as error:
        print(f'导入失败：{error}', file=sys.stderr, flush=True)
        if root is not None:
            from tkinter import messagebox
            messagebox.showerror('导入失败', str(error))
        return 1
    finally:
        if root is not None:
            root.destroy()


if __name__ == '__main__':
    raise SystemExit(main())
