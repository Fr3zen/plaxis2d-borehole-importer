# 【模块中文说明】
# 钻孔导入的独立软件工作进程。
# 外层程序可使用几何处理环境，而本脚本必须由与 PLAXIS 匹配的 Python 运行，避免 SDK 版本冲突。
# 接收已验证 JSON 和新工程路径，调用 import_native，将回读结果写入指定 result 文件。
# 接口异常打印完整追踪并返回非零退出码，供父进程判定失败；未包含计算调用。
# 
"""Run only under the Python distribution shipped with the installed PLAXIS."""
import argparse
import json
from pathlib import Path
import sys
import traceback

from excel_to_plaxis import import_native

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--config', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    parser.add_argument('--result', type=Path, required=True)
    parser.add_argument('--plaxis', type=Path, required=True)
    parser.add_argument('--mesh', action='store_true')
    args = parser.parse_args()
    try:
        data = json.loads(args.config.read_text(encoding='utf-8'))
        result = import_native(data, args.output, args.plaxis, args.mesh)
        args.result.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
        print(f'已保存工程：{args.output}', flush=True)
    except Exception:
        traceback.print_exc()
        sys.exit(1)
