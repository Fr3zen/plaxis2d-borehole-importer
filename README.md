将 Excel 中的钻孔位置、孔口标高和各土层层底标高转换为 PLAXIS 2D 原生钻孔及地层，减少重复录入工作。支持零厚度层和地层尖灭，生成后可继续在 PLAXIS 中编辑。

本项目通过 PLAXIS 官方 Python 脚本接口建立工程，仅包含自有导入工具代码。使用者需自行安装并配置可使用脚本接口的 PLAXIS 及其匹配 Python 环境。

## 功能

- 从 `.xlsx` 文件读取多钻孔、多层位数据，按剖面距离排序。
- 检查钻孔编号、位置、数值和地层标高顺序。
- 允许层底等于层顶，即零厚度；支持统一层位下的尖灭及再次出现。
- 生成离线剖面预览和数据检查结果。
- 创建 PLAXIS 原生钻孔和土层，回读界面标高并与输入数据比较。
- 保存新的 `.p2dx` 工程，可选生成网格。

## 使用范围与限制

- 面向 Windows 和 PLAXIS 2D。
- 需要事先统一不同钻孔的层位对应关系；不会自动判断断层或完成地质解释。
- 当前导入地层几何；土体材料、水位、渗流边界和施工阶段需另行设置。
- 不执行变形或渗流计算。
- 孔间层界面由 PLAXIS 原生分层机制连接；离线预览用于核对输入，不是软件截图。
- 不覆盖已有工程或同名导入记录目录。

开发环境曾使用 PLAXIS 2D 2024.1.0.1060、普通 Python 3.7、Shapely 2.0.7、Matplotlib 3.5.3，以及匹配 PLAXIS 的 Python 3.8 / plxscripting 1.0.4。该信息用于说明已使用的版本，不代表所有其他版本均已验证。

## 文件结构

请保留以下相对位置：

```text
plaxis2d-borehole-importer/
├─ README.md
├─ LICENSE                 # 发布者选择许可证后添加
├─ .gitignore
├─ requirements.txt
├─ generate_strata.py
└─ excel-import/
   ├─ excel_to_plaxis.py
   ├─ plaxis_worker.py
   └─ 钻孔数据模板.xlsx
```

`generate_strata.py` 必须位于 `excel-import` 的上一级目录。只复制 `excel-import` 会缺少几何处理模块。

## 安装环境

程序使用两套 Python 环境：普通 Python 负责读表、几何检查和预览；与 PLAXIS 匹配的官方 Python 环境负责调用 `plxscripting`。

在项目根目录打开 PowerShell，给普通 Python 安装依赖：

```powershell
python -m pip install -r requirements.txt
```

这里的 `python` 应指向你准备运行导入器的普通 Python。若电脑未配置该命令，请用完整路径替代，例如：

```powershell
& "D:\Python\python.exe" -m pip install -r requirements.txt
```

`requirements.txt` 中的两个依赖固定为开发时使用过的版本。若安装失败，应结合实际 Python 版本检查依赖兼容性。

PLAXIS 和对应的 Python / plxscripting 环境需由使用者自行配置。本仓库不附带 PLAXIS、官方 Python 发行版或 UDSM DLL。

## Excel 数据格式

打开 `excel-import/钻孔数据模板.xlsx`，编辑“钻孔数据”工作表。第一行为表头，从第二行开始每行一个钻孔。

| 钻孔编号 | 沿剖面距离 x（m） | 孔口标高（m） | 填土层底标高（m） | 黏土层底标高（m） | 砂土层底标高（m） |
|---|---:|---:|---:|---:|---:|
| ZK1 | 0 | 10 | 8 | 1 | -15 |
| ZK2 | 30 | 12 | 9 | -2 | -15 |
| ZK3 | 70 | 9 | 6 | 0 | -15 |

填写要求：

1. 距离和标高单位均为 m，所有标高采用相同基准。
2. x 为沿二维剖面的距离，不能直接填未经投影的平面坐标。
3. 填写层底**标高**，而非埋深：层底标高 = 孔口标高 − 层底埋深。
4. 地层列从浅到深排列，新增层位请沿用模板“某土层底标高（m）”的命名方式。
5. 各孔必须填写每个统一层位的层底；最后一层的层底为模型底界。
6. 某层缺失时，令其层底等于层顶，表示零厚度；不得留空。全剖面零厚度层保留层位次序，但不人为增加面积。
7. 层底不能高于层顶；数值 0 表示实际标高为 0，不能代替缺失数据。
8. 钻孔编号和 x 坐标必须分别唯一；行顺序可以打乱。
9. 数值单元格不要保留公式，请先复制并粘贴为值。
10. 数据表右侧说明与地层列之间至少留一列空白；新增地层时相应右移说明。
11. 只支持 `.xlsx`；旧 `.xls` 请先另存为 `.xlsx`。

## 先检查数据

下面命令均在项目根目录运行。首次使用建议先检查示例模板：

```powershell
python .\excel-import\excel_to_plaxis.py --excel ".\excel-import\钻孔数据模板.xlsx" --output ".\output\示例检查.p2dx" --validate-only
```

该操作读取 Excel 并生成预览和诊断文件，不启动 PLAXIS，也不生成 `.p2dx` 工程。检查资料保存在 `output/示例检查_import/`。

重新运行请使用新输出名称，或将旧检查结果移出原位置，避免同名目录冲突。

## 导入 PLAXIS

先保存并关闭正在运行的 PLAXIS Input。程序会启动自己的 Input 实例，建立新的工程。导入期间请勿手动修改模型；软件若出现登录或许可提示，请按自己的软件环境处理。

```powershell
python .\excel-import\excel_to_plaxis.py --excel ".\excel-import\钻孔数据模板.xlsx" --output ".\output\示例模型.p2dx" --plaxis "E:\plaxis2024 2d\Plaxis2DXInput.exe" --plaxis-python "C:\ProgramData\Seequent\PLAXIS Python Distribution V2\python\python.exe"
```

请修改：

- `--excel`：自己的 Excel 文件路径。
- `--output`：一个尚不存在的新工程路径。
- `--plaxis`：本机 `Plaxis2DXInput.exe` 的完整路径。
- `--plaxis-python`：与本机 PLAXIS 匹配的 Python 可执行文件路径。

路径带空格时保留双引号。不要把最后两个参数都设为 PLAXIS 主程序路径，它们分别指向软件和 Python。

也可以不传 `--excel`，通过窗口选择输入文件和新工程位置：

```powershell
python .\excel-import\excel_to_plaxis.py --plaxis "E:\plaxis2024 2d\Plaxis2DXInput.exe" --plaxis-python "C:\ProgramData\Seequent\PLAXIS Python Distribution V2\python\python.exe"
```

窗口模式需要普通 Python 具有 Tkinter。

## 命令参数

| 参数 | 作用 |
|---|---|
| `--excel` | 输入 `.xlsx`；不提供时显示文件选择窗口 |
| `--sheet` | 工作表名，默认为“钻孔数据” |
| `--output` | 新的 `.p2dx` 工程路径 |
| `--plaxis` | PLAXIS Input 程序路径 |
| `--plaxis-python` | 匹配 PLAXIS 的 Python 路径 |
| `--validate-only` | 仅检查数据、生成预览 |
| `--mesh` | 实际导入后尝试划分中等网格，不运行计算 |

## 输出与检查

实际导入成功后，应同时保留 `.p2dx` 和配套 `.p2dxdat` 数据目录。复制工程时请将两者一起复制。

工程旁的 `工程名_import` 目录保存数据快照、预览、检查结果和导入报告。原钻孔编号、层位名称与软件对象的对应关系也会记录。

完成后请在 PLAXIS 中检查剖面范围、层序和标高，再继续分配材料和设置水条件。

## 常见问题

| 问题 | 处理方法 |
|---|---|
| 找不到 `generate_strata` | 检查根目录是否有 `generate_strata.py`，并保留上述目录结构 |
| 找不到 Shapely 或 Matplotlib | 用实际启动导入器的同一套 Python 安装 `requirements.txt` |
| 找不到 PLAXIS 或官方 Python | 核对 `--plaxis` 和 `--plaxis-python`，不要照抄别人电脑的路径 |
| 提示 Input 已运行 | 先保存并关闭已有 Input，再导入 |
| 提示目标工程或记录目录已存在 | 更换 `--output` 文件名 |
| 提示层底顺序错误 | 检查是否把埋深当作标高，或层底是否高于层顶；相等的零厚度允许 |
| 公式、空白或文字引起错误 | 改为明确的数值；公式粘贴为值；缺层按零厚度填写 |
| 模型没有材料、水位或阶段 | 当前功能只导入地层几何，请另行设置 |

模型写入中途失败时，Input 可能留有部分未保存的几何；请检查错误信息和诊断文件，不要将其当作完整工程。再次导入前关闭该实例，并使用新输出名称。

## 反馈问题

提交 Issue 时请说明系统、PLAXIS 和 Python 版本、操作步骤、完整错误信息；如需附 Excel，请使用能复现问题的虚构或脱敏数据。不要上传连接密码或私人工程资料。

## 许可证与项目关系

正式开源发布前，请由发布者选择许可证，并将完整许可文本放在仓库根目录的 `LICENSE` 文件中。本 README 不替发布者指定许可证。

本项目是独立的辅助导入工具，与 PLAXIS 官方无隶属关系。PLAXIS 及第三方依赖遵循各自的许可。