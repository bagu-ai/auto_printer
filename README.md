# 标签自动打印工具

将 Excel 清单映射到 BarTender 模板，提供标签校样、打印预检和批量任务提交。真实打印需要 Windows、BarTender Automation / Enterprise、正确的模板和可用的打印机；模拟模式不连接打印机。

## v1.3.3 的行为

- 默认使用提交模式：每行完整设置字段和份数，调用打印接口后检查返回结果，不等待设备完成反馈。直接打开 EXE 与 `启动标签打印.cmd` 使用相同默认模式。
- BarTender COM 会话在独立子进程内串行运行；父进程提供取消与超时监护。发生错误、断链或超时后停止后续提交。
- 提交前记录不可变任务快照，逐行写入 JSONL 增量日志；结束后保存 JSON 汇总。无法确定是否已经提交的行标为 `unknown`，不会自动重发。
- 扫描、校样与打印互斥；清理结束前不启动新任务。关闭 BarTender 使用不保存更改的参数，关闭异常会显示出来。

`submitted` 仅表示打印接口报告任务已提交，不代表纸张已经输出，也不证明标签内容、份数全部正确。取消不能撤回已经提交到打印队列的任务；异常后必须结合队列和实物确认实际缺少的标签。

已有现场记录来自一台 Windows / BarTender 2016 R1 / USB001 打印工作站：提交模式四批合计 160 行、258 份均记录为 `submitted`，使用者反馈该模式可以完成任务。该记录不等于逐张实物验收，也不能替代其他模板、打印机、驱动或工作站的验证。

## 使用发布文件夹

1. 将新发布文件夹完整放在 D 盘可写目录，保留 EXE、`_internal`、`模板` 和启动脚本在一起，不覆盖旧版本目录。
2. 双击 `标签自动打印工具_防卡住版.exe` 或 `启动标签打印.cmd`。
3. 导入 Excel（`.xls` / `.xlsx`），选择工作表，核对识别的表头、日期和每行份数。
4. 选择本站点已验证的 `.btw` 模板，核对“Excel 列 ↔ 模板字段”映射及实际打印机队列。
5. 先选择一行试打，核对编号、内容、位置和份数，再提交所选批次。预检提示不能代替实物核对。
6. 出现错误或 `unknown` 时先保留日志并核对队列、实物，不直接整批重打。

完整操作说明见 [使用说明.txt](使用说明.txt)；版本切换和异常处理见 [修复版使用说明.txt](修复版使用说明.txt)。

模板中需动态填充的对象应在 BarTender 的“数据源 → 名称”设置名称，如 `发货新编号`、`客户香型`、`发货粒径`。未命名对象保持模板原样，不参与映射。应使用与现场 BarTender 版本兼容的模板。

## 数据与版本切换

默认状态目录是 `程序目录\运行数据\`；源码运行时，“程序目录”指仓库根目录。

| 内容 | 默认位置 |
|---|---|
| 字段映射 | `运行数据\profiles\` |
| 任务增量日志与汇总 | `运行数据\logs\` |
| 预览缓存 | `运行数据\previews\` |
| 窗口配置与待核对标记 | `运行数据\window.ini` |
| 程序临时文件与 COM 缓存 | `运行数据\.tmp\`、`运行数据\.cache\` |
| 模板 | 程序目录下 `模板\` |

可用 `LABEL_PRINTER_HOME` 指定另一个可写状态目录。新目录不会自动继承旧版本的配置或待核对标记；更换版本前先核对未完成批次。需要保留映射时，可复制旧版 `profiles` 到新状态目录后重新核对模板路径、字段和打印机。模板路径变化可能需要重新选择映射。

## 从源码开发与验证

源码开发可复用已准备好的 Windows / Python 3.12.10 环境，也可在 D 盘仓库内建立 `.venv`。以下命令在仓库根目录执行，`python` 应指向准备使用的解释器；`requirements-build.txt` 固定开发、测试和构建所需依赖，包括 PySide6、pywin32、pytest 和 PyInstaller。复用已有环境时，将命令中的 `.\.venv\Scripts\python.exe` 替换为该环境解释器的绝对路径。

```powershell
New-Item -ItemType Directory -Force .tmp, .cache | Out-Null
$env:TEMP = (Resolve-Path .tmp).Path
$env:TMP = $env:TEMP
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install --cache-dir .cache\pip -r requirements-build.txt

# 单元测试；缺少本地真实样例时，相关可选用例会跳过
.\.venv\Scripts\python.exe -m pytest -q --basetemp .tmp\pytest

# 故障恢复及界面/Excel 模拟自检，不连接真实 BarTender 或打印机
.\.venv\Scripts\python.exe run_app.py --self-test .tmp\self-test
.\.venv\Scripts\python.exe run_app.py --ui-self-test .tmp\ui-self-test

# 打开模拟界面
.\.venv\Scripts\python.exe -m label_printer --mock
```

自检使用生成的模拟数据，无需客户 Excel 清单。界面自检需要 Windows 中文字体，缺少字体时应先补齐该运行条件。需要实际打印时，在已配置 BarTender 的工作站使用 `.\.venv\Scripts\python.exe -m label_printer`，并先试打核对。自动化测试和模拟自检不验证真实 COM 环境、驱动、纸张输出或设备兼容性；本地通过也不代表远程 CI 已执行。

现有 GitHub 工作流配置 Windows / Python 3.13，并继续调用 `tools/smoke_mock.py`；该脚本现在转调用 `--ui-self-test` 的模拟验证流程。本次未修改工作流，Python 3.13 的远程 CI 结果以合并请求中的实际检查为准，不能从 Python 3.12.10 本地结果推断。

## 构建发布文件夹

先完成上面的依赖安装，再在仓库根目录执行：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tools\build_release.ps1
```

默认使用仓库内 `.venv\Scripts\python.exe`，也可通过 `-Python` 传入已安装依赖的解释器绝对路径。`build_exe.bat` 调用同一构建流程。

输出为 `dist\v1.3.3\标签自动打印工具_防卡住版\`，包含 EXE、`_internal`、模板、修复版使用说明和唯一的启动脚本。构建工作文件、缓存和临时文件分别位于仓库内 `build`、`.cache`、`.tmp`。构建不会关闭正在运行的打印程序或清空打印队列。

当前流程生成可携文件夹，不自动生成 ZIP，也不依赖客户 Word/PDF 手册。需要压缩发布时，将整个文件夹作为一个包保存，不能只分享 EXE。本仓库构建不包含额外的现场诊断工具包。

## 开发排查选项

| 环境变量 | 行为 |
|---|---|
| `LABEL_PRINTER_WAIT_FOR_COMPLETION` | 默认 `0`：提交模式；`1`：等待完成反馈，仅供排查。 |
| `LABEL_PRINTER_PRINT_TIMEOUT_MS` | 默认 `60000`；外围监护再增加 15 秒。提交模式的阻塞保护由外围监护保证，不假定内层超时生效。 |
| `LABEL_PRINTER_HOME` | 覆盖运行数据目录，使用 D 盘可写目录。 |
| `LABEL_PRINTER_FAST_MODE` | 提交模式下不启用重复值省略；等待模式的可选优化仍须验证模板脚本兼容性。 |

## 代码结构

- `label_printer/core/excel_reader.py`：Excel 加载、表头识别、值和份数解析。
- `label_printer/core/profiles.py`：字段映射与打印预检。
- `label_printer/core/bartender_backend.py`：单进程内的 BarTender COM 操作与结果检查。
- `label_printer/core/isolated_backend.py`：独立进程通信、取消、超时与有限清理。
- `label_printer/core/job_journal.py`：任务快照和增量日志。
- `label_printer/core/paths.py`、`runtime.py`：状态目录、缓存与临时文件。
- `label_printer/ui/main_window.py`：界面、扫描/校样/打印调度和异常后的核对提示。
- `tests/`：单元与模拟行为测试；`--self-test`、`--ui-self-test` 提供独立自检入口。

本次变更不新增业务清单、真实运行日志或凭据。原仓库已跟踪的样例 XLS、`docs/` 和 `.research/` 保持原样；这不表示整个仓库不含样例或文档。具体提交范围、验证边界与回退说明见 [DELIVERY.md](DELIVERY.md)。
