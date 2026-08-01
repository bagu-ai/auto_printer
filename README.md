# 标签自动打印工具

把 Excel 样品清单拖进窗口，选好 BarTender 模板和「Excel 列 ↔ 模板字段」映射，一键自动打印全部标签。

## 使用流程（打印工作站）

1. 把 `标签自动打印工具.exe` 和 `模板` 文件夹放在同一目录（`模板` 里放 `.btw` 文件）
2. 双击运行（要求本机已装 **BarTender Automation / Enterprise 版**，并接好标签打印机）
3. 拖入 Excel 清单（.xls / .xlsx）
   - 自动识别表头行；多工作表文件可选「工作表」
   - 日期列自动格式化为 `YYYY-MM-DD`（不会打出 `00:00:00`）
   - 「送样数量」列如 `10g*3瓶` 自动解析为每行默认张数（3 张）
4. 选择模板（中文版 / 英文版）→ 后台自动读取模板里的命名数据源
5. 核对「字段映射」：每个模板字段取哪一列（或固定文本、留空）。同名自动匹配，映射按模板自动保存，下次直接用
6. 在左侧数据工作台选择一行，右侧「标签校样」会自动刷新；可用「适合窗口 / 100% / ±缩放 / 独立查看」按像素核对（BarTender 直接导出 300 DPI 单张标签，不消耗标签纸）
7. 数据工作台默认显示 10–14 行，可搜索任意列、拖动中间分隔条；运行记录默认收起，失败时自动展开
8. 勾选要打印的行、修改每行张数
8. 选打印机（默认用模板里保存的，选择会随映射一起记住）
9. 点「试打当前行」或「全部打印」——打印前自动**预检**（空值、重复编号、未映射字段）
10. 每次打印的任务记录保存在日志里，便于查错补打

## 数据存放位置

| 内容 | 位置 |
|---|---|
| 映射配置 | `%APPDATA%\标签自动打印工具\profiles\` |
| 打印任务日志 | `%APPDATA%\标签自动打印工具\logs\` |
| 预览图缓存 | `%APPDATA%\标签自动打印工具\previews\` |
| 模板（只读资源） | exe 同目录 `模板\` |

> 配置文件按「模板名+路径指纹」保存，不同目录的同名模板互不干扰；v1.0 的旧配置会自动迁移。

## 模板准备（一次性）

模板里需要动态填充的对象必须有「数据源名称」：

BarTender 中双击文本/条码对象 → 数据源 → **名称** 填一个有意义的名字（如 `发货新编号`、`客户香型`、`发货粒径`）→ 保存模板。

工具按这个名字列出可映射字段；没命名的对象不会出现在映射表里（保持模板原样打印）。

## 开发

```bash
pip install -r requirements.txt

# 单元测试（43 个：Excel 解析/格式化/多表、映射配置、预检、worker 线程模型、路径）
pytest

# 开发机无 BarTender：干跑模式演示全流程（打印动作只写日志）
python -m label_printer --mock

# 离屏自测（不弹窗：加载→映射→预览PNG→线程打印→任务日志）
python tools/smoke_mock.py

# 连接真 BarTender 运行
python -m label_printer

# 打包 exe（自动重建发布 zip）
build_exe.bat
```

## 架构要点

- **COM 线程亲和**：BarTender COM 对象只在 QThread 内创建、使用、销毁（`pythoncom.CoInitialize` 配对），绝不跨线程共享——真机打印稳定性的关键
- **后端抽象**：`PrintBackend` 接口 + BarTender / mock 双实现，UI 只面向接口；开发机无 BarTender 也能演示全流程
- **标签校样**：BarTender `ExportImage` 直接导出 300 DPI 单张标签设计面（60×30mm 为约 708×354 px），避免整页预览缩小发糊；保留源图并支持 fit/100%/25%–400% 缩放与独立查看
- **打印结果校验**：`Format.Print` 等到任务提交 spooler 并检查 `BtPrintResult` 与 Messages，失败行带 BarTender 原始错误

## 目录结构

```
label_printer/
  __main__.py            入口（--mock 干跑模式）
  core/
    excel_reader.py      Excel 加载、表头探测、单元格显示值格式化、张数解析
    profiles.py          映射配置（模板名+路径哈希）、预检 validate_jobs
    paths.py             resource_dir（只读资源）/ state_dir（%APPDATA%）
    backend.py           打印后端抽象接口
    bartender_backend.py BarTender COM 实现（目标机）
    mock_backend.py      干跑实现（开发机调试）
  ui/main_window.py      PySide6 主界面（三个 QThread worker：扫描/预览/打印）
tests/                   pytest 单测（真实 xls 样例存在才跑，核心用例全部临时生成）
tools/smoke_mock.py      离屏全流程自测
.github/workflows/ci.yml Windows CI（pytest + 冒烟）
```
