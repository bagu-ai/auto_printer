# 标签自动打印工具

把 Excel 样品清单拖进窗口，选好 BarTender 模板和「Excel 列 ↔ 模板字段」映射，一键自动打印全部标签。

## 使用流程（打印工作站）

1. 把 `标签自动打印工具.exe` 和 `模板` 文件夹放在同一目录（`模板` 里放 `.btw` 文件）
2. 双击运行（要求本机已装 **BarTender Automation / Enterprise 版**，并接好标签打印机）
3. 拖入 Excel 清单（.xls / .xlsx）→ 自动识别表头行与每行默认张数（按「送样数量」如 `10g*3瓶` → 3 张）
4. 选择模板（中文版 / 英文版）→ 工具自动列出模板里的命名数据源
5. 核对「字段映射」：每个模板字段取哪一列（或固定文本、留空）。同名自动匹配，映射会按模板自动保存，下次直接用
6. 在数据预览里勾选要打印的行、修改每行张数
7. 选打印机（默认用模板里保存的），先 **试打第一张** 核对，再 **全部打印**

## 模板准备（一次性）

模板里需要动态填充的对象必须有「数据源名称」：

BarTender 中双击文本/条码对象 → 数据源 → **名称** 填一个有意义的名字（如 `发货新编号`、`客户香型`、`发货粒径`）→ 保存模板。

工具按这个名字列出可映射字段；没命名的对象不会出现在映射表里（保持模板原样打印）。

## 开发

```bash
pip install -r requirements.txt

# 单元测试（Excel 解析、张数解析、映射配置）
pytest

# 开发机无 BarTender：干跑模式演示全流程（打印动作只写日志）
python -m label_printer --mock

# 离屏自测（不弹窗，跑通 加载→映射→模拟打印）
python tools/smoke_mock.py

# 连接真 BarTender 运行
python -m label_printer

# 打包 exe
build_exe.bat
```

## 目录结构

```
label_printer/
  __main__.py            入口（--mock 干跑模式）
  core/
    excel_reader.py      Excel 加载、表头自动探测、张数解析
    profiles.py          映射配置 JSON 读写（按模板保存）
    backend.py           打印后端抽象接口
    bartender_backend.py BarTender COM 实现（目标机）
    mock_backend.py      干跑实现（开发机调试）
  ui/main_window.py      PySide6 主界面
tests/                   pytest 单测
tools/smoke_mock.py      离屏全流程自测
profiles/                运行时生成的映射配置（exe 同目录）
```
