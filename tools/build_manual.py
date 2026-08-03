"""生成客户版《标签自动打印工具 使用说明书》Word + PDF。

用法：python tools/build_manual.py
输出：docs/标签自动打印工具_使用说明书_v1.2.1.docx / .pdf
依赖：python-docx + 本机 Word（COM 导出 PDF）
"""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
IMG = ROOT / "docs" / "images"

from docx import Document  # noqa: E402
from docx.enum.table import WD_TABLE_ALIGNMENT  # noqa: E402
from docx.enum.text import WD_ALIGN_PARAGRAPH  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402
from docx.shared import Cm, Pt, RGBColor  # noqa: E402

ACCENT = RGBColor(0x08, 0x7E, 0xA4)
INK = RGBColor(0x17, 0x21, 0x2B)
GRAY = RGBColor(0x5E, 0x6B, 0x78)


def set_cn_font(style_or_run, name_cn="微软雅黑", name_en="Segoe UI"):
    """同时设置中英文/东亚字体，避免 Word 里中文回退到宋体。"""
    font = style_or_run.font
    font.name = name_en
    rpr = style_or_run.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        from docx.oxml import OxmlElement
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:ascii"), name_en)
    rfonts.set(qn("w:hAnsi"), name_en)
    rfonts.set(qn("w:eastAsia"), name_cn)


def setup_styles(doc):
    normal = doc.styles["Normal"]
    normal.font.size = Pt(11)
    normal.font.color.rgb = INK
    set_cn_font(normal)
    normal.paragraph_format.line_spacing = 1.4
    normal.paragraph_format.space_after = Pt(6)

    for name, size, color in (("Heading 1", 17, ACCENT), ("Heading 2", 14, INK),
                              ("Heading 3", 12, INK)):
        st = doc.styles[name]
        st.font.size = Pt(size)
        st.font.bold = True
        st.font.color.rgb = color
        set_cn_font(st)
        st.paragraph_format.space_before = Pt(16 if name == "Heading 1" else 10)
        st.paragraph_format.space_after = Pt(6)


def p(doc, text, bold=False, size=None, color=None, align=None, style=None):
    par = doc.add_paragraph(style=style)
    run = par.add_run(text)
    run.bold = bold
    if size:
        run.font.size = Pt(size)
    if color:
        run.font.color.rgb = color
    if align is not None:
        par.alignment = align
    return par


def bullet(doc, text, bold_prefix=""):
    par = doc.add_paragraph(style="List Bullet")
    if bold_prefix:
        r = par.add_run(bold_prefix)
        r.bold = True
    par.add_run(text)
    return par


class StepCounter:
    """每章重新从 1 开始的手工编号（List Number 样式会全文连号，故不用）。"""

    def __init__(self):
        self.n = 0

    def reset(self):
        self.n = 0


_steps = StepCounter()


def step(doc, text):
    _steps.n += 1
    par = doc.add_paragraph()
    par.paragraph_format.left_indent = Cm(0.75)
    par.paragraph_format.first_line_indent = Cm(-0.75)
    num = par.add_run(f"{_steps.n}. ")
    num.bold = True
    num.font.color.rgb = ACCENT
    par.add_run(text)
    return par


def picture(doc, name, caption, width_cm=15.5):
    path = IMG / name
    if not path.exists():
        print(f"[warn] 缺少插图 {name}，跳过")
        return
    par = doc.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    par.add_run().add_picture(str(path), width=Cm(width_cm))
    cap = p(doc, caption, size=9, color=GRAY, align=WD_ALIGN_PARAGRAPH.CENTER)
    cap.paragraph_format.space_after = Pt(12)


def table(doc, header, rows, widths=None):
    t = doc.add_table(rows=1 + len(rows), cols=len(header))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for j, h in enumerate(header):
        cell = t.rows[0].cells[j]
        cell.text = ""
        run = cell.paragraphs[0].add_run(h)
        run.bold = True
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            t.rows[i + 1].cells[j].text = str(val)
    if widths:
        for j, w in enumerate(widths):
            for r in t.rows:
                r.cells[j].width = Cm(w)
    doc.add_paragraph()
    return t


def build():
    doc = Document()
    setup_styles(doc)
    for section in doc.sections:
        section.page_width, section.page_height = Cm(21), Cm(29.7)
        section.top_margin = section.bottom_margin = Cm(2.2)
        section.left_margin = section.right_margin = Cm(2.5)

    # ---------- 封面 ----------
    for _ in range(5):
        doc.add_paragraph()
    p(doc, "标签自动打印工具", bold=True, size=30, color=ACCENT,
      align=WD_ALIGN_PARAGRAPH.CENTER)
    p(doc, "使 用 说 明 书", size=20, align=WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_paragraph()
    p(doc, "（Excel 清单 → 条码标签 批量打印）", size=12, color=GRAY,
      align=WD_ALIGN_PARAGRAPH.CENTER)
    for _ in range(8):
        doc.add_paragraph()
    p(doc, "版本：V1.2.1", size=12, align=WD_ALIGN_PARAGRAPH.CENTER)
    p(doc, "适用模板：60×30mm 中文版 / 英文版标签", size=12,
      align=WD_ALIGN_PARAGRAPH.CENTER)
    p(doc, "2026 年 8 月", size=12, color=GRAY, align=WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_page_break()

    # ---------- 1 软件简介 ----------
    doc.add_heading("一、这个软件是做什么的", 1)
    p(doc, "以前打印样品标签，需要在 BarTender 里对照 Excel 清单，一行一行地"
           "手工抄写编号、香型、粒径，再逐行点打印。行数多时既慢又容易抄错。")
    p(doc, "这个工具把整个过程变成：把 Excel 清单拖进窗口 → 核对内容 → 一键全部打印。"
           "每一行打印什么、打几张，都由清单自动带出，打印前还能在屏幕上看到"
           "和实际标签一模一样的“校样”图。")
    table(doc,
          ["对比项", "手工方式", "本工具"],
          [["35 行清单录入", "逐行手抄约 30 分钟", "拖入文件，几秒钟"],
           ["每行打印张数", "人工记忆输入", "按“送样数量”自动带出（如 10g*3瓶 → 3 张）"],
           ["内容校对", "打完才发现错", "打印前屏幕校样 + 自动预检"],
           ["打印记录", "无", "每次打印自动存日志，可查可补打"]],
          widths=[4.5, 5.5, 6.5])

    # ---------- 2 运行环境 ----------
    doc.add_heading("二、使用前需要准备什么", 1)
    table(doc,
          ["项目", "要求"],
          [["电脑系统", "Windows 10 / 11（打印工作站电脑）"],
           ["BarTender", "已安装 BarTender 2016 Automation / Enterprise 版（打印工作站已装）"],
           ["打印机", "东芝 TOSHIBA B-EX4T1 或 ARGOX CP-2140M，驱动已装好、能正常打印"],
           ["Excel 清单", ".xls 或 .xlsx 格式；第一行附近要有表头（如 序号、发货新编号、客户香型…）"],
           ["标签模板", ".btw 文件，且需要自动填充的对象已起好“数据源名称”（见第八章）"]],
          widths=[4, 12])
    p(doc, "注意：本工具不需要安装 Python 或任何其他软件，双击 exe 即可使用。", bold=True)

    # ---------- 3 安装 ----------
    doc.add_heading("三、安装（只需一次）", 1)
    _steps.reset()
    step(doc, "把收到的压缩包「标签自动打印工具_v1.2.1.zip」解压到一个固定位置，"
              "例如 D:\\标签打印\\（不要放在桌面临时文件夹里，以后别移动它）。")
    step(doc, "解压后里面有：标签自动打印工具.exe、「模板」文件夹（内含中文版/英文版 .btw 模板）、使用说明.txt。")
    step(doc, "在 exe 上点右键 →「发送到 → 桌面快捷方式」，以后从桌面双击启动。")
    step(doc, "第一次启动需要解压资源，约等 10–20 秒，属于正常现象；之后启动很快。")
    p(doc, "如果杀毒软件提示风险，请选择“信任 / 加入白名单”。"
           "（程序为打包的绿色软件，部分杀软会误报。）")

    # ---------- 4 认识界面 ----------
    doc.add_heading("四、认识界面", 1)
    picture(doc, "main_window.png", "图 1　主界面（已载入一份 35 行清单）")
    bullet(doc, "：显示后端连接状态和「数据行数 / 待打印张数」总览。", "① 顶部状态条")
    bullet(doc, "：选择 Excel 清单、工作表、表头行和标签模板。", "② 输入区")
    bullet(doc, "：清单的每一行都在这里，勾选哪些行打印、每行打几张；上方可搜索。", "③ 数据工作台（左）")
    bullet(doc, "：右侧两个标签页——「标签校样」看当前行的标签效果图；「字段映射」配置模板字段取 Excel 哪一列。", "④ 校样 / 映射（右）")
    bullet(doc, "：选择打印机、试打当前行、一键打印全部。主按钮上会实时显示「打印 N 行 · N 张」。", "⑤ 打印操作条（底部）")
    bullet(doc, "：默认收起，点击标题可展开；出现错误时会自动展开。", "⑥ 运行记录")
    p(doc, "界面布局可以调整：拖动中间的分隔条改变左右宽度；窗口大小、分隔位置、"
           "标签页选择都会被记住，下次打开自动恢复。")

    # ---------- 5 标准流程 ----------
    doc.add_heading("五、标准打印流程（照做即可）", 1)
    _steps.reset()
    step(doc, "双击启动软件，确认右上角状态显示「BarTender 已就绪」。")
    step(doc, "把 Excel 清单文件直接拖进窗口（或点「选择 Excel」）。"
              "软件会自动识别表头行和数据；如果识别错了，改「表头第 N 行」后点「重新读取」。")
    step(doc, "文件里有多个工作表时，在「工作表」下拉框切换到正确的那一张。")
    step(doc, "在「标签模板」下拉框选择中文版或英文版模板，稍等一两秒让软件读出模板字段。")
    step(doc, "第一次使用某个模板时，点右侧「字段映射」标签页核对：每个模板字段取 Excel 的哪一列。"
              "同名的会自动配好；也可以选「固定文本」手动输入固定内容，或「留空」。"
              "映射会自动保存，下次用同一模板不用再配。")
    step(doc, "需要在每行内容前后加固定文字时（这些文字 Excel 里没有），用「前缀」「后缀」两列。"
              "例如给「编号」前缀填 编号：、「香型」前缀填 香型：、「粒径」前缀填 粒径：，"
              "打印出来就是“编号：TB-71717　香型：红酒　粒径：3.0mm”这样的效果。"
              "前缀/后缀同样随映射自动保存，改完右侧校样会立即刷新。")
    picture(doc, "mapping_panel.png",
            "图 2　字段映射：自动匹配 Excel 列，并用「前缀」给每行加固定文字", width_cm=8.5)
    step(doc, "回到「标签校样」标签页，点数据表里的任意一行，右侧会显示这一行打印出来的样子"
              "（由 BarTender 按 300 DPI 真实渲染，和印出来一致）。核对方法见第六章。")
    step(doc, "在数据工作台勾选要打印的行；需要改某行张数时直接改「张数」列的数字"
              "（默认按“送样数量”自动解析，如 10g*3瓶 → 3 张）。")
    step(doc, "选择打印机。一般保持「使用模板默认打印机」即可；你的选择会随映射一起记住。")
    step(doc, "先点「试打当前行」，打出一张实物核对：内容、位置、清晰度都对，再点蓝色主按钮"
              "「打印 N 行 · N 张」批量打印。")
    step(doc, "批量打印前会弹出确认框，显示将发送的行数和总张数，确认无误点「是」。")
    p(doc, "打印过程中可以点「取消」中断；每打印完一行，运行记录里都有对应日志。", bold=False)

    # ---------- 6 校样 ----------
    doc.add_heading("六、标签校样怎么用（打印前核对）", 1)
    picture(doc, "preview_panel.png", "图 3　标签校样区", width_cm=8.5)
    bullet(doc, "：完整显示标签，且不会把小图强行放大（保持清晰）。", "适合窗口")
    bullet(doc, "：按原始像素 1:1 显示，用来核对文字边缘、条码细节。", "100%")
    bullet(doc, "：在 25%–400% 之间缩放；也可以按住 Ctrl 滚鼠标滚轮。", "− / +")
    bullet(doc, "：放大后按住鼠标左键拖动，查看标签四周边缘。", "平移")
    bullet(doc, "：双击校样图在「适合窗口」和「100%」之间切换。", "双击")
    bullet(doc, "：弹出一个大窗口专门核对标签，适合精细检查。", "独立查看")
    p(doc, "校样图下方的信息栏显示源图像素尺寸和当前缩放比例。"
           "修改字段映射、切换打印机、选择其他行时，校样会自动刷新。")

    # ---------- 7 预检 ----------
    doc.add_heading("七、打印前的自动检查（预检）", 1)
    p(doc, "点打印或试打时，软件会先自动检查，发现问题会提示：")
    table(doc,
          ["检查项", "说明", "处理方式"],
          [["空值提醒", "某列在勾选的行里有空内容，打出来对应位置是空白", "确认可以接受就继续；否则改 Excel 或调整映射"],
           ["重复编号", "勾选的行里出现相同编号", "回 Excel 核实是否录重"],
           ["未映射字段", "模板字段没有配取值来源", "到「字段映射」标签页补配，或确认留空"]],
          widths=[3.5, 7.5, 5.5])
    p(doc, "红色错误必须处理才能打印；黄色警告确认后可以继续。")

    # ---------- 8 模板准备 ----------
    doc.add_heading("八、模板准备（一次性，由管理员完成）", 1)
    _steps.reset()
    p(doc, "模板里需要自动填充内容的文本或条码，必须先在 BarTender 里起名字，"
           "软件才知道往哪里填：")
    step(doc, "用 BarTender 打开 .btw 模板。")
    step(doc, "双击要自动填充的对象（如编号文本）→ 左侧选「数据源」→ 在「名称」一栏填上有意义的名字，"
              "例如：编号、香型、粒径（或 发货新编号、客户香型、发货粒径）。")
    step(doc, "每个需要动态内容的对象都起好名字后，保存模板。")
    p(doc, "没起名字的对象不会出现在映射表里，按模板原样打印（适合固定不变的内容）。",
      color=GRAY)

    # ---------- 9 数据位置 ----------
    doc.add_heading("九、数据保存在哪里", 1)
    table(doc,
          ["内容", "位置"],
          [["字段映射配置", "C:\\Users\\<用户名>\\AppData\\Roaming\\标签自动打印工具\\profiles\\"],
           ["打印任务日志", "…\\标签自动打印工具\\logs\\（每次打印一份 .json，记录每行内容与成败）"],
           ["校样缓存", "…\\标签自动打印工具\\previews\\（只保留最新一张，不占空间）"],
           ["标签模板", "exe 同目录的「模板」文件夹"]],
          widths=[4.5, 11.5])
    p(doc, "重装软件或移动 exe 都不会丢配置和日志；映射按「模板名 + 路径」分别保存，"
           "不同目录的同名模板互不干扰。")

    # ---------- 10 FAQ ----------
    doc.add_heading("十、常见问题", 1)
    faq = [
        ("双击 exe 没反应 / 提示 BarTender 不可用",
         "确认本机 BarTender（Automation/Enterprise 版）能正常打开。弹窗里也可以先选"
         "「干跑模式」演示流程（不会真正打印），用于培训或检查配置。"),
        ("字段映射表是空的",
         "模板里没有命名数据源，请按第八章给对象起名字；或者模板还在读取中，稍等一两秒。"),
        ("校样显示“生成失败”",
         "看运行记录里的具体原因。常见是模板路径被移动、或 BarTender 正被其他窗口占用（"
         "先把 BarTender 设计器里打开的同一模板关掉）。如果提示「为第一个参数指定的"
         "颜色不正确」，说明用的是 V1.2 旧版软件，请换用 V1.2.1 及以上版本（已兼容 "
         "BarTender 2016）。"),
        ("加前缀后文字超出标签边界 / 被裁掉",
         "模板文本框在当前字号下放不下了。在 BarTender 里双击该文本对象 → 左侧「文本格式」"
         "→「自动调整大小」选「缩小文字以适应对象」，保存模板；之后内容过长会自动缩小，"
         "不会再被裁掉。（改模板请在打印工作站的 BarTender 2016 上做，别用更高版本保存。）"),
        ("打印出来的内容不对",
         "先在「标签校样」用 100% 或独立查看核对，再「试打当前行」；不对就调整字段映射"
         "（映射自动保存）。"),
        ("预检提示“某列有 N 行为空”",
         "这些行在该列没有内容，打印出来对应位置空白。确认无误可继续；有问题回去改 Excel。"),
        ("打印中断了 / 想查打了哪些",
         "每次打印都有任务日志：%APPDATA%\\标签自动打印工具\\logs\\ 里按时间命名的 "
         ".json 文件，记录每一行的内容和成功/失败。"),
        ("窗口显示不正常 / 想恢复默认布局",
         "直接关闭软件重开即可；布局信息保存在本机，删除 "
         "HKEY_CURRENT_USER\\Software\\BG0\\LabelPrinter 可恢复默认。"),
        ("杀毒软件报毒",
         "打包软件偶有误报，加白名单即可。"),
    ]
    for q, a in faq:
        p(doc, "Q：" + q, bold=True)
        par = p(doc, "A：" + a)
        par.paragraph_format.space_after = Pt(10)

    # ---------- 11 注意事项 ----------
    doc.add_heading("十一、打印现场注意事项", 1)
    bullet(doc, "批量打印前，务必先「试打当前行」确认实物效果；换模板、换纸、换碳带后都要重新试打一张。")
    bullet(doc, "确认打印机联机、标签纸和碳带充足，再点批量打印。")
    bullet(doc, "打印进行中不要关闭软件；如需停止请点「取消」。")
    bullet(doc, "Excel 清单修改后，点「重新读取」让软件载入最新内容。")
    bullet(doc, "日期列在 Excel 里是什么格式，打印出来就是什么格式；软件会自动把日期显示为 2026-01-03 这样的形式。")

    doc.add_paragraph()
    p(doc, "如遇到本说明书未覆盖的问题，请截图（含运行记录）并联系软件提供方。",
      color=GRAY)

    out_docx = ROOT / "docs" / "标签自动打印工具_使用说明书_v1.2.1.docx"
    doc.save(out_docx)
    print("docx:", out_docx)
    return out_docx


def to_pdf(docx_path: Path):
    import win32com.client
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(str(docx_path), ReadOnly=True)
        pdf = docx_path.with_suffix(".pdf")
        doc.SaveAs2(str(pdf), FileFormat=17)  # wdFormatPDF
        doc.Close(False)
        print("pdf:", pdf)
    finally:
        word.Quit(False)


if __name__ == "__main__":
    path = build()
    to_pdf(path)
