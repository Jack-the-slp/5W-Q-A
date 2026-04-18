"""
5W提问法 Word图卡生成器
- 从网上抓取场景图片
- 生成可打印的Word文档（图+5W问答）
- 保存到桌面
"""
import os, json, ssl, urllib.request, urllib.error
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import time

# ============ 路径配置 ============
BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
SCENESFile = os.path.join(BASE_DIR, "5w_scenes.json")
IMG_DIR    = os.path.join(BASE_DIR, "5w_images")
DESKTOP    = "/mnt/c/Users/123/Desktop"

os.makedirs(IMG_DIR, exist_ok=True)

# ============ 图片获取（使用本地已下载图片） ============
def get_local_image(scene_id):
    """返回本地图片路径，如果不存在返回None"""
    cache = os.path.join(IMG_DIR, f"{scene_id}.jpg")
    if os.path.exists(cache):
        return cache
    # 尝试png格式
    cache_png = os.path.join(IMG_DIR, f"{scene_id}.png")
    if os.path.exists(cache_png):
        return cache_png
    return None

# ============ Word 样式工具 ============
def set_cell_bg(cell, hex_color):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_color)
    tcPr.append(shd)

def add_run(para, text, bold=False, color=None, size=10, font="微软雅黑"):
    run = para.add_run(text)
    run.font.name  = font
    run.font.size  = Pt(size)
    run.font.bold  = bold
    if color:
        run.font.color.rgb = RGBColor(*bytes.fromhex(color))
    # 设置中文字体
    r = run._r
    rPr = r.get_or_add_rPr()
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:eastAsia"), font)
    rPr.insert(0, rFonts)
    return run

def make_table_row(table, w_values, bg=None, text_size=9):
    row = table.add_row()
    for i, (cell, val) in enumerate(zip(row.cells, w_values)):
        cell.text = ""
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after  = Pt(1)
        p.paragraph_format.left_indent  = Pt(3)
        add_run(p, val, size=text_size)
        if bg:
            set_cell_bg(cell, bg)
    return row

# ============ 生成单个场景Word ============
def generate_scene_word(scene, img_path, output_dir):
    doc  = Document()
    page = doc.sections[0]
    page.page_width  = Cm(21)
    page.page_height = Cm(29.7)
    page.left_margin   = Cm(1.8)
    page.right_margin  = Cm(1.8)
    page.top_margin    = Cm(1.5)
    page.bottom_margin = Cm(1.5)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.space_before = Pt(0)
    title.paragraph_format.space_after  = Pt(4)
    add_run(title, f"【{scene['name']}】5W提问练习", bold=True, color="1A5F8A", size=14)

    # 图片区域（居中）
    if img_path and os.path.exists(img_path):
        pic_para = doc.add_paragraph()
        pic_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = pic_para.add_run()
        run.add_picture(img_path, width=Inches(5.5))
        pic_para.paragraph_format.space_after = Pt(6)

    # 5W 问题表格
    questions = scene.get("questions", {})
    W_COLORS = {
        "谁 Who":          ("谁", "D6EAF8"),
        "什么 What":       ("什么", "D5F5E3"),
        "哪里 Where":      ("哪里", "FCF3CF"),
        "什么时候 When":    ("什么时候", "FADBD8"),
        "为什么 Why":       ("为什么", "E8DAEF"),
        "做什么 What doing":("做什么", "FDEBD0"),
    }

    table = doc.add_table(rows=0, cols=2)
    table.style = "Table Grid"
    table.autofit = False

    col_widths = [Cm(3.5), Cm(13.5)]
    for row in table.columns[0].cells:
        row.width = col_widths[0]
    for row in table.columns[1].cells:
        row.width = col_widths[1]

    for w_key, qs in questions.items():
        if not qs:
            continue
        label, bg = W_COLORS.get(w_key, (w_key.split()[0], "F0F0F0"))

        # 标签行（跨两列）
        lbl_row = table.add_row()
        lbl_cell = lbl_row.cells[0].merge(lbl_row.cells[1])
        lbl_cell.text = ""
        lp = lbl_cell.paragraphs[0]
        lp.paragraph_format.space_before = Pt(2)
        lp.paragraph_format.space_after  = Pt(2)
        lp.paragraph_format.left_indent  = Pt(6)
        lp.alignment = WD_ALIGN_PARAGRAPH.LEFT
        add_run(lp, f"◆ {label}", bold=True, color="1A5F8A", size=10)
        set_cell_bg(lbl_cell, bg)

        # 问题行
        for q in qs:
            row = table.add_row()
            row.cells[0].text = ""
            row.cells[1].text = ""
            # 左单元格：方框（勾选区）
            p0 = row.cells[0].paragraphs[0]
            p0.paragraph_format.space_before = Pt(1)
            p0.paragraph_format.space_after  = Pt(1)
            p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_run(p0, "□", size=11)
            # 右单元格：问题
            p1 = row.cells[1].paragraphs[0]
            p1.paragraph_format.space_before = Pt(1)
            p1.paragraph_format.space_after  = Pt(1)
            p1.paragraph_format.left_indent  = Pt(4)
            add_run(p1, q, size=9)

    # 页脚
    footer  = doc.sections[0].footer
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run(fp, "上海小螺号能力开发中心  ·  5W语言训练图卡", color="888888", size=8)

    filename = f"5W_{scene['name']}.docx"
    out_path = os.path.join(output_dir, filename)
    doc.save(out_path)
    return out_path

# ============ 主函数 ============
def main():
    with open(SCENESFile, "r", encoding="utf-8") as f:
        data = json.load(f)
    scenes = data["scenes"]

    print(f"找到 {len(scenes)} 个场景，开始生成...")
    print(f"输出目录: {DESKTOP}\n")

    for scene in scenes:
        print(f"处理: {scene['name']}")
        img_path = get_local_image(scene["id"])
        out_path  = generate_scene_word(scene, img_path, DESKTOP)
        print(f"  → {os.path.basename(out_path)}")

    print(f"\n全部完成！文件已保存到桌面。")

if __name__ == "__main__":
    main()
