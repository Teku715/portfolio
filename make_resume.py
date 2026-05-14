from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

doc = Document()

# 设置默认字体
style = doc.styles['Normal']
style.font.name = '微软雅黑'
style.font.size = Pt(11)
style._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')

# 页边距
for section in doc.sections:
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(2)
    section.right_margin = Cm(2)

def add_title(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(text)
    run.font.size = Pt(13)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0x1a, 0x1a, 0x2e)
    run.font.name = '微软雅黑'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    pPr = p._element.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), '2563EB')
    pBdr.append(bottom)
    pPr.append(pBdr)

def add_item(doc, title, meta='', content=''):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(1)
    if meta:
        run1 = p.add_run(title)
        run1.font.size = Pt(11)
        run1.font.bold = True
        run1.font.name = '微软雅黑'
        run1._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
        run2 = p.add_run(f'  {meta}')
        run2.font.size = Pt(10)
        run2.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
    else:
        run = p.add_run(title)
        run.font.size = Pt(11)
        run.font.bold = True
        run.font.name = '微软雅黑'
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    if content:
        p2 = doc.add_paragraph(content)
        p2.paragraph_format.space_before = Pt(1)
        p2.paragraph_format.space_after = Pt(2)
        for run in p2.runs:
            run.font.size = Pt(10)
            run.font.color.rgb = RGBColor(0x44, 0x44, 0x44)
            run.font.name = '微软雅黑'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')

# ===== 标题区（带照片） =====
table = doc.add_table(rows=1, cols=2)
table.style = 'Table Grid'
# 去掉边框
tbl = table._tbl
tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement('w:tblPr')
tblBorders = OxmlElement('w:tblBorders')
for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
    border = OxmlElement(f'w:{border_name}')
    border.set(qn('w:val'), 'none')
    tblBorders.append(border)
tblPr.append(tblBorders)

cell_left = table.cell(0, 0)
cell_right = table.cell(0, 1)

# 左：文字信息
p_name = cell_left.paragraphs[0]
p_name.alignment = WD_ALIGN_PARAGRAPH.LEFT
run = p_name.add_run('特列努尔')
run.font.size = Pt(22)
run.font.bold = True
run.font.color.rgb = RGBColor(0x1a, 0x1a, 0x2e)
run.font.name = '微软雅黑'
run._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')

p2 = cell_left.add_paragraph()
run2 = p2.add_run('Java后端开发实习生  |  南京  |  可远程')
run2.font.size = Pt(12)
run2.font.color.rgb = RGBColor(0x25, 0x63, 0xeb)
run2.font.name = '微软雅黑'
run2._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')

p3 = cell_left.add_paragraph()
run3 = p3.add_run('18761851809  |  1079086024@qq.com  |  南京')
run3.font.size = Pt(11)
run3.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
run3.font.name = '微软雅黑'
run3._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')

# 右：照片（右上角）
p_photo = cell_right.paragraphs[0]
p_photo.alignment = WD_ALIGN_PARAGRAPH.RIGHT
run_photo = p_photo.add_run()
run_photo.add_picture(r'D:\openclaw_workspace\portfolio-gh\photo.jpg', width=Inches(0.9))

# 设置列宽
cell_left.width = Inches(5.0)
cell_right.width = Inches(1.3)

doc.add_paragraph()

# ===== 教育背景 =====
add_title(doc, '教育背景')
add_item(doc, '南京邮电大学', '2025级 · 大一', '物联网学院 · 网络工程  |  GPA 3.15 / 4.0  |  英语四级')

# ===== 专业技能 =====
add_title(doc, '专业技能')
skills = [
    ('Java后端', 'Spring Boot / MySQL / MyBatis / H2数据库 / 接口开发'),
    ('AI辅助开发', 'Vibe Coding / GitHub Copilot / Coze平台 / 硅基流动API(Qwen2.5-7B)'),
    ('前端 & 工具', 'Vue3 / Element Plus / Vite / Git / Docker'),
]
for title, desc in skills:
    p = doc.add_paragraph(style='List Bullet')
    run1 = p.add_run(f'{title}：')
    run1.font.bold = True
    run1.font.size = Pt(10)
    run1.font.name = '微软雅黑'
    run1._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    run2 = p.add_run(desc)
    run2.font.size = Pt(10)
    run2.font.color.rgb = RGBColor(0x44, 0x44, 0x44)
    run2.font.name = '微软雅黑'
    run2._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')

# ===== 比赛经历 =====
add_title(doc, '比赛经历')
add_item(doc, '中兴捧月 · 校园课程助手智能体（兴享智家）', '2026.05')
add_item(doc, '', '', 'AI智能助手，接入硅基流动Qwen2.5-7B，可查课程表、查成绩、问答。独立完成前后端开发，Java/Spring Boot后端5个接口，Vue3前端。')
add_item(doc, '计算机设计大赛 · 老人信息管理系统', '2026.04')
add_item(doc, '', '', '独立完成Java后端开发，提供5个REST API接口（增删改查），MySQL数据库。')

# ===== 项目经历 =====
add_title(doc, '项目经历')
add_item(doc, '星火工作室考核 · CStudyMate C语言AI学伴', '2026.05')
add_item(doc, '', '', 'Spring Boot + MyBatis + H2数据库，前端Vue3 + Element Plus，接入硅基流动Qwen2.5-7B实现AI问答。完成功能：AI问答、学情统计、练习中心。')
add_item(doc, 'TLIAS 智能学习平台', '2026.03')
add_item(doc, '', '', 'Spring Boot + Vue3 + JWT，用户登录认证与JWT令牌鉴权，Filter接口权限拦截，AOP操作日志记录。')

# ===== 个人评价 =====
add_title(doc, '个人评价')
evals = [
    '从新疆到南京，一路独立走过来，有完整全栈项目落地经验',
    'AI辅助编程能力：Vibe Coding、GitHub Copilot、Coze平台、硅基流动API，能利用AI工具快速定位问题、提升开发效率',
    '学习能力强，大一期间自学Spring Boot、Vue3、MyBatis并落地项目，目标暑假找Java后端实习攒经验',
]
for e in evals:
    p = doc.add_paragraph(style='List Bullet')
    run = p.add_run(e)
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(0x44, 0x44, 0x44)
    run.font.name = '微软雅黑'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')

doc.save(r'D:\openclaw_workspace\portfolio-gh\resume_final.docx')
print('done')