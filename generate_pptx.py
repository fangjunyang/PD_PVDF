from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# 主题色
BG = RGBColor(245, 247, 250)
DARK = RGBColor(20, 30, 40)
BLUE = RGBColor(42, 93, 167)
LIGHT = RGBColor(92, 134, 202)
GRAY = RGBColor(100, 110, 120)
ACCENT = RGBColor(45, 160, 120)

# 标题页
slide = prs.slides.add_slide(prs.slide_layouts[0])
slide.background.fill.solid()
slide.background.fill.fore_color.rgb = BG

# 标题
box = slide.shapes.add_textbox(Inches(0.8), Inches(1.2), Inches(11.5), Inches(1.2))
text = box.text_frame
text.text = '基于 PVDF 压电薄膜的智能配电柜局部放电超声在线监测'
text.paragraphs[0].font.size = Pt(28)
text.paragraphs[0].font.bold = True
text.paragraphs[0].font.color.rgb = DARK
text.paragraphs[0].alignment = PP_ALIGN.LEFT

# 副标题
box2 = slide.shapes.add_textbox(Inches(0.8), Inches(2.3), Inches(8.0), Inches(0.7))
text2 = box2.text_frame
text2.text = '建筑电气与智能化专业 | 学生汇报版'
text2.paragraphs[0].font.size = Pt(18)
text2.paragraphs[0].font.color.rgb = GRAY

# 颜色条
bar = slide.shapes.add_shape(1, Inches(0.8), Inches(3.3), Inches(10.0), Inches(0.12))
bar.fill.solid()
bar.fill.fore_color.rgb = BLUE
bar.line.fill.background()

# 结尾小字
box3 = slide.shapes.add_textbox(Inches(0.8), Inches(5.5), Inches(4.0), Inches(0.5))
text3 = box3.text_frame
text3.text = '2026年'
text3.paragraphs[0].font.size = Pt(16)
text3.paragraphs[0].font.color.rgb = GRAY

# 内容页模板

def add_title(slide, title, color=BLUE):
    box = slide.shapes.add_textbox(Inches(0.6), Inches(0.4), Inches(10.8), Inches(0.6))
    tf = box.text_frame
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(24)
    p.font.bold = True
    p.font.color.rgb = color

# 目录页
slide = prs.slides.add_slide(prs.slide_layouts[6])
slide.background.fill.solid()
slide.background.fill.fore_color.rgb = BG
add_title(slide, '目录')

bullets = [
    '1. 研究背景与课题意义',
    '2. 局部放电超声监测概述',
    '3. PVDF压电薄膜传感器工作原理',
    '4. 信号调理电路设计',
    '5. 典型应用案例',
    '6. 总结与展望'
]

textbox = slide.shapes.add_textbox(Inches(1.0), Inches(1.4), Inches(9.0), Inches(4.0))
tf = textbox.text_frame
for i, item in enumerate(bullets):
    p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
    p.text = item
    p.level = 0
    p.font.size = Pt(20)
    p.font.color.rgb = DARK
    p.bullet = True
    p.space_after = Pt(10)

# 背景页
slide = prs.slides.add_slide(prs.slide_layouts[6])
slide.background.fill.solid()
slide.background.fill.fore_color.rgb = BG
add_title(slide, '研究背景与课题意义')
textbox = slide.shapes.add_textbox(Inches(1.0), Inches(1.5), Inches(10.0), Inches(4.0))
tf = textbox.text_frame
items = [
    '配电柜是电力系统关键设备，承担供电、分配与控制任务。',
    '局部放电（PD）是绝缘老化和缺陷形成的早期信号。',
    '传统人工巡检难以及时发现隐性故障。',
    '在线监测能够实现早期预警和设备状态评估。',
    '本课题围绕 PVDF 压电薄膜传感器展开研究。'
]
for idx, item in enumerate(items):
    p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
    p.text = item
    p.level = 0
    p.font.size = Pt(20)
    p.font.color.rgb = DARK
    p.bullet = True
    p.space_after = Pt(12)

# 超声监测概述页
slide = prs.slides.add_slide(prs.slide_layouts[6])
slide.background.fill.solid(); slide.background.fill.fore_color.rgb = BG
add_title(slide, '局部放电超声监测概述')
textbox = slide.shapes.add_textbox(Inches(1.0), Inches(1.5), Inches(10.0), Inches(4.0))
tf = textbox.text_frame
items2 = [
    '局部放电会在绝缘缺陷处产生冲击性声学信号。',
    '超声信号可通过柜体结构传播并被传感器接收。',
    '该方法具有非侵入、在线、实时的优势。',
    '典型特征参数包括脉冲幅值、频率分布、能量包络等。',
    '本报告仅聚焦 PVDF 压电薄膜传感器。'
]
for idx, item in enumerate(items2):
    p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
    p.text = item
    p.level = 0; p.font.size = Pt(20); p.font.color.rgb = DARK; p.bullet = True; p.space_after = Pt(12)

# 工作原理页
slide = prs.slides.add_slide(prs.slide_layouts[6])
slide.background.fill.solid(); slide.background.fill.fore_color.rgb = BG
add_title(slide, 'PVDF压电薄膜传感器工作原理')
textbox = slide.shapes.add_textbox(Inches(0.8), Inches(1.5), Inches(6.4), Inches(3.6))
tf = textbox.text_frame
items3 = [
    'PVDF 是高分子压电材料。',
    '机械应力或振动会改变分子极化状态。',
    '产生电荷并形成电压信号。',
    '超声作用于薄膜后，输出与应变强度有关。',
    '这种电-机转换机制适合局部放电检测。'
]
for idx, item in enumerate(items3):
    p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
    p.text = item
    p.font.size = Pt(19); p.font.color.rgb = DARK; p.bullet = True; p.space_after = Pt(10)

# 右侧示意图（简单框图）
shape = slide.shapes.add_shape(1, Inches(7.7), Inches(1.9), Inches(4.0), Inches(2.5))
shape.fill.solid(); shape.fill.fore_color.rgb = LIGHT
shape.line.color.rgb = BLUE
shape.line.width = Pt(1.5)
# Add text inside
text_frame = shape.text_frame
text_frame.word_wrap = True
text_frame.text = '超声波\n→ PVDF薄膜\n→ 电荷/电压输出\n→ 传感器信号'
for p in text_frame.paragraphs:
    p.alignment = PP_ALIGN.CENTER
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255,255,255)

# 关键特性页
slide = prs.slides.add_slide(prs.slide_layouts[6])
slide.background.fill.solid(); slide.background.fill.fore_color.rgb = BG
add_title(slide, 'PVDF传感器关键特性')
textbox = slide.shapes.add_textbox(Inches(1.0), Inches(1.6), Inches(10.0), Inches(4.2))
tf = textbox.text_frame
items4 = [
    '柔性强，适合复杂表面安装。',
    '频率响应适合超声检测。',
    '安装方便，适合贴装式布置。',
    '工频和环境噪声较强，需合理调理。',
    '适合智能配电柜在线监测场景。'
]
for idx, item in enumerate(items4):
    p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
    p.text = item
    p.font.size = Pt(20); p.font.color.rgb = DARK; p.bullet = True; p.space_after = Pt(12)

# 信号调理页
slide = prs.slides.add_slide(prs.slide_layouts[6])
slide.background.fill.solid(); slide.background.fill.fore_color.rgb = BG
add_title(slide, '信号调理电路总体框架')
textbox = slide.shapes.add_textbox(Inches(1.0), Inches(1.6), Inches(10.0), Inches(3.6))
tf = textbox.text_frame
items5 = [
    'PVDF传感器 → 前置放大 → 滤波与抑制 → ADC采样',
    '前置放大器提高微弱电荷信号的信噪比。',
    '带通/低通滤波去除工频和高频杂波。',
    '整形和触发帮助捕捉局部放电脉冲。',
    '后端数字处理用于监测、判定与报警。'
]
for idx, item in enumerate(items5):
    p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
    p.text = item
    p.font.size = Pt(20); p.font.color.rgb = DARK; p.bullet = True; p.space_after = Pt(10)

# 典型应用案例页1
slide = prs.slides.add_slide(prs.slide_layouts[6])
slide.background.fill.solid(); slide.background.fill.fore_color.rgb = BG
add_title(slide, '典型应用案例一：实验室模拟试验')
textbox = slide.shapes.add_textbox(Inches(1.0), Inches(1.6), Inches(10.0), Inches(4.0))
tf = textbox.text_frame
items6 = [
    '在实验室模拟缺陷模型中进行局部放电试验。',
    'PVDF传感器贴装在柜体关键位置，采集超声脉冲。',
    '时域和频域分析能够识别异常放电信号。',
    '结果说明：PVDF传感器对局部放电响应灵敏。',
    '为现场在线监测提供了有效实验依据。'
]
for idx, item in enumerate(items6):
    p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
    p.text = item
    p.font.size = Pt(20); p.font.color.rgb = DARK; p.bullet = True; p.space_after = Pt(10)

# 典型应用案例页2
slide = prs.slides.add_slide(prs.slide_layouts[6])
slide.background.fill.solid(); slide.background.fill.fore_color.rgb = BG
add_title(slide, '典型应用案例二：配电柜现场在线监测')
textbox = slide.shapes.add_textbox(Inches(1.0), Inches(1.6), Inches(10.0), Inches(4.0))
tf = textbox.text_frame
items7 = [
    '动态布点安装在关键区域，例如母线与端子附近。',
    '持续监测局部放电超声信号并进行阈值判定。',
    '能够实现早期故障预警和状态评估。',
    '对电力设备的安全运行具有实际工程意义。',
    '其工程应用价值主要体现在在线、实时和非侵入。'
]
for idx, item in enumerate(items7):
    p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
    p.text = item
    p.font.size = Pt(20); p.font.color.rgb = DARK; p.bullet = True; p.space_after = Pt(10)

# 优点与局限页
slide = prs.slides.add_slide(prs.slide_layouts[6])
slide.background.fill.solid(); slide.background.fill.fore_color.rgb = BG
add_title(slide, '优点与局限')
textbox = slide.shapes.add_textbox(Inches(1.0), Inches(1.6), Inches(10.0), Inches(4.2))
tf = textbox.text_frame
items8 = [
    '优点：安装方便、柔性强、适合在线监测。',
    '优点：对超声信号响应灵敏。',
    '局限：环境噪声和机械振动可能影响检测。',
    '局限：安装位置和算法设计对判定结果影响明显。',
    '工程中应结合布置优化和多维分析进一步提高可靠性。'
]
for idx, item in enumerate(items8):
    p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
    p.text = item
    p.font.size = Pt(20); p.font.color.rgb = DARK; p.bullet = True; p.space_after = Pt(12)

# 总结页
slide = prs.slides.add_slide(prs.slide_layouts[6])
slide.background.fill.solid(); slide.background.fill.fore_color.rgb = BG
add_title(slide, '总结')
textbox = slide.shapes.add_textbox(Inches(1.0), Inches(1.6), Inches(10.0), Inches(4.2))
tf = textbox.text_frame
items9 = [
    'PVDF 压电薄膜是适合配电柜局部放电超声在线监测的优选传感器之一。',
    '它具备柔性、灵敏度和安装便利等优势。',
    '通过合理的调理电路和信号处理，可实现可靠的故障预警。',
    '在智能电网和智能配电柜发展中具有较强应用前景。',
    '该研究方案具有工程推广价值。'
]
for idx, item in enumerate(items9):
    p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
    p.text = item
    p.font.size = Pt(20); p.font.color.rgb = DARK; p.bullet = True; p.space_after = Pt(12)

# 结束页
slide = prs.slides.add_slide(prs.slide_layouts[6])
slide.background.fill.solid(); slide.background.fill.fore_color.rgb = BG
add_title(slide, '谢谢！')
textbox = slide.shapes.add_textbox(Inches(1.0), Inches(2.2), Inches(10.0), Inches(1.5))
tf = textbox.text_frame
p = tf.paragraphs[0]
p.text = 'Q&A'
p.font.size = Pt(32)
p.font.bold = True
p.font.color.rgb = BLUE
p.alignment = PP_ALIGN.CENTER

# 保存
prs.save('PVDF_配电柜局部放电在线监测.pptx')
print('PPTX generated successfully: PVDF_配电柜局部放电在线监测.pptx')
