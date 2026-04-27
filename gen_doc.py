#!/usr/bin/env python3
from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

doc = Document()

# 页面设置
for section in doc.sections:
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)

style = doc.styles['Normal']
style.font.name = 'Microsoft YaHei'
style.font.size = Pt(11)
style.paragraph_format.space_after = Pt(6)
style.paragraph_format.line_spacing = 1.5

# ── 封面 ──
doc.add_paragraph()
doc.add_paragraph()
title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title.add_run('Ozon 本土公司利润计算原理')
run.font.size = Pt(28)
run.font.bold = True
run.font.color.rgb = RGBColor(0, 91, 255)

subtitle = doc.add_paragraph()
subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = subtitle.add_run('ОСНО 一般税制 · 从中国发货 · 2026年最新政策')
run.font.size = Pt(14)
run.font.color.rgb = RGBColor(107, 114, 128)

doc.add_paragraph()
date_p = doc.add_paragraph()
date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = date_p.add_run('2026年4月')
run.font.size = Pt(12)
run.font.color.rgb = RGBColor(156, 163, 175)

doc.add_page_break()

# ── 目录 ──
doc.add_heading('目录', level=1)
toc_items = [
    '一、概述',
    '二、核心概念：单件落地成本',
    '三、从售价到净利润：完整扣费链路',
    '四、ОСНО 税务计算详解',
    '五、物流费用计算（FBO / FBS / rFBS）',
    '六、退货成本分摊',
    '七、盈亏平衡分析',
    '八、敏感性分析',
    '九、定价反推逻辑',
    '十、佣金率对利润的影响',
    '附录：示例参数与完整计算过程',
]
for item in toc_items:
    p = doc.add_paragraph(item)
    p.paragraph_format.space_after = Pt(2)

doc.add_page_break()

# ── 一、概述 ──
doc.add_heading('一、概述', level=1)
doc.add_paragraph(
    '本文档详细说明 Ozon 利润计算器的计算原理，适用于在俄罗斯 Ozon 平台以本土公司（ООО 或 ИП）身份销售商品的中国卖家。'
    '计算器支持 6 种俄罗斯税制，本文以最常用的 ОСНО（一般税制）为主线进行说明。'
)
doc.add_paragraph(
    '计算器的核心目标是：输入商品参数后，实时计算单件净利润、利润率、ROI，并提供月度经营数据、盈亏平衡分析、敏感性分析、定价反推和批次规划。'
)

# ── 二、核心概念 ──
doc.add_heading('二、核心概念：单件落地成本', level=1)
doc.add_paragraph('单件落地成本是每卖出一件商品的最低成本底线，由三部分组成：')

table = doc.add_table(rows=4, cols=3)
table.style = 'Light Grid Accent 1'
table.alignment = WD_TABLE_ALIGNMENT.CENTER
headers = ['成本项', '说明', '示例（₽）']
for i, h in enumerate(headers):
    table.rows[0].cells[i].text = h
data = [
    ['商品货值', '中国工厂出货价（已折算为卢布）', '600'],
    ['头程运费分摊', '中国→俄罗斯国际物流，按件均摊', '200'],
    ['清关费用', '海关关税 + 清关服务费，按件均摊', '0'],
]
for r, row_data in enumerate(data):
    for c, val in enumerate(row_data):
        table.rows[r+1].cells[c].text = val

doc.add_paragraph()
p = doc.add_paragraph()
run = p.add_run('落地成本 = 货值 + 头程 + 清关 = 600 + 200 + 0 = 800 ₽')
run.font.bold = True

# ── 三、完整扣费链路 ──
doc.add_heading('三、从售价到净利润：完整扣费链路', level=1)
doc.add_paragraph('以售价 2000₽、ОСНО 税制、FBO 发货模式为例，单件利润的计算过程如下：')

table = doc.add_table(rows=13, cols=4)
table.style = 'Light Grid Accent 1'
table.alignment = WD_TABLE_ALIGNMENT.CENTER
headers = ['序号', '费用项目', '计算方式', '金额（₽）']
for i, h in enumerate(headers):
    table.rows[0].cells[i].text = h

rows_data = [
    ['', '售价（起点）', '买家支付价', '+2,000.00'],
    ['①', 'Ozon 平台佣金', '售价 × 佣金率（9%）', '-180.00'],
    ['②', '收单手续费 Эквайринг', '售价 × 1.5%（免НДС）', '-30.00'],
    ['③', '物流配送费（FBO）', '仓储 + 拣货 + 最后一公里', '-88.60'],
    ['④', '退货物流分摊', '退货件数 × 退货物流费 ÷ 实际成交', '-3.16'],
    ['⑤', '退货损耗分摊', '退货件数 × 损耗率 × 落地成本 ÷ 实际成交', '-12.63'],
    ['⑥', '商品落地成本', '货值 + 头程 + 清关', '-800.00'],
    ['⑦', '广告推广分摊', '月广告费 ÷ 月实际成交件数', '-52.63'],
    ['⑧', '其他杂费分摊', '月杂费 ÷ 月实际成交件数', '-21.05'],
    ['⑨', 'НДС 增值税', '销项税 − 进项税（详见第四章）', '-62.89'],
    ['⑩', '所得税 25%', '应税利润 × 25%（详见第四章）', '-77.97'],
    ['', '净利润（终点）', '售价 − 所有扣费', '+671.07'],
]
for r, row_data in enumerate(rows_data):
    for c, val in enumerate(row_data):
        cell = table.rows[r+1].cells[c]
        cell.text = val
        if r == len(rows_data) - 1:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.bold = True
                    run.font.color.rgb = RGBColor(0, 185, 107)

doc.add_paragraph()
doc.add_paragraph('利润率 = 净利润 ÷ 售价 = 671.07 ÷ 2000 = 33.6%')
doc.add_paragraph('ROI = 净利润 ÷ 落地成本 = 671.07 ÷ 800 = 83.9%')

p = doc.add_paragraph()
run = p.add_run('注意：佣金和物流费按"下单件数"计算（包含退货），不是按成交件数。这是 Ozon 的实际扣费逻辑。')
run.font.italic = True
run.font.color.rgb = RGBColor(255, 105, 0)

# ── 四、ОСНО 税务 ──
doc.add_heading('四、ОСНО 税务计算详解', level=1)
doc.add_paragraph('ОСНО（一般税制）包含两层税，必须分别计算：')

doc.add_heading('4.1 第一层：НДС 增值税（22%）', level=2)
doc.add_paragraph('НДС 采用"销项减进项"的抵扣机制，实际税负远低于 22%。')

table = doc.add_table(rows=5, cols=2)
table.style = 'Light Grid Accent 1'
table.alignment = WD_TABLE_ALIGNMENT.CENTER
nds_data = [
    ['计算步骤', '公式与结果'],
    ['销项税（卖出时产生）', '售价 × 22/122 = 2000 × 22/122 = 360.66 ₽'],
    ['进项税基数（可抵扣成本）', '货物成本 + 佣金 + 物流 + 退货物流 + 广告'],
    ['进项税（买入时已付）', '进项基数 × 22/122'],
    ['应缴 НДС', 'max(销项税 − 进项税, 0)'],
]
for r, row_data in enumerate(nds_data):
    for c, val in enumerate(row_data):
        table.rows[r].cells[c].text = val

doc.add_paragraph()
p = doc.add_paragraph()
run = p.add_run('关键点：')
run.font.bold = True
doc.add_paragraph('• 收单手续费（Эквайринг）免 НДС，不能作为进项抵扣')
doc.add_paragraph('• 其他杂费（包装、人工）通常也不含 НДС 进项')
doc.add_paragraph('• 进项越多，实际缴纳的 НДС 越少')

doc.add_heading('4.2 第二层：所得税 налог на прибыль（25%）', level=2)
doc.add_paragraph('所得税基于"不含税利润"计算：')

table = doc.add_table(rows=4, cols=2)
table.style = 'Light Grid Accent 1'
table.alignment = WD_TABLE_ALIGNMENT.CENTER
it_data = [
    ['计算步骤', '公式'],
    ['不含税收入', '售价 × 100/122'],
    ['不含税支出', '进项基数 × 100/122 + 收单费 + 杂费'],
    ['所得税', 'max(不含税收入 − 不含税支出, 0) × 25%'],
]
for r, row_data in enumerate(it_data):
    for c, val in enumerate(row_data):
        table.rows[r].cells[c].text = val

doc.add_paragraph()
p = doc.add_paragraph()
run = p.add_run('总税负 = НДС + 所得税')
run.font.bold = True
run.font.size = Pt(13)

# ── 五、物流 ──
doc.add_heading('五、物流费用计算', level=1)
doc.add_paragraph('Ozon 提供三种发货模式，物流费用结构不同：')

doc.add_heading('5.1 FBO（平台仓发货）', level=2)
doc.add_paragraph('商品提前备货到 Ozon 仓库，由平台完成拣货、打包、配送。')
table = doc.add_table(rows=5, cols=3)
table.style = 'Light Grid Accent 1'
headers = ['费用项', '说明', '示例']
for i, h in enumerate(headers):
    table.rows[0].cells[i].text = h
fbo_data = [
    ['仓储费', '按天按件收取，与商品体积相关', '0.12 ₽/天/件'],
    ['库存天数', '商品在仓库的平均存放天数', '30 天'],
    ['拣货打包费', '每件订单的拣货和打包费用', '25 ₽/件'],
    ['最后一公里配送', '从仓库到买家的配送费', '60 ₽/件'],
]
for r, row_data in enumerate(fbo_data):
    for c, val in enumerate(row_data):
        table.rows[r+1].cells[c].text = val

doc.add_paragraph()
p = doc.add_paragraph()
run = p.add_run('FBO 物流费/件 = 仓储费×库存天数 + 拣货打包费 + 最后一公里 = 0.12×30 + 25 + 60 = 88.60 ₽')
run.font.bold = True

doc.add_heading('5.2 FBS（商家仓发货）', level=2)
doc.add_paragraph('商家自己存货，接单后交给 Ozon 分拣配送。')
doc.add_paragraph('FBS 物流费/件 = FBS 物流费 + FBS 处理费 = 120 + 25 = 145 ₽')

doc.add_heading('5.3 rFBS（直接快递）', level=2)
doc.add_paragraph('商家接单后自行通过快递直寄买家。')
doc.add_paragraph('rFBS 物流费/件 = 快递配送费 = 180 ₽')

# ── 六、退货 ──
doc.add_heading('六、退货成本分摊', level=1)
doc.add_paragraph('退货会产生两笔额外成本，需要分摊到每件实际成交的商品上：')

doc.add_paragraph('1. 退货物流费：Ozon 向卖家收取的退货运费')
doc.add_paragraph('   月退货物流 = 月下单量 × 退货率 × 退货物流单价')
doc.add_paragraph()
doc.add_paragraph('2. 退货损耗：退回商品中不可再售的部分')
doc.add_paragraph('   月退货损耗 = 月下单量 × 退货率 × 损耗率 × 单件落地成本')
doc.add_paragraph()

p = doc.add_paragraph()
run = p.add_run('示例（月销100件，退货率5%，损耗率30%）：')
run.font.bold = True
doc.add_paragraph('• 退货件数 = 100 × 5% = 5 件')
doc.add_paragraph('• 实际成交 = 100 − 5 = 95 件')
doc.add_paragraph('• 退货物流分摊/件 = 5 × 60 ÷ 95 = 3.16 ₽')
doc.add_paragraph('• 退货损耗分摊/件 = 5 × 30% × 800 ÷ 95 = 12.63 ₽')

# ── 七、盈亏平衡 ──
doc.add_heading('七、盈亏平衡分析', level=1)

doc.add_heading('7.1 盈亏平衡售价', level=2)
doc.add_paragraph('在当前成本结构下，让净利润恰好为 0 的最低售价。低于此价格就会亏损。')
doc.add_paragraph('计算方式：将所有变动成本（佣金、收单费、物流等）表示为售价的比例，求解使利润为零的售价。')

doc.add_heading('7.2 盈亏平衡销量', level=2)
doc.add_paragraph('在当前售价下，覆盖固定成本（广告费 + 杂费）所需的最低月销量。')
doc.add_paragraph('计算方式：月固定成本 ÷ 单件边际贡献（净利润 + 广告分摊 + 杂费分摊）')

# ── 八、敏感性 ──
doc.add_heading('八、敏感性分析', level=1)
doc.add_paragraph('敏感性分析展示关键参数变动对利润的影响，帮助卖家评估风险：')
doc.add_paragraph('• 售价 ±10% / ±20% 时，净利润和利润率的变化')
doc.add_paragraph('• 货值 ±10% / ±20% 时，净利润和利润率的变化')
doc.add_paragraph()
doc.add_paragraph('这两个参数是影响利润最敏感的变量。售价每提高 10%，利润率可能提升 5-8 个百分点；货值每降低 10%，利润率可能提升 3-5 个百分点。')

# ── 九、定价反推 ──
doc.add_heading('九、定价反推逻辑', level=1)
doc.add_paragraph('定价反推的目标：给定目标利润率，反推出需要设定的售价。')
doc.add_paragraph()
doc.add_paragraph('由于 ОСНО 的税务计算涉及非线性的销项/进项抵扣，无法用简单公式直接求解。计算器采用二分法（Binary Search）：')
doc.add_paragraph('1. 设定搜索范围：落地成本的 0.5 倍 ~ 50 倍')
doc.add_paragraph('2. 取中间值作为试探售价')
doc.add_paragraph('3. 用完整的计算逻辑算出该售价下的利润率')
doc.add_paragraph('4. 如果利润率低于目标，提高售价；如果高于目标，降低售价')
doc.add_paragraph('5. 重复 300 次迭代，精度达到 0.005%')

doc.add_paragraph()
doc.add_paragraph('计算器提供 0%~40% 共 9 档利润率的对照表，方便快速参考。')

# ── 十、佣金影响 ──
doc.add_heading('十、佣金率对利润的影响', level=1)
doc.add_paragraph('Ozon 不同类目的佣金率差异巨大（5%~55%），是影响利润的最关键因素之一。')

table = doc.add_table(rows=6, cols=3)
table.style = 'Light Grid Accent 1'
table.alignment = WD_TABLE_ALIGNMENT.CENTER
headers = ['类目', '佣金率', '利润影响']
for i, h in enumerate(headers):
    table.rows[0].cells[i].text = h
comm_data = [
    ['食品饮料', '5%', '利润空间最大，适合走量'],
    ['电子数码 / 汽车用品', '9%', '利润健康，主力类目'],
    ['家居日用品', '12%', '利润适中，竞争激烈'],
    ['服装鞋帽（1500~5000₽）', '45%', '利润极度压缩，需高售价支撑'],
    ['个人卫生用品', '55%', '几乎无利润空间，慎入'],
]
for r, row_data in enumerate(comm_data):
    for c, val in enumerate(row_data):
        table.rows[r+1].cells[c].text = val

doc.add_paragraph()
p = doc.add_paragraph()
run = p.add_run('建议：优先选择佣金率 ≤ 15% 的类目，佣金率超过 30% 的类目需要非常高的售价才能盈利。')
run.font.bold = True
run.font.color.rgb = RGBColor(255, 105, 0)

# ── 附录 ──
doc.add_page_break()
doc.add_heading('附录：示例参数与完整计算过程', level=1)

doc.add_heading('示例参数', level=2)
table = doc.add_table(rows=15, cols=2)
table.style = 'Light Grid Accent 1'
table.alignment = WD_TABLE_ALIGNMENT.CENTER
params = [
    ['参数', '值'],
    ['Ozon 售价', '2,000 ₽'],
    ['商品货值', '600 ₽'],
    ['头程运费', '200 ₽'],
    ['清关费用', '0 ₽'],
    ['商品类目', '电子数码（佣金 9%）'],
    ['月销量', '100 件'],
    ['退货率', '5%'],
    ['退货损耗率', '30%'],
    ['发货模式', 'FBO'],
    ['仓储费', '0.12 ₽/天/件'],
    ['库存天数', '30 天'],
    ['拣货打包费', '25 ₽/件'],
    ['最后一公里', '60 ₽/件'],
]
for r, row_data in enumerate(params):
    for c, val in enumerate(row_data):
        table.rows[r].cells[c].text = val

doc.add_heading('完整计算过程', level=2)

steps = [
    '1. 落地成本 = 600 + 200 + 0 = 800 ₽/件',
    '2. 退货件数 = 100 × 5% = 5 件',
    '3. 实际成交 = 100 − 5 = 95 件',
    '4. 月收入 = 95 × 2000 = 190,000 ₽',
    '5. 月佣金 = 100 × 2000 × 9% = 18,000 ₽（按下单件数算）',
    '6. 月收单费 = 190,000 × 1.5% = 2,850 ₽',
    '7. 月物流费 = 100 × (0.12×30 + 25 + 60) = 8,860 ₽（按下单件数算）',
    '8. 月退货物流 = 5 × 60 = 300 ₽',
    '9. 月退货损耗 = 5 × 30% × 800 = 1,200 ₽',
    '10. 月货物成本 = 95 × 800 + 1,200 = 77,200 ₽',
    '11. 月总成本 = 18,000 + 2,850 + 8,860 + 300 + 77,200 + 5,000 + 2,000 = 114,210 ₽',
    '',
    'НДС 计算：',
    '12. 销项税 = 190,000 × 22/122 = 34,262.30 ₽',
    '13. 进项基数 = 77,200 + 18,000 + 8,860 + 300 + 5,000 = 109,360 ₽',
    '14. 进项税 = 109,360 × 22/122 = 19,722.95 ₽',
    '15. 应缴 НДС = 34,262.30 − 19,722.95 = 14,539.35 ₽',
    '',
    '所得税计算：',
    '16. 不含税收入 = 190,000 × 100/122 = 155,737.70 ₽',
    '17. 不含税支出 = 109,360 × 100/122 + 2,850 + 2,000 = 94,487.05 ₽',
    '18. 应税利润 = 155,737.70 − 94,487.05 = 61,250.65 ₽',
    '19. 所得税 = 61,250.65 × 25% = 15,312.66 ₽',
    '',
    '最终结果：',
    '20. 月税费合计 = 14,539.35 + 15,312.66 = 29,852.01 ₽',
    '21. 月净利润 = 190,000 − 114,210 − 29,852.01 = 45,937.99 ₽',
    '22. 单件净利润 = 45,937.99 ÷ 95 = 483.56 ₽',
    '23. 利润率 = 483.56 ÷ 2000 = 24.2%',
    '24. ROI = 483.56 ÷ 800 = 60.4%',
]

for step in steps:
    if step == '':
        doc.add_paragraph()
    elif step.startswith('НДС') or step.startswith('所得税') or step.startswith('最终'):
        p = doc.add_paragraph()
        run = p.add_run(step)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0, 91, 255)
    else:
        doc.add_paragraph(step)

# ── 页脚说明 ──
doc.add_paragraph()
doc.add_paragraph()
p = doc.add_paragraph()
run = p.add_run('本文档由 Ozon 利润计算器自动生成 · 2026年4月')
run.font.size = Pt(9)
run.font.color.rgb = RGBColor(156, 163, 175)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER

# 保存
output_path = '/Users/swb/Downloads/ozon-profit-calc/Ozon利润计算原理.docx'
doc.save(output_path)
print(f'已生成: {output_path}')
