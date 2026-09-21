from pathlib import Path
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference, Series
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation


ROOT = Path(r"D:\chatGPT\实体3")
SOURCE = ROOT / "ai-digital-product-30day-playbook.html"
OUTPUT = ROOT / "ai-digital-product-30day-playbook-analysis.xlsx"

NAVY = "1F4E78"
BLUE = "D9EAF7"
LIGHT_BLUE = "EEF6FB"
GOLD = "FFF2CC"
GREEN = "E2F0D9"
RED = "FCE4D6"
GRAY = "F2F2F2"
WHITE = "FFFFFF"
INPUT_FILL = PatternFill("solid", fgColor=GOLD)
OUTPUT_FILL = PatternFill("solid", fgColor=BLUE)
HEADER_FILL = PatternFill("solid", fgColor=NAVY)
SECTION_FILL = PatternFill("solid", fgColor=LIGHT_BLUE)
THIN = Side(style="thin", color="D9E1F2")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def setup_sheet(ws, title, subtitle, widths, freeze="A4"):
    ws.sheet_view.showGridLines = False
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(widths))
    ws.cell(1, 1, title)
    ws.cell(1, 1).font = Font(size=18, bold=True, color=WHITE)
    ws.cell(1, 1).fill = HEADER_FILL
    ws.cell(1, 1).alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 32
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(widths))
    ws.cell(2, 1, subtitle)
    ws.cell(2, 1).font = Font(size=10, color="5C6672")
    ws.cell(2, 1).alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[2].height = 30
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.freeze_panes = freeze


def add_table(ws, headers, rows, start_row=3, input_cols=None, number_formats=None):
    input_cols = set(input_cols or [])
    for c, value in enumerate(headers, 1):
        cell = ws.cell(start_row, c, value)
        cell.fill = HEADER_FILL
        cell.font = Font(bold=True, color=WHITE)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[start_row].height = 28
    for r_idx, row in enumerate(rows, start_row + 1):
        for c_idx, value in enumerate(row, 1):
            cell = ws.cell(r_idx, c_idx, value)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = BORDER
            if c_idx in input_cols:
                cell.fill = INPUT_FILL
            elif isinstance(value, str) and value.startswith("="):
                cell.fill = OUTPUT_FILL
            if number_formats and c_idx in number_formats:
                cell.number_format = number_formats[c_idx]
    ws.auto_filter.ref = f"A{start_row}:{get_column_letter(len(headers))}{start_row + len(rows)}"
    return start_row + len(rows)


def add_section(ws, row, text, last_col):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=last_col)
    ws.cell(row, 1, text)
    ws.cell(row, 1).fill = SECTION_FILL
    ws.cell(row, 1).font = Font(bold=True, color=NAVY)
    ws.cell(row, 1).alignment = Alignment(vertical="center")
    ws.row_dimensions[row].height = 24


wb = Workbook()
wb.remove(wb.active)
wb.properties.title = "AI 数字产品小店商业模式分析与从零到一执行工作簿"
wb.properties.subject = "人民币、中国市场、创业者自用的顾问决策版"
wb.properties.creator = "Codex 商业模式分析"
wb.properties.description = "基于 ai-digital-product-30day-playbook.html 的结构化分析，包含三情景公式模型。"

ws = wb.create_sheet("00_结论摘要")
setup_sheet(ws, "00 结论摘要", "一句话判断：值得用不超过 1000 元做 30 天验证，但当前更接近“低成本现金流试验”，还不是已证明可规模化的数字产品生意。", [20, 24, 38, 38, 14, 18])
summary_rows = [
    ["核心结论", "值得做小规模验证", "原手册把选品、制作、上架、引流和逐日 SOP 打通，最大优势是启动成本和执行门槛低。", "先跑通 30 天，不以首月利润为目标。"],
    ["最佳机会", "电商内容模板包 + 卫浴垂直 SOP/服务", "通用模板负责获客和首单，卫浴建材场景有行业经验壁垒，适合做高客单利润款。", "首发主推 SKU 1，同步准备 SKU 2 和 SKU 3。"],
    ["主要风险", "流量不可控、交付范围失控、平台合规", "30 天目标只要求 3 单，说明转化仍是最大不确定项；低价代做容易吃掉时间。", "先卖标准品，只开有限修改次数，所有入口指向同一成交页。"],
    ["商业模型", "内容获客 -> 低价模板破冰 -> 标准包/垂直 SOP 赚钱 -> 代做和订阅复购", "数字产品边际成本低，但首月需要人工交付和大量内容测试。", "用贡献利润和回本周期判断是否加码，不只看订单数。"],
    ["30 天成功标准", "30 条内容、20 次咨询、3 单成交、1 个订阅、明确有效内容主题", "与手册里程碑一致，且能把“流量”转成“真实付费行为”。", "若未达标，先用一对一诊断复聊 10 位高意向用户。"],
    ["建议下一步", "先做 3 个动作：上架 9.9 元引流品、发布 3 条带钩子内容、主动私信 10 位目标客户", "这是最低成本获取真实反馈的路径。", "7 天内没有咨询，优先改标题、封面和钩子，不重做全套产品。"],
]
add_table(ws, ["分析项", "判断", "依据/解释", "行动建议"], summary_rows, 3)
ws["F3"] = "工作簿导航"
ws["F3"].fill = HEADER_FILL
ws["F3"].font = Font(bold=True, color=WHITE)
for i, name in enumerate(wb.sheetnames[1:], 4):
    ws.cell(i, 6, name).hyperlink = f"#'{name}'!A1"
    ws.cell(i, 6).font = Font(color="0563C1", underline="single")

ws = wb.create_sheet("01_商业模式画布")
setup_sheet(ws, "01 商业模式画布", "将手册中的商业模式显性化；重点看获客、付费、交付和复购能否形成闭环。", [18, 34, 34, 24, 14])
rows = [
    ["客户细分", "淘宝/拼多多/抖音店主、微商、实体店老板；卫浴/瓷砖/灯具/全屋定制门店与经销商。", "手册 SKU 1-3 的目标客户描述", "优先验证电商店主和卫浴门店老板", "高"],
    ["客户任务", "快速产出可发布内容；降低文案试错；让门店持续获客；用 AI 提升运营效率。", "模块一与模块四", "访谈 10 位用户验证最高频任务", "高"],
    ["痛点", "不会写、没时间写、请代运营贵、AI 工具有但不会用、担心没效果。", "候选商品与话术章节", "区分“学习需求”和“代做需求”", "高"],
    ["价值主张", "拿模板就能用，先低价体验，再升级标准包、垂直 SOP 和代做服务。", "SKU 1-5 与 Offer 阶梯", "用首单结果验证价值是否足够具体", "高"],
    ["渠道", "小红书内容与私信、朋友圈、面包多/爱发电成交、微信私域承接。", "模块三、四、七", "所有公开内容统一导流到微信或成交页", "高"],
    ["收入来源", "9.9 元引流品、19.9-69 元模板包、99-199 元 SOP、69 元代做体验、29.9 元/月订阅、399-599 元运营包。", "模块一、三", "首月只主推 2 个 SKU，避免选择过载", "高"],
    ["关键成本", "AI 工具、图片/字体版权、平台抽佣、营销测试、人工交付时间。", "模块八", "工具月成本控制在 300 元内", "中"],
    ["关键资源", "行业经验、案例素材、提示词和模板化交付流程、稳定内容产能。", "卫浴 SOP 差异化说明", "先沉淀案例库和标准交付模板", "高"],
    ["关键伙伴/平台", "DeepSeek/Kimi、飞书、Canva/稿定、面包多/爱发电、小红书、微信。", "模块八工具清单", "每个任务准备一个备用工具", "中"],
]
add_table(ws, ["模块", "核心判断", "原手册依据/假设", "待验证项", "优先级"], rows, 3)

ws = wb.create_sheet("02_用户需求分析")
setup_sheet(ws, "02 用户需求分析", "把“想用 AI”拆成可付费任务；优先服务那些现在就有发布/获客压力的人。", [18, 28, 28, 24, 22, 22, 16])
rows = [
    ["电商/内容运营新手", "快速写主图、详情页、种草文并按时发布", "不会写、没时间写、反复改稿", "免费模板、通用 AI、外包", "看到可直接套用的样例", "担心结果模板化、不会用", "高"],
    ["淘宝/拼多多/抖音店主", "用低成本内容带来咨询和成交", "内容产能低，投流成本高", "自己摸索、请兼职、代运营", "有行业案例和交付截图", "价格敏感、效果不确定", "高"],
    ["卫浴/建材门店老板", "让本地客户看到案例并到店咨询", "缺摄影、缺选题、不会持续更新", "装修公司代做、员工随便发", "同行真实案例和 30 天日历", "不确定小红书是否适合本地", "高"],
    ["经销商/渠道业务员", "给门店提供可复制运营支持", "需要批量赋能门店但缺工具", "品牌方物料、临时培训", "SOP 可复制、可培训", "内部落地执行力不足", "中"],
    ["买过模板但不用的人", "直接拿到可发布成品", "模板买了不会套用", "继续买工具、找代写", "先体验 69 元小额代做", "担心交付慢和修改无穷", "中"],
    ["想学 AI 的运营新人", "建立可复用的工作流", "信息碎片化、提示词质量不稳", "免费教程、社群、课程", "能解决具体工作问题", "只收藏不行动", "低"],
]
add_table(ws, ["人群", "核心任务/JTBD", "痛点", "当前替代方案", "付费触发", "购买阻力", "优先级"], rows, 3)

ws = wb.create_sheet("03_产品与价值主张")
setup_sheet(ws, "03 产品与价值主张", "每个 SKU 都必须回答：谁买、买完能做什么、为什么买你的、交付边界是什么。", [12, 22, 32, 30, 22, 22, 14, 14])
rows = [
    ["SKU 1", "电商小红书种草文模板包", "30 个标题公式、5 篇笔记模板、话题标签、使用说明、季度更新", "电商卖家专用、飞书+PDF、版本化更新", "不会写但愿意自己发的店主", "只想代写、完全不愿动手的人", "引流+标准", "A"],
    ["SKU 2", "电商主图/详情页提示词包", "主图卖点提示词、8 模块详情页提示词、2 个行业样例、违禁词检查清单", "填空式、配合样例、强调可用性", "有产品但不会提炼卖点的店主/运营", "无明确产品资料、期待自动成交的人", "主力利润", "A"],
    ["SKU 3", "卫浴门店小红书运营 SOP", "30 天日历、20 选题、门店模板、评论话术、飞书表", "行业经验壁垒、真实门店场景", "卫浴/建材门店与经销商", "非建材场景、预算极低的纯学习者", "利润+差异化", "B+"],
    ["SKU 4", "卫浴门店标题生成器", "50 个卫浴场景标题", "低价测试付款意愿和蓝海需求", "想低成本试水的小门店", "需要完整运营方案的人", "引流测试", "A-"],
    ["SKU 5", "代做体验+月度订阅", "1 条种草文、1 条主图文案、模板、后续订阅", "先体验后订阅、48 小时交付、3 轮修改", "买过模板但不肯自己动手的店主", "要求不限量、频繁改稿的人", "复购后端", "B"],
]
add_table(ws, ["编号", "产品", "交付内容", "差异化", "适用客户", "不适用客户", "商业角色", "结论"], rows, 3)

ws = wb.create_sheet("04_选品评分矩阵")
setup_sheet(ws, "04 选品评分矩阵", "在原手册四维打分基础上增加毛利、交付规模化、渠道匹配、壁垒、合规和复购维度；评分 1-5，权重合计 100%。", [18, 10, 24], freeze="A15")
criteria = [
    ["需求强度", 0.15, "是否高频、强痛点、有明确使用场景"],
    ["付费意愿", 0.15, "客户是否愿意为省时间和可直接使用付费"],
    ["首单速度", 0.15, "能否在 30 天内完成破冰成交"],
    ["毛利空间", 0.10, "扣平台费、履约和退款后是否仍有空间"],
    ["交付规模化", 0.10, "是否能模板化、低人力重复交付"],
    ["渠道匹配", 0.10, "是否适合小红书内容、私域和低预算冷启动"],
    ["竞争区隔", 0.10, "是否有垂直经验、案例或结构化壁垒"],
    ["合规安全性", 0.05, "版权、宣传、平台规则和退款风险是否可控"],
    ["复购/升单", 0.10, "是否自然导向订阅、代做或更高客单服务"],
]
add_table(ws, ["评分维度", "权重", "判断说明"], criteria, 3)
score_rows = [
    ["SKU 1 种草文模板包", 4, 4, 5, 5, 5, 5, 3, 5, 3],
    ["SKU 2 主图详情提示词包", 4, 4, 4, 5, 5, 4, 3, 5, 3],
    ["SKU 3 卫浴运营 SOP", 3, 4, 2, 5, 3, 3, 5, 4, 5],
    ["SKU 4 卫浴标题生成器", 3, 3, 5, 5, 5, 5, 4, 5, 2],
    ["SKU 5 代做体验+订阅", 4, 4, 2, 3, 2, 3, 4, 3, 5],
]
start = 15
headers = ["产品", "需求强度", "付费意愿", "首单速度", "毛利空间", "交付规模化", "渠道匹配", "竞争区隔", "合规安全性", "复购/升单", "加权总分", "等级", "结论"]
add_table(ws, headers, score_rows, start)
for r in range(start + 1, start + 1 + len(score_rows)):
    ws.cell(r, 11, f"=SUMPRODUCT(B{r}:J{r},$B$4:$B$12)")
    ws.cell(r, 11).number_format = "0.00"
    ws.cell(r, 12, f'=IF(K{r}>=4,"A",IF(K{r}>=3.2,"B","C"))')
for r, conclusion in zip(range(start + 1, start + 1 + len(score_rows)), ["首批主推", "同步准备", "高价值后置", "引流测试", "成交后测试"]):
    ws.cell(r, 13, conclusion)
for r in range(start + 1, start + 1 + len(score_rows)):
    for c in range(2, 11):
        ws.cell(r, c).fill = INPUT_FILL
        ws.cell(r, c).alignment = Alignment(horizontal="center")
    ws.cell(r, 11).fill = OUTPUT_FILL
    ws.cell(r, 12).fill = OUTPUT_FILL
ws.conditional_formatting.add(f"K{start+1}:K{start+len(score_rows)}", ColorScaleRule(start_type="num", start_value=2.5, start_color="F8696B", mid_type="num", mid_value=3.5, mid_color="FFEB84", end_type="num", end_value=4.5, end_color="63BE7B"))

ws = wb.create_sheet("05_产品组合与Offer阶梯")
setup_sheet(ws, "05 产品组合与 Offer 阶梯", "用低价品完成第一次付费，用标准品建立主营收入，用垂直 SOP 和服务拉高利润。", [12, 18, 22, 18, 26, 24, 26])
rows = [
    ["入口", "免费体验版", "10 个标题 + 1 篇模板", "0 元", "换微信/建立信任", "领取后 24 小时内跟进", "咨询率不低于 20%"],
    ["破冰", "SKU 1 / SKU 4", "低价模板包", "9.9-19.9 元", "完成首单和付款动作", "成交后推标准包", "首月至少 3 单"],
    ["主力", "SKU 1 / SKU 2", "完整模板与提示词包", "19.9-69 元", "贡献主要毛收入", "持续迭代版本", "至少 2 个 SKU 出单"],
    ["利润", "SKU 3", "卫浴门店运营 SOP", "99-199 元", "利用行业经验赚利润", "可加 1 次答疑", "有 2 个精准客户即可上线"],
    ["复购", "月度订阅", "每月更新场景包", "29.9 元/月", "提升 LTV", "达到 5 个付费用户后上线", "首月目标 1 个订阅"],
    ["高客单", "代做/运营包", "定制交付或月度运营", "399-599 元", "验证服务上限", "订单稳定后再开放", "交付占时不超 40% 后评估"],
]
add_table(ws, ["阶段", "产品", "形态", "价格", "商业目的", "升级动作", "成功标准"], rows, 3)

ws = wb.create_sheet("06_定价与套餐")
setup_sheet(ws, "06 定价与套餐", "价格不是越低越好；低价只用于破冰，标准包和垂直 SOP 承担利润。", [16, 16, 16, 16, 16, 20, 22, 18])
rows = [
    ["SKU 1 引流版", 9.9, 9.9, 0.5, 0.59, "新人券/限时", "数字商品发货后不退款", "用于测试付款意愿", "=IF(B4=0,0,(B4-D4-E4)/B4)"],
    ["SKU 1 标准版", 29.9, 19.9, 0.5, 1.79, "限时 19.9", "内容问题 48 小时内处理", "主推破冰品", "=IF(B5=0,0,(B5-D5-E5)/B5)"],
    ["SKU 2 提示词包", 49.0, 39.9, 0.5, 2.94, "和 SKU 1 组合优惠", "数字商品发货后不退款", "主力利润品", "=IF(B6=0,0,(B6-D6-E6)/B6)"],
    ["SKU 3 卫浴 SOP", 129.0, 99.0, 1.0, 7.74, "前 10 名 99 元", "可含 1 次答疑", "垂直利润品", "=IF(B7=0,0,(B7-D7-E7)/B7)"],
    ["SKU 5 代做体验", 69.0, 69.0, 25.0, 4.14, "买模板包可抵扣 19.9", "48 小时交付，3 轮修改", "服务测试品", "=IF(B8=0,0,(B8-D8-E8)/B8)"],
    ["月度订阅", 29.9, 29.9, 1.0, 1.79, "首月体验价", "按月更新，不承诺销量", "复购产品", "=IF(B9=0,0,(B9-D9-E9)/B9)"],
]
add_table(ws, ["套餐", "建议价", "最低可接受价", "直接履约成本", "平台费估算", "优惠策略", "退款/售后", "商业角色", "估算毛利率"], rows, 3, number_formats={2:"¥0.00",3:"¥0.00",4:"¥0.00",5:"¥0.00",9:"0.0%"})
ws.column_dimensions["I"].width = 14

ws = wb.create_sheet("07_单位经济模型")
setup_sheet(ws, "07 单位经济模型", "黄色单元格为输入项，蓝色为公式输出；刷新后用实际数据替换默认假设。", [28, 16, 14, 48])
inputs = [
    ["月度曝光", 3000, "次", "小红书/公众号/朋友圈合计可触达内容曝光"],
    ["主页点击率", 0.08, "%", "曝光后点击主页或内容的比例"],
    ["咨询率", 0.20, "%", "主页点击后进入私信/咨询的比例"],
    ["成交率", 0.15, "%", "咨询后完成付款的比例"],
    ["平均客单价", 29.9, "元", "首月以低价破冰，后续可通过组合包提升"],
    ["退款率", 0.05, "%", "数字商品与代做服务的综合退款比例"],
    ["平台/支付费率", 0.06, "%", "按上架平台实际抽佣替换"],
    ["单笔履约成本", 0.5, "元", "网盘、人工交付、素材等直接成本"],
    ["月度营销费用", 300, "元", "首月控制在 300 元内"],
    ["月度固定工具成本", 200, "元", "AI、图片、模板工具合计"],
    ["复购系数", 1.35, "倍", "订阅、复购和升单对 LTV 的放大系数"],
]
add_table(ws, ["输入项", "数值", "单位", "说明"], inputs, 3, input_cols=[2], number_formats={2:"0.00"})
for r in (5, 6, 7, 9, 10):
    ws.cell(r, 2).number_format = "0.0%"
for r in (4, 12, 13):
    ws.cell(r, 2).number_format = "0"
for r in (8, 11):
    ws.cell(r, 2).number_format = "¥0.00"
ws["B14"].number_format = "0.00"
outputs = [
    ["主页点击", "=B4*B5", "次"],
    ["咨询数", "=B17*B6", "人"],
    ["订单数", "=B18*B7", "单"],
    ["毛收入", "=B19*B8", "元"],
    ["退款金额", "=B20*B9", "元"],
    ["净收入", "=B20-B21", "元"],
    ["平台/支付费", "=B22*B10", "元"],
    ["履约成本", "=B19*B11", "元"],
    ["贡献利润", "=B22-B23-B24", "元"],
    ["单笔贡献利润", "=IF(B19=0,0,B25/B19)", "元"],
    ["CAC", "=IF(B19=0,0,B12/B19)", "元"],
    ["毛利率", "=IF(B22=0,0,B25/B22)", "%"],
    ["LTV", "=B26*B14", "元"],
    ["LTV/CAC", "=IF(B27=0,0,B29/B27)", "倍"],
    ["固定成本回本订单", "=IF(B26=0,0,CEILING(B13/B26,1))", "单"],
    ["盈亏平衡订单", "=IF(B26=0,0,CEILING((B12+B13)/B26,1))", "单"],
    ["月经营利润", "=B25-B12-B13", "元"],
    ["月经营利润率", "=IF(B22=0,0,B33/B22)", "%"],
]
start = 16
add_table(ws, ["输出项", "公式结果", "单位"], outputs, start)
for r in range(start + 1, start + 1 + len(outputs)):
    ws.cell(r, 2).fill = OUTPUT_FILL
    if r in (28, 34):
        ws.cell(r, 2).number_format = "0.0%"
    elif r in (17, 18, 19, 30, 31, 32):
        ws.cell(r, 2).number_format = "0.00"
    else:
        ws.cell(r, 2).number_format = "¥0.00"

ws = wb.create_sheet("08_三情景测算")
setup_sheet(ws, "08 三情景测算", "保守、基准、进取三套参数；本表所有结果由公式计算，修改输入即可联动。", [20, 14, 14, 14, 14, 16])
scenario_inputs = [
    ["月度曝光", 1500, 3000, 6000, "次"],
    ["主页点击率", 0.04, 0.08, 0.12, "%"],
    ["咨询率", 0.12, 0.20, 0.28, "%"],
    ["成交率", 0.10, 0.15, 0.22, "%"],
    ["平均客单价", 19.9, 29.9, 39.9, "元"],
    ["退款率", 0.06, 0.05, 0.04, "%"],
    ["平台/支付费率", 0.07, 0.06, 0.05, "%"],
    ["单笔履约成本", 0.5, 0.5, 0.5, "元"],
    ["月度营销费用", 200, 300, 500, "元"],
    ["月度固定工具成本", 100, 200, 300, "元"],
    ["复购系数", 1.20, 1.35, 1.60, "倍"],
]
add_table(ws, ["输入项", "保守", "基准", "进取", "单位"], scenario_inputs, 3, input_cols=[2, 3, 4], number_formats={2:"0.00",3:"0.00",4:"0.00"})
for r in (4, 7, 8, 9, 10, 12, 13, 14):
    ws.cell(r, 2).number_format = "0"
    ws.cell(r, 3).number_format = "0"
    ws.cell(r, 4).number_format = "0"
for r in (5, 6, 7, 9, 10):
    for c in (2, 3, 4):
        ws.cell(r, c).number_format = "0.0%"
for r in (8, 11):
    for c in (2, 3, 4):
        ws.cell(r, c).number_format = "¥0.00"
for c in (2, 3, 4):
    ws.cell(14, c).number_format = "0.00"
add_section(ws, 16, "公式输出", 6)
scenario_outputs = [
    ["主页点击", "={c}4*{c}5", "次"],
    ["咨询数", "={c}17*{c}6", "人"],
    ["订单数", "={c}18*{c}7", "单"],
    ["毛收入", "={c}19*{c}8", "元"],
    ["退款金额", "={c}20*{c}9", "元"],
    ["净收入", "={c}20-{c}21", "元"],
    ["平台/支付费", "={c}22*{c}10", "元"],
    ["履约成本", "={c}19*{c}11", "元"],
    ["贡献利润", "={c}22-{c}23-{c}24", "元"],
    ["单笔贡献利润", "=IF({c}19=0,0,{c}25/{c}19)", "元"],
    ["CAC", "=IF({c}19=0,0,{c}12/{c}19)", "元"],
    ["毛利率", "=IF({c}22=0,0,{c}25/{c}22)", "%"],
    ["LTV", "={c}26*{c}14", "元"],
    ["LTV/CAC", "=IF({c}27=0,0,{c}29/{c}27)", "倍"],
    ["固定成本回本订单", "=IF({c}26=0,0,CEILING({c}13/{c}26,1))", "单"],
    ["盈亏平衡订单", "=IF({c}26=0,0,CEILING(({c}12+{c}13)/{c}26,1))", "单"],
    ["月经营利润", "={c}25-{c}12-{c}13", "元"],
    ["月经营利润率", "=IF({c}22=0,0,{c}33/{c}22)", "%"],
]
for i, (label, formula, unit) in enumerate(scenario_outputs):
    r = 17 + i
    ws.cell(r, 1, label).border = BORDER
    ws.cell(r, 5, unit).border = BORDER
    for c_idx, col in zip((2, 3, 4), ("B", "C", "D")):
        cell = ws.cell(r, c_idx, formula.format(c=col))
        cell.fill = OUTPUT_FILL
        cell.border = BORDER
        cell.number_format = "0.00"
for r in (28, 34):
    for c in (2, 3, 4):
        ws.cell(r, c).number_format = "0.0%"
for r in (20, 21, 22, 23, 24, 25, 26, 27, 29, 33):
    for c in (2, 3, 4):
        ws.cell(r, c).number_format = "¥0.00"
for r in (31, 32):
    for c in (2, 3, 4):
        ws.cell(r, c).number_format = "0"
ws.conditional_formatting.add("E17:E34", CellIsRule(operator="greaterThan", formula=["0"], fill=OUTPUT_FILL))
add_section(ws, 36, "情景结论", 6)
conclusions = [
    ["保守", "先验证是否有人付款；低于 1 单就继续做内容和私信，不增加付费投放。"],
    ["基准", "达到手册首月最低目标附近；重点看 CAC 是否小于单笔贡献利润。"],
    ["进取", "只有在咨询和复购真实存在时才加码；不要让进取情景倒逼过度承诺或无限改稿。"],
]
add_table(ws, ["情景", "管理动作"], conclusions, 37)
add_section(ws, 42, "图表数据", 6)
chart_rows = [
    ["毛收入", "=C20", "=D20", "=E20"],
    ["贡献利润", "=C25", "=D25", "=E25"],
    ["月经营利润", "=C33", "=D33", "=E33"],
]
add_table(ws, ["指标", "保守", "基准", "进取"], chart_rows, 43)
chart = BarChart()
chart.type = "col"
chart.title = "三情景收入与利润"
chart.y_axis.title = "人民币"
chart.x_axis.title = "指标"
data = Reference(ws, min_col=2, max_col=4, min_row=43, max_row=46)
cats = Reference(ws, min_col=1, min_row=44, max_row=46)
chart.add_data(data, titles_from_data=True, from_rows=False)
chart.set_categories(cats)
chart.height = 7.5
chart.width = 15
ws.add_chart(chart, "H4")

ws = wb.create_sheet("09_获客渠道与漏斗")
setup_sheet(ws, "09 获客渠道与漏斗", "基准情景漏斗用于判断免费内容和私域是否足以带来第一批付费用户。", [16, 18, 16, 16, 18, 18, 22, 22])
funnel = [
    ["曝光", "='08_三情景测算'!D4", "", "来自基准情景"],
    ["主页点击", "='08_三情景测算'!D17", "=IF(B4=0,0,B5/B4)", "点击率"],
    ["咨询/私信", "='08_三情景测算'!D18", "=IF(B5=0,0,B6/B5)", "主页点击转咨询"],
    ["订单", "='08_三情景测算'!D19", "=IF(B6=0,0,B7/B6)", "咨询转成交"],
    ["复购/订阅估算", "=B7*0.25", "=IF(B7=0,0,B8/B7)", "按 25% 订单转化估算"],
]
add_table(ws, ["漏斗阶段", "数量", "阶段转化率", "口径"], funnel, 3, number_formats={2:"0.0",3:"0.0%"})
channel_rows = [
    ["小红书内容", "曝光+信任", "0-100 元", "高", "30 天发布 30 条，3 条带明确钩子", "有 1 篇 500+ 浏览且产生咨询", "连续 2 周无咨询则改标题/封面"],
    ["小红书私信", "高意向转化", "0 元", "高", "主动私信 20 位提问用户", "每天至少 3 次有效互动", "出现违规风险则改由评论答疑"],
    ["朋友圈", "熟人信任", "0 元", "中高", "连续 7 天展示产品、过程、案例", "至少 1 个熟人咨询", "无人互动则减少频率"],
    ["面包多/爱发电", "成交与自动发货", "平台抽佣 3%-10%", "高", "上架 SKU 1 和 SKU 4", "购买链接可用且自动发货", "抽佣或类目不符则换平台"],
    ["公众号", "内容沉淀", "0 元", "中", "同步长文，承接免费版领取", "7 天内新增 10 个关注", "无增长则只保留同步分发"],
    ["付费投放", "放大验证", "首月 0-200 元", "低", "只对已出单内容做小额测试", "CAC 小于单笔贡献利润", "首月无自然转化时不投放"],
    ["联盟/互推", "借力分发", "按成交分成", "中", "找同行交换体验版或做联合直播", "带来至少 3 个有效咨询", "对方用户不匹配则停止"],
]
add_table(ws, ["渠道", "作用", "预算/成本", "匹配度", "验证动作", "通过标准", "退出条件"], channel_rows, 10)
chart = BarChart()
chart.type = "bar"
chart.title = "基准获客漏斗"
chart.add_data(Reference(ws, min_col=2, min_row=4, max_row=8), titles_from_data=False)
chart.set_categories(Reference(ws, min_col=1, min_row=4, max_row=8))
chart.height = 7
chart.width = 13
ws.add_chart(chart, "J4")

ws = wb.create_sheet("10_30天执行计划")
setup_sheet(ws, "10 30 天执行计划", "前 7 天逐日执行，8-30 天按周目标拆成每日动作；状态列可下拉选择。", [8, 8, 18, 44, 26, 20, 12, 24, 12, 14])
plan_rows = [
    ["Day 0", "准备", "准备日", "检查设备与账号清单；注册工具；建立产品文档和数据复盘表；写一句 30 天目标。", "工具与账号可用；飞书 2 个文档", "电脑/手机/飞书", 200, "账号注册完成率 100%", "未开始", "创业者"],
    ["Day 1", "第1周", "基建日", "注册小红书、公众号、收款平台、飞书和至少 3 个 AI 工具；写入 5 个 SKU。", "5 行 SKU 表与平台账号", "小红书/公众号/飞书", 100, "5 个 SKU 完成建档", "未开始", "创业者"],
    ["Day 2", "第1周", "框架日", "建立种草文模板包目录；收集 30 个标题；归纳 10 个公式；测试提示词。", "产品目录和标题库", "小红书/飞书/DeepSeek", 0, "30 个标题、10 个公式", "未开始", "创业者"],
    ["Day 3", "第1周", "初稿日", "补齐 30 个标题公式；写 5 篇模板；整理 4 组话题标签；写说明页。", "产品 V1 初稿", "DeepSeek/飞书", 0, "产品目录全部有内容", "未开始", "创业者"],
    ["Day 4", "第1周", "完善日", "导出 PDF；做 3 张封面；制作免费体验版；建立 SKU 2 的 10 条提示词。", "PDF、封面、体验版", "飞书/Canva/稿定", 0, "可交付文件不少于 4 个", "未开始", "创业者"],
    ["Day 5", "第1周", "上架日", "上架 SKU 1 和 SKU 4；提交小红书店铺申请；写 5 个详情页；设置微信个人号。", "2 个购买链接与 5 份详情页", "面包多/爱发电/小红书/微信", 50, "2 个 SKU 可付款", "未开始", "创业者"],
    ["Day 6", "第1周", "内容生产日", "写 1 条 AI 日报、1 条工具实测、1 条垂直内容，并完成 3 张封面。", "3 篇待发布内容", "DeepSeek/Canva", 0, "3 篇内容可发布", "未开始", "创业者"],
    ["Day 7", "第1周", "发布复盘日", "发布 3 篇内容；发 2 条朋友圈；填第一周数据；确定下周 7 个选题。", "首周内容和复盘表", "小红书/微信/飞书", 0, "至少 1 份免费版发出", "未开始", "创业者"],
]
for day in range(8, 15):
    task = "发布 1 条内容并回复全部评论/私信；其中至少 3 天主动私信高意向用户并发送免费体验版。"
    if day == 10:
        task += " 根据评论反馈完成 1 次标题或说明页优化。"
    plan_rows.append([f"Day {day}", "第2周", "增长验证", task, "1 条内容、互动记录、1 次产品迭代", "小红书/微信/飞书", 0, "周内至少 1 篇 500+ 浏览、3 次咨询", "未开始", "创业者"])
for day in range(15, 22):
    task = "发布 1 条内容；回复私信；主动联系 2 位潜在客户；记录报价和成交阻力。"
    if day == 16:
        task += " 上架 9.9 元引流品并发布置顶领取内容。"
    if day == 19:
        task += " 上线 29.9 元月度订阅，先向已成交用户推荐。"
    plan_rows.append([f"Day {day}", "第3周", "转化测试", task, "咨询记录、报价记录、至少 1 单", "小红书/微信/面包多", 0, "20 份领取、5 次咨询、1 单", "未开始", "创业者"])
for day in range(22, 31):
    task = "稳定发布 1 条内容；回复互动；优化最佳 SKU 详情页；收集好评或交付截图。"
    if day == 25:
        task += " 对 10 位问过问题但未成交的人做二次跟进。"
    if day == 30:
        task += " 完成 30 天复盘，确定 2 个主推 SKU和有效内容主题。"
    plan_rows.append([f"Day {day}", "第4周", "复盘与放大", task, "30 条内容、复盘、2 个主推 SKU", "小红书/微信/飞书", 0, "累计 3 单、1 个订阅", "未开始", "创业者"])
add_table(ws, ["日期", "周次", "阶段", "关键任务", "产出物", "渠道/工具", "预算(元)", "核心指标", "状态", "负责人"], plan_rows, 3, input_cols=[9], number_formats={7:"¥0"})
last_plan_row = 3 + len(plan_rows)
dv = DataValidation(type="list", formula1='"未开始,进行中,已完成,已延期"', allow_blank=True)
ws.add_data_validation(dv)
dv.add(f"I4:I{last_plan_row}")
ws.conditional_formatting.add(f"I4:I{last_plan_row}", CellIsRule(operator="equal", formula=['"已完成"'], fill=PatternFill("solid", fgColor=GREEN)))
ws.conditional_formatting.add(f"I4:I{last_plan_row}", CellIsRule(operator="equal", formula=['"已延期"'], fill=PatternFill("solid", fgColor=RED)))

ws = wb.create_sheet("11_指标看板")
setup_sheet(ws, "11 指标看板", "先看真实行为和现金流，再看播放量；指标必须能触发下一步动作。", [20, 34, 18, 18, 28, 28])
metric_rows = [
    ["内容发布数", "实际发布的小红书/公众号内容条数", "30 条/30 天", "每日", "连续 2 天为 0", "立即补发草稿或最低限度内容"],
    ["内容有效率", "带来主页点击或咨询的内容占比", "≥10%", "每周", "低于 5%", "只保留最佳选题并改写标题/封面"],
    ["主页点击", "内容带来的主页访问次数", "基准 240 次/月", "每日", "连续 3 天为 0", "检查钩子是否在正文而非结尾"],
    ["咨询数", "私信、评论或微信中的有效咨询", "20 次/30 天", "每日", "7 天少于 2 次", "主动答疑 20 条热门笔记"],
    ["成交数", "完成付款的订单数", "3 单/30 天", "每日", "Day 21 仍为 0", "免费诊断 5 位高意向用户"],
    ["客单价", "成交金额/成交数", "≥29.9 元", "每日", "低于 19.9 元", "增加组合包和标准版推荐"],
    ["退款率", "退款金额/毛收入", "≤5%", "每周", "高于 8%", "检查交付范围与宣传承诺"],
    ["贡献利润", "净收入-平台费-履约成本", ">0", "每周", "连续 2 周为负", "停止低毛利渠道并提高客单"],
    ["CAC", "营销费用/新增订单", "<单笔贡献利润", "每周", "连续 2 周高于 LTV", "暂停付费投放，改自然内容"],
    ["LTV/CAC", "客户终身价值/获客成本", "≥2.0", "每月", "低于 1.2", "只保留自然流量和复购路径"],
]
add_table(ws, ["指标", "定义", "30 天目标", "频率", "红线", "触发动作"], metric_rows, 3)
add_section(ws, 16, "7 日数据录入示例（黄色区域可改）", 6)
daily_headers = ["日期", "曝光", "主页点击", "咨询", "订单", "收入", "营销费用", "成交率", "CAC"]
daily_rows = [[f"第{i}天", 0, 0, 0, 0, 0, 0, f"=IF(C{17+i}=0,0,E{17+i}/C{17+i})", f"=IF(E{17+i}=0,0,G{17+i}/E{17+i})"] for i in range(1, 8)]
add_table(ws, daily_headers, daily_rows, 17, input_cols=[2,3,4,5,6,7], number_formats={6:"¥0",7:"¥0",8:"0.0%",9:"¥0.00"})
chart = LineChart()
chart.title = "7 日订单变化"
chart.y_axis.title = "订单"
chart.add_data(Reference(ws, min_col=5, min_row=18, max_row=24), titles_from_data=False)
chart.set_categories(Reference(ws, min_col=1, min_row=18, max_row=24))
chart.height = 7
chart.width = 12
ws.add_chart(chart, "H17")

ws = wb.create_sheet("12_风险与合规审计")
setup_sheet(ws, "12 风险与合规审计", "风险管理采用概率 × 影响；高严重度项必须在放大投放或扩 SKU 前处理。", [22, 10, 10, 12, 10, 28, 34, 34, 14, 14])
risk_rows = [
    ["只发内容不接单", 4, 5, "=B4*C4", '=IF(D4>=15,"高",IF(D4>=8,"中","低"))', "有浏览但无咨询/订单", "每个内容入口统一指向主页、置顶和文末购买链接", "主动私信 20 位提问用户，先换取真实反馈", "创业者", "每周日"],
    ["低价代做挤占时间", 3, 5, "=B5*C5", '=IF(D5>=15,"高",IF(D5>=8,"中","低"))', "改稿超过 3 轮且仍在低价订单", "只卖套餐，写明 48 小时交付和 3 轮修改", "暂停接单，把流程模板化后再恢复", "交付负责人", "每周日"],
    ["内容与产品失衡", 3, 4, "=B6*C6", '=IF(D6>=15,"高",IF(D6>=8,"中","低"))', "连续 3 天没有产品更新或客户跟进", "每日固定 30 分钟产品/成交时间", "暂停资讯号 1 天，只处理产品和客户", "创业者", "每周日"],
    ["依赖单一 AI 工具", 3, 3, "=B7*C7", '=IF(D7>=15,"高",IF(D7>=8,"中","低"))', "工具限流导致断更或无法交付", "每个任务保留 1 个备用工具", "切换备用工具并记录替代流程", "创业者", "每月 1 日"],
    ["版权与原创风险", 3, 5, "=B8*C8", '=IF(D8>=15,"高",IF(D8>=8,"中","低"))', "直接搬运他人文案、图片或字体", "只用免费商用素材；改写并注明来源", "下架争议内容并重新制作", "内容负责人", "每月 1 日"],
    ["夸大宣传与退款纠纷", 3, 5, "=B9*C9", '=IF(D9>=15,"高",IF(D9>=8,"中","低"))', "承诺销量、爆款或保证结果", "不承诺具体销量；售后规则写进详情页", "停止相关话术，主动处理退款", "创业者", "上新前"],
    ["平台规则变化", 3, 4, "=B10*C10", '=IF(D10>=15,"高",IF(D10>=8,"中","低"))', "私信引流或店铺类目被限制", "不在私信高频发微信号；关注平台提示", "切换面包多/爱发电或内容渠道", "渠道负责人", "每周一"],
    ["数据隐私风险", 2, 4, "=B11*C11", '=IF(D11>=15,"高",IF(D11>=8,"中","低"))', "收集联系方式或客户资料无授权", "最小化收集；明确用途；不转卖数据", "删除无关数据并补充告知", "创业者", "每月 1 日"],
    ["现金流误判", 3, 4, "=B12*C12", '=IF(D12>=15,"高",IF(D12>=8,"中","低"))', "只看收入，不看平台费、退款和工具费", "每周更新单位经济模型", "减少非必要工具和投放", "财务负责人", "每周日"],
    ["过早扩 SKU", 4, 3, "=B13*C13", '=IF(D13>=15,"高",IF(D13>=8,"中","低"))', "2 个主推 SKU 未稳定就新增产品", "首月最多主推 2 个 SKU，其余只做后台测试", "冻结新品，集中优化出单产品的详情页", "创业者", "每两周"],
]
add_table(ws, ["风险", "概率(1-5)", "影响(1-5)", "严重度", "等级", "触发信号", "预防措施", "应急动作", "负责人", "复核日期"], risk_rows, 3, input_cols=[2,3], number_formats={4:"0"})
ws.conditional_formatting.add("D4:D13", ColorScaleRule(start_type="num", start_value=1, start_color="63BE7B", mid_type="num", mid_value=9, mid_color="FFEB84", end_type="num", end_value=25, end_color="F8696B"))
ws.conditional_formatting.add("E4:E13", CellIsRule(operator="equal", formula=['"高"'], fill=PatternFill("solid", fgColor=RED)))
ws.conditional_formatting.add("E4:E13", CellIsRule(operator="equal", formula=['"中"'], fill=PatternFill("solid", fgColor=GOLD)))
ws.conditional_formatting.add("E4:E13", CellIsRule(operator="equal", formula=['"低"'], fill=PatternFill("solid", fgColor=GREEN)))
add_section(ws, 16, "合规底线清单", 10)
compliance_rows = [
    ["经营范围", "早期可用个人身份在平台小规模经营；连续 2-3 个月稳定出单后评估个体工商户。", "当月复核"],
    ["收入记录", "从第一天记录日期、平台、金额、客户和退款。", "每日"],
    ["版权", "不搬运他人文案；图片和字体只用免费商用或已授权素材。", "每次上新"],
    ["AI 标识", "按平台要求标注 AI 生成或 AI 辅助。", "每次发布"],
    ["宣传合规", "不承诺销量、爆款、保证收益；不写绝对化用语。", "每次发布"],
    ["退款政策", "以平台规则为准；详情页写清数字商品交付和售后边界。", "每次上新"],
    ["私域引流", "遵守平台私信和联系方式规则，避免高频发送微信号。", "每周"],
]
add_table(ws, ["项目", "要求", "频率"], compliance_rows, 17)

ws = wb.create_sheet("13_原手册映射")
setup_sheet(ws, "13 原手册映射", "逐章保证分析可追溯：原手册负责执行细节，本工作簿负责商业判断、优先排序和验证边界。", [10, 16, 34, 34, 20, 14])
mapping_rows = [
    ["一", "选品策略，行 189-299", "四维打分、14 个候选商品、5 个首批 SKU、飞书+PDF 载体。", "原手册侧重可做性和执行门槛，缺少毛利、复购和交付负荷权重。", "保留 5 个 SKU；新增加权选品矩阵，先用 SKU 1 破冰、SKU 3 赚利润。", "最高"],
    ["二", "产品制作，行 300-477", "种草文模板、代做服务、工作流包、交付模板和报价。", "制作流程完整，但代做是人力型收入，容易被低价改稿拖住。", "标准品优先；代做只作为后端测试，明确 3 轮修改和 48 小时交付。", "高"],
    ["三", "上架销售，行 479-580", "面包多/爱发电主阵地，小红书店铺同步，微信私域承接。", "平台选择合理，但抽佣和类目规则需实时核验。", "所有成交入口统一；首月只上架 2 个 SKU 并测试自动发货。", "高"],
    ["四", "引流体系，行 582-758", "AI 资讯号做曝光，垂直内容做信任，知识付费做筛选。", "内容量设计足够，但风险是曝光多、购买入口弱。", "每条内容只保留一个核心钩子；每周复盘带来咨询的内容。", "最高"],
    ["五", "30 天计划，行 760-895", "前 7 天逐日，后 3 周按最低指标推进。", "目标门槛低，适合验证但不证明可规模化。", "保留最低指标，同时增加 CAC、贡献利润和回本订单监控。", "最高"],
    ["六", "每日 SOP，行 896-935", "工作日 2 小时、周末 4 小时、最低限度日 30 分钟。", "可执行性强，但需要防止内容生产挤占产品迭代和交付。", "每日至少保留 30 分钟给产品、成交和客户跟进。", "高"],
    ["七", "获客与转化话术，行 936-1002", "冷启动私信、评论区互动、客户沟通和推荐话术。", "话术丰富，但仍需用真实回复率验证。", "先测 3 套私信话术，只保留咨询率最高的一套。", "高"],
    ["八", "工具与费用，行 1003-1043", "首月工具费用可控在 300 元内，后续再采购高阶工具。", "与低成本验证策略一致；最大隐性成本是时间。", "按单位经济模型跟踪工具费和营销费，不做非必要订阅。", "中"],
    ["九", "风险与合规，行 1044-1087", "五类运营风险、营业执照、税务、版权、平台红线。", "风险识别较全，但缺少严重度和触发动作。", "转为概率×影响矩阵，并设置负责人和复核日期。", "高"],
    ["十", "里程碑与后续，行 1088-1134", "Day 3/7/14/21/30 指标，涨价、上新、自动化和全职判断标准。", "里程碑方向正确，但规模化需要至少 2-3 个月数据。", "首月只决定是否继续验证；涨价和扩品以连续出单为前提。", "高"],
]
add_table(ws, ["模块", "章节/行号", "原文核心", "顾问判断", "采用方式", "优先级"], mapping_rows, 3)

ws = wb.create_sheet("14_假设与数据来源")
setup_sheet(ws, "14 假设与数据来源", "明确区分原手册事实、顾问推断和待验证假设，避免把估值当结论。", [24, 20, 34, 14, 16, 34, 14])
assumption_rows = [
    ["预算上限", "原手册事实", "总预算 ≤1000 元，第 1 周控制在 100-300 元。", "高", "成本控制", "记录每周实际支出", "每周日"],
    ["首月目标", "原手册事实", "30 天卖出第一单，最低里程碑为 3 单和 1 个订阅。", "高", "验收目标", "指标看板每日更新", "每日"],
    ["价格带", "原手册事实", "9.9 元引流、19.9-69 元标准、99-199 元 SOP、29.9 元订阅。", "高", "定价", "按真实成交价修正", "首单后"],
    ["平台抽佣", "待验证", "原手册仅标注 3%-10%，单位经济暂按 6% 估算。", "中", "毛利润", "注册后核对实际费率", "上架前"],
    ["退款率", "顾问假设", "首月按 5% 估算，代做服务可能更高。", "中", "净收入", "记录退款原因和比例", "每周"],
    ["自然流量", "顾问假设", "基准 3000 次曝光、8% 主页点击、20% 咨询、15% 成交。", "低", "订单预测", "用 30 天真实数据替换", "Day 30"],
    ["履约成本", "顾问假设", "数字品 0.5 元/单，代做 25 元/单。", "中", "贡献利润", "按交付耗时重算", "每周"],
    ["工具成本", "顾问假设", "基准 200 元/月，进取 300 元/月。", "中", "现金流", "与费用总表对账", "每周"],
    ["复购系数", "顾问假设", "基准 1.35，用于 LTV 估算，不等于已实现复购。", "低", "LTV", "达到 5 个付费用户后验证", "第 2 个月"],
    ["CAC", "顾问假设", "自然流量按营销费用/订单估算，不含创业者时间成本。", "中", "投放判断", "加入时间成本做二次核算", "Day 30"],
    ["市场与币种", "已确认口径", "中国大陆市场，人民币，创业者自用。", "高", "全部财务表", "不需要调整", "固定"],
    ["不承诺销量", "原手册事实与合规要求", "手册明确不承诺爆款和具体销量。", "高", "营销合规", "详情页和话术逐条检查", "每次发布"],
]
add_table(ws, ["项目", "类型", "内容", "置信度", "影响范围", "验证方法", "复核时间"], assumption_rows, 3)

wb.calculation.fullCalcOnLoad = True
wb.active = 0

wb.save(OUTPUT)
print(f"Created: {OUTPUT}")
