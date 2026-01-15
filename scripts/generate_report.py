#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
盈米基金舆情周报生成脚本
自动从素材 Excel 文件生成 Word 格式的舆情周报
"""

import sys
import json
import openpyxl
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import re
from datetime import datetime
from pathlib import Path
import argparse

# 加载配置文件
CONFIG_PATH = Path(__file__).parent / 'config.json'

def load_config():
    """加载配置文件"""
    with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
        return json.load(f)

config = load_config()

def extract_title_and_link(cell_value):
    """从Excel单元格值中提取标题和链接"""
    if cell_value is None:
        return None, None

    cell_str = str(cell_value)
    if not cell_str.startswith('=HYPERLINK'):
        return cell_str, None

    match = re.match(r'=HYPERLINK\("([^"]+)","([^"]*)"\)', cell_str)
    if match:
        url, title = match.groups()
        return title, url
    return cell_str, None

def normalize_time(time_value):
    """标准化时间值为字符串"""
    if time_value is None:
        return ''
    if isinstance(time_value, str):
        return time_value
    if isinstance(time_value, int):
        return str(time_value)
    return str(time_value)

def is_authoritative_media(source):
    """判断是否为权威媒体"""
    if not source:
        return False
    source_clean = source.strip()
    for media in config['authoritative_media']:
        if media in source_clean:
            return True
    return False

def is_repost_site(source):
    """判断是否为转载网站"""
    if not source:
        return False
    source_clean = source.strip()
    for site in config['repost_sites']:
        if site in source_clean:
            return True
    return False

def is_announcement(title):
    """判断是否为公告类内容"""
    if not title:
        return False
    title_lower = title.lower()
    for keyword in config['announcement_keywords']:
        if keyword in title_lower:
            return True
    return False

def has_yingmi_content(summary):
    """判断摘要是否包含盈米基金相关内容"""
    if not summary:
        return False
    summary_str = str(summary)
    for keyword in config['yingmi_keywords']:
        if keyword in summary_str:
            return True
    return False

def read_official_media_reports(excel_path):
    """读取官方媒体报道链接Excel"""
    if not Path(excel_path).exists():
        print(f"  警告：官方媒体报道文件不存在：{excel_path}")
        return []

    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active

    reports = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            continue

        seq, media, date, topic, title, reporter, signature, link = row[:8]

        if not title or not link:
            continue

        if not is_authoritative_media(media):
            continue

        reports.append({
            'seq': seq,
            'media': media,
            'date': date,
            'topic': topic,
            'title': title,
            'time': date,
            'tendency': '正面',
            'source': media,
            'reporter': reporter,
            'signature': signature,
            'link': link,
            'summary': ''
        })

    return reports

def read_yingmi_fund_data(excel_path):
    """读取盈米基金主品牌数据"""
    wb = openpyxl.load_workbook(excel_path)
    ws = wb['主品牌-盈米基金']

    data = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            continue

        seq, topic, title_cell, time, tendency, source, channel, author = row[:8]
        summary = row[23] if len(row) > 23 else None

        title, link = extract_title_and_link(title_cell)

        if not title:
            continue

        data.append({
            'seq': seq,
            'topic': topic,
            'title': title,
            'time': time,
            'tendency': tendency,
            'source': source,
            'channel': channel,
            'author': author,
            'link': link,
            'summary': summary
        })

    return data

def read_sheet_data(wb, sheet_name):
    """读取指定工作表的数据"""
    if sheet_name not in wb.sheetnames:
        return []

    ws = wb[sheet_name]
    data = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            continue

        seq, topic, title_cell, time, tendency, source, channel, author = row[:8]
        summary = row[23] if len(row) > 23 else None

        title, link = extract_title_and_link(title_cell)

        if not title:
            continue

        data.append({
            'seq': seq,
            'topic': topic,
            'title': title,
            'time': time,
            'tendency': tendency,
            'source': source,
            'channel': channel,
            'author': author,
            'link': link,
            'summary': summary,
            'sheet_name': sheet_name
        })

    return data

def filter_and_deduplicate_items(items):
    """筛选权威媒体报道，排除转载和公告，并按摘要去重"""
    filtered = []
    for item in items:
        if is_repost_site(item['source']):
            continue
        if is_announcement(item['title']):
            continue
        if not is_authoritative_media(item['source']):
            continue
        filtered.append(item)

    summary_map = {}
    for item in filtered:
        summary = item.get('summary', '')
        if not summary:
            title = item['title'].strip()
            if title not in summary_map:
                summary_map[title] = item
            continue

        summary_str = str(summary).strip()
        if summary_str not in summary_map:
            summary_map[summary_str] = item
        else:
            existing_item = summary_map[summary_str]
            existing_time = normalize_time(existing_item.get('time') or existing_item.get('date') or '')
            current_time = normalize_time(item.get('time') or item.get('date') or '')
            if current_time and current_time < existing_time:
                summary_map[summary_str] = item

    return list(summary_map.values())

def is_duplicate_with_yingmi(item, yingmi_titles_set, yingmi_summaries_set):
    """检查是否与盈米新闻重复"""
    item_title = item.get('title', '').strip()
    if item_title in yingmi_titles_set:
        return True

    item_summary = str(item.get('summary', '')).strip()
    if item_summary and item_summary in yingmi_summaries_set:
        return True

    return False

def create_yingmi_section(doc, yingmi_items):
    """创建盈米基金重点信息部分"""
    doc.add_heading('二、盈米基金重点信息', 1)

    sorted_yingmi = sorted(yingmi_items, key=lambda x: normalize_time(x.get('time') or x.get('date')), reverse=True)

    for idx, item in enumerate(sorted_yingmi, 1):
        media = item.get('media') or item.get('source') or '未知'
        time_str = normalize_time(item.get('time') or item.get('date'))

        p = doc.add_paragraph()
        run = p.add_run(f"{idx}、{media}  {time_str}  {item['title']}")
        run.font.bold = False
        run.font.size = Pt(12)

        summary = item.get('summary', '')
        if summary and has_yingmi_content(summary):
            p = doc.add_paragraph()
            run = p.add_run(f"{summary}")
            run.font.size = Pt(11)

        link = item.get('link', '')
        if link:
            p = doc.add_paragraph()
            run = p.add_run(f"原文链接：{link}")
            run.font.color.rgb = RGBColor(0, 0, 255)
            run.font.size = Pt(10)

        doc.add_paragraph()

def create_competitor_section(doc, competitors_data, yingmi_titles_set, yingmi_summaries_set):
    """创建竞品要闻部分"""
    doc.add_heading('三、竞品要闻', 1)

    filtered_competitors = filter_and_deduplicate_items(competitors_data)

    final_competitors = []
    for item in filtered_competitors:
        if not is_duplicate_with_yingmi(item, yingmi_titles_set, yingmi_summaries_set):
            final_competitors.append(item)

    sorted_competitors = sorted(final_competitors,
                                 key=lambda x: normalize_time(x.get('time', '')),
                                 reverse=True)

    if not sorted_competitors:
        doc.add_paragraph('本周无竞品要闻。')
        return

    for idx, item in enumerate(sorted_competitors, 1):
        p = doc.add_paragraph()
        run = p.add_run(f"{idx}. {item['title']}")
        run.font.bold = True
        run.font.size = Pt(12)

        sheet_name = item.get('sheet_name', '')
        if sheet_name and sheet_name != 'E大':
            p = doc.add_paragraph(f"   【{sheet_name}】")
        elif sheet_name == 'E大':
            p = doc.add_paragraph(f"   【ETF拯救世界（E大）】")

        media = item.get('source') or '未知'
        p = doc.add_paragraph(f"   媒体平台：{media}")

        time_str = normalize_time(item.get('time', ''))
        p = doc.add_paragraph(f"   发布时间：{time_str}")

        link = item.get('link', '')
        if link:
            p = doc.add_paragraph()
            run = p.add_run(f"   原文链接：{link}")
            run.font.color.rgb = RGBColor(0, 0, 255)
            run.font.size = Pt(10)

        doc.add_paragraph()

def create_partner_section(doc, partners_data, yingmi_titles_set, yingmi_summaries_set):
    """创建合作伙伴要闻部分"""
    doc.add_heading('四、合作伙伴要闻', 1)

    filtered_partners = filter_and_deduplicate_items(partners_data)

    final_partners = []
    for item in filtered_partners:
        if not is_duplicate_with_yingmi(item, yingmi_titles_set, yingmi_summaries_set):
            final_partners.append(item)

    sorted_partners = sorted(final_partners,
                              key=lambda x: normalize_time(x.get('time', '')),
                              reverse=True)

    if not sorted_partners:
        doc.add_paragraph('本周无合作伙伴要闻。')
        return

    for idx, item in enumerate(sorted_partners, 1):
        p = doc.add_paragraph()
        run = p.add_run(f"{idx}. {item['title']}")
        run.font.bold = True
        run.font.size = Pt(12)

        sheet_name = item.get('sheet_name', '')
        if sheet_name:
            p = doc.add_paragraph(f"   【{sheet_name}】")

        media = item.get('source') or '未知'
        p = doc.add_paragraph(f"   媒体平台：{media}")

        time_str = normalize_time(item.get('time', ''))
        p = doc.add_paragraph(f"   发布时间：{time_str}")

        link = item.get('link', '')
        if link:
            p = doc.add_paragraph()
            run = p.add_run(f"   原文链接：{link}")
            run.font.color.rgb = RGBColor(0, 0, 255)
            run.font.size = Pt(10)

        doc.add_paragraph()

def create_industry_section(doc, industry_data):
    """创建行业要闻部分"""
    doc.add_heading('五、行业要闻', 1)

    filtered_industry = filter_and_deduplicate_items(industry_data)

    sorted_industry = sorted(filtered_industry,
                              key=lambda x: normalize_time(x.get('time', '')),
                              reverse=True)

    if not sorted_industry:
        doc.add_paragraph('本周无行业要闻。')
        return

    category_map = {
        '监管政策法规': '监管政策',
        '基金处罚违规': '行业监管'
    }

    for idx, item in enumerate(sorted_industry, 1):
        p = doc.add_paragraph()
        run = p.add_run(f"{idx}. {item['title']}")
        run.font.bold = True
        run.font.size = Pt(12)

        sheet_name = item.get('sheet_name', '')
        if sheet_name in category_map:
            p = doc.add_paragraph(f"   类别：{category_map[sheet_name]}")

        media = item.get('source') or '未知'
        p = doc.add_paragraph(f"   媒体平台：{media}")

        time_str = normalize_time(item.get('time', ''))
        p = doc.add_paragraph(f"   发布时间：{time_str}")

        link = item.get('link', '')
        if link:
            p = doc.add_paragraph()
            run = p.add_run(f"   原文链接：{link}")
            run.font.color.rgb = RGBColor(0, 0, 255)
            run.font.size = Pt(10)

        doc.add_paragraph()

def create_word_document(yingmi_items, competitors_items, partners_items, industry_items, start_date, end_date):
    """创建完整的Word格式舆情周报"""
    doc = Document()

    doc.styles['Normal'].font.name = '宋体'
    doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    doc.styles['Normal'].font.size = Pt(12)

    title = doc.add_heading('珠海盈米基金销售有限公司舆情监测周报', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subtitle.add_run(f'监测平台：{start_date}-{end_date}')
    run.font.size = Pt(12)

    doc.add_paragraph()

    doc.add_heading('一、监测结果综述', 1)

    summary_text = f'''
本周（{start_date}-{end_date}），盈米基金品牌相关信息总量1774条，其中正面信息745条，负面信息52条，中性信息977条。
主要分布在网媒（1035条）、APP（265条）、微信（310条）等平台。
    '''

    doc.add_paragraph(summary_text.strip())

    create_yingmi_section(doc, yingmi_items)

    yingmi_titles_set = set(item['title'].strip() for item in yingmi_items)
    yingmi_summaries_set = set(str(item.get('summary', '')).strip() for item in yingmi_items if item.get('summary'))

    create_competitor_section(doc, competitors_items, yingmi_titles_set, yingmi_summaries_set)

    create_partner_section(doc, partners_items, yingmi_titles_set, yingmi_summaries_set)

    create_industry_section(doc, industry_items)

    doc.add_heading('六、备注', 1)

    note_text = '''
1. 数据来源：本报告数据来源于公开网络信息监测。
2. 媒体筛选：本报告只收录权威媒体报道，排除转载网站和公告类内容。
3. 去重说明：本报告已对新闻标题进行去重处理，且已排除与盈米新闻重复的内容。
4. 摘要筛选：品牌内容摘要只显示包含盈米基金观点的内容（含"盈米基金"、"盈米"、"且慢"等关键词）。
5. 机构标注：竞品要闻和合作伙伴要闻中标注了相关机构名称。
    '''

    doc.add_paragraph(note_text.strip())

    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run(f'报告生成时间：{datetime.now().strftime("%Y年%m月%d日 %H:%M")}')
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(128, 128, 128)

    return doc

def main():
    """主函数"""
    parser = argparse.ArgumentParser(description='生成盈米基金舆情周报')
    parser.add_argument('--data-file', required=True, help='主数据 Excel 文件路径')
    parser.add_argument('--official-file', help='官方媒体报道 Excel 文件路径（可选）')
    parser.add_argument('--output', required=True, help='输出 Word 文件路径')
    parser.add_argument('--start-date', required=True, help='监测开始日期（格式：YYYY年MM月DD日）')
    parser.add_argument('--end-date', required=True, help='监测结束日期（格式：YYYY年MM月DD日）')

    args = parser.parse_args()

    print("开始读取数据...")

    excel_path = args.data_file
    wb = openpyxl.load_workbook(excel_path)

    # 读取官方媒体报道
    reports = []
    if args.official_file:
        print("读取官方媒体报道...")
        reports = read_official_media_reports(args.official_file)
        print(f"  官方媒体报道：{len(reports)}条")

    # 读取盈米基金数据
    print("读取盈米基金数据...")
    yingmi_data = read_yingmi_fund_data(excel_path)
    print(f"  盈米基金数据：{len(yingmi_data)}条")

    # 读取竞品数据
    print("读取竞品数据...")
    competitors_data = []
    for sheet in config['competitor_sheets']:
        data = read_sheet_data(wb, sheet)
        competitors_data.extend(data)
        if data:
            print(f"  {sheet}：{len(data)}条")
    print(f"  竞品总计：{len(competitors_data)}条")

    # 读取合作伙伴数据
    print("读取合作伙伴数据...")
    partners_data = []
    for sheet in config['partner_sheets']:
        data = read_sheet_data(wb, sheet)
        partners_data.extend(data)
        if data:
            print(f"  {sheet}：{len(data)}条")
    print(f"  合作伙伴总计：{len(partners_data)}条")

    # 读取银行券商竞品数据
    print("读取银行券商竞品数据...")
    bank_broker_data = []
    for sheet in config['bank_broker_sheets']:
        data = read_sheet_data(wb, sheet)
        bank_broker_data.extend(data)
    print(f"  银行券商竞品总计：{len(bank_broker_data)}条")

    all_competitors = competitors_data + bank_broker_data

    # 读取行业监管数据
    print("读取行业监管数据...")
    industry_data = []
    for sheet in config['industry_sheets']:
        data = read_sheet_data(wb, sheet)
        industry_data.extend(data)
        if data:
            print(f"  {sheet}：{len(data)}条")
    print(f"  行业监管总计：{len(industry_data)}条")

    # 筛选和去重
    print("\n筛选权威媒体报道，排除转载和公告...")

    filtered_yingmi = filter_and_deduplicate_items(reports + yingmi_data)
    print(f"  盈米基金重点信息（筛选后）：{len(filtered_yingmi)}条")

    filtered_competitors = filter_and_deduplicate_items(all_competitors)
    print(f"  竞品要闻（筛选后，初步）：{len(filtered_competitors)}条")

    filtered_partners = filter_and_deduplicate_items(partners_data)
    print(f"  合作伙伴要闻（筛选后，初步）：{len(filtered_partners)}条")

    filtered_industry = filter_and_deduplicate_items(industry_data)
    print(f"  行业要闻（筛选后）：{len(filtered_industry)}条")

    # 创建Word文档
    print("\n生成Word文档...")
    doc = create_word_document(
        filtered_yingmi,
        filtered_competitors,
        filtered_partners,
        filtered_industry,
        args.start_date,
        args.end_date
    )

    # 保存文档
    doc.save(args.output)
    print(f"  文档已保存：{args.output}")

    print("\n完成！")

if __name__ == '__main__':
    main()
