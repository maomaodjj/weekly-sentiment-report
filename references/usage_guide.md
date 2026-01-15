# 舆情周报生成使用指南

## 环境准备

### 安装依赖

```bash
pip3 install openpyxl python-docx
```

## 脚本使用

### 基本用法

```bash
python3 scripts/generate_report.py \
  --data-file "/path/to/main_data.xlsx" \
  --official-file "/path/to/official_reports.xlsx" \
  --output "/path/to/output_report.docx" \
  --start-date "2026年01月04日" \
  --end-date "2026年01月12日"
```

### 参数说明

- `--data-file`：主数据文件（必需），包含所有监测数据的 Excel 文件
- `--official-file`：官方媒体报道文件（可选），单独的官方媒体报道链接 Excel
- `--output`：输出文件路径（必需），生成的 Word 文档保存位置
- `--start-date`：开始日期（必需），格式：YYYY年MM月DD日
- `--end-date`：结束日期（必需），格式：YYYY年MM月DD日

## 完整示例

### 示例 1：使用素材目录中的文件

```bash
python3 scripts/generate_report.py \
  --data-file "/Users/yingmi/Desktop/舆情周报AI/素材：1月4日-1月12日舆情周报素材/20260105080000-20260112080000.xlsx" \
  --official-file "/Users/yingmi/Desktop/舆情周报AI/素材：1月4日-1月12日舆情周报素材/1月4日-1月12日官方媒体报道链接情况.xlsx" \
  --output "/Users/yingmi/Desktop/舆情周报AI/2026年1月4日-1月12日盈米基金舆情周报.docx" \
  --start-date "2026年01月04日" \
  --end-date "2026年01月12日"
```

### 示例 2：不使用官方媒体报道文件

```bash
python3 scripts/generate_report.py \
  --data-file "/path/to/main_data.xlsx" \
  --output "/path/to/output_report.docx" \
  --start-date "2026年01月04日" \
  --end-date "2026年01月12日"
```

## 输入文件要求

### 主数据文件（data-file）

必须是有效的 Excel 文件（.xlsx），包含以下工作表：

**盈米相关**：
- 主品牌-盈米基金（必需）
- 关键人物
- 主产品-且慢
- 且慢-主推产品
- 主产品-蜂鸟
- 主产品-啟明

**竞品**：
- E大、基金豆、翼支付、创必承、普益、7分钟理财、小帮规划
- 壹钱包、理财魔方、钱耳朵、新竹财富、朝财进宝、慢牛易富圈
- 智谷趋势、财商邦、菜鸟理财、价值发现者、阿牛定投、东方金匠
- 青澜家办、定投从零开始、盈添、智者理财

**合作伙伴**：
- 中航证券、粤开证券

**银行券商竞品**：
- 易方达、国联证券、中金财富、华泰证券、国泰君安
- 招商银行、平安银行、东方证券、天天基金、雪球
- 腾讯理财通、蚂蚁财富、基煜基金、汇成基金

**行业监管**：
- 监管政策法规、基金处罚违规

每个工作表的表头必须包含：
- 序号、监测专题名称、标题、发布时间、倾向性、来源网站、来源频道、作者
- 第 24 列（X 列）为摘要内容

### 官方媒体报道文件（official-file）

可选文件，包含以下列：
- 序号、媒体、发布时间、选题、标题、记者、署名、链接

## 配置文件

配置文件位于 `scripts/config.json`，可以自定义：

### 权威媒体列表

添加或删除权威媒体名称：

```json
"authoritative_media": [
  "中国证券报",
  "上海证券报",
  ...
]
```

### 转载网站列表

需要排除的转载网站：

```json
"repost_sites": [
  "证券之星",
  "东方财富网",
  ...
]
```

### 竞品工作表列表

指定需要读取的竞品工作表名称：

```json
"competitor_sheets": [
  "E大",
  "基金豆",
  ...
]
```

## 输出说明

生成的 Word 文档包含以下部分：

1. **标题**：珠海盈米基金销售有限公司舆情监测周报
2. **监测结果综述**：信息总量统计
3. **盈米基金重点信息**：
   - 序号、媒体、日期、标题
   - 摘要（仅包含盈米关键词）
   - 原文链接
4. **竞品要闻**：
   - 序号、标题
   - 机构名称标注
   - 媒体平台、发布时间、原文链接
5. **合作伙伴要闻**：格式同竞品要闻
6. **行业要闻**：格式同竞品要闻，增加类别标注
7. **备注**：数据来源、筛选规则说明
8. **报告生成时间**：自动生成的时间戳

## 常见问题

### Q: 如何修改权威媒体列表？

A: 编辑 `scripts/config.json` 文件中的 `authoritative_media` 数组。

### Q: 如何添加新的竞品工作表？

A: 在 Excel 中添加新工作表，然后在 `config.json` 的 `competitor_sheets` 数组中添加工作表名称。

### Q: 为什么某些新闻没有被包含？

A: 可能的原因：
- 来源不在权威媒体列表中
- 来源在转载网站列表中
- 标题包含公告关键词
- 与盈米新闻重复（针对竞品和合作伙伴）

### Q: 如何修改输出格式？

A: 编辑 `scripts/generate_report.py` 中的 `create_word_document()` 函数。

## 故障排除

### 错误：ModuleNotFoundError: No module named 'openpyxl'

**解决**：安装依赖库
```bash
pip3 install openpyxl python-docx
```

### 错误：Key: '主品牌-盈米基金' not found

**解决**：确保 Excel 文件包含"主品牌-盈米基金"工作表

### 警告：官方媒体报道文件不存在

**说明**：官方媒体报道文件是可选的，如果没有提供会显示警告但继续执行
