# 盈米基金舆情周报生成 Skill

自动生成盈米基金舆情监测周报的 Claude Skill。

## 功能特性

- 📊 **自动数据读取**：从 Excel 素材文件读取多工作表数据
- 🔍 **智能筛选**：筛选权威媒体报道，排除转载和公告
- 🔄 **自动去重**：按标题和摘要去重，保留最早版本
- 📝 **统一格式**：所有模块使用统一的输出格式
- 💡 **观点提取**：支持从原文提取盈米基金观点

## 报告内容

生成的 Word 周报包含：
1. 监测结果综述
2. 盈米基金重点信息
3. 竞品要闻
4. 合作伙伴要闻
5. 行业要闻
6. 备注说明

## 安装

### 依赖安装

```bash
pip3 install openpyxl python-docx
```

### Skill 安装

```bash
# 方式1：使用 .skill 文件安装
claude skill install weekly-sentiment-report.skill

# 方式2：手动安装
cp -r weekly-sentiment-report ~/.claude/skills/
```

## 使用方法

### 基本用法

```bash
python3 scripts/generate_report.py \
  --data-file "/path/to/main_data.xlsx" \
  --official-file "/path/to/official_reports.xlsx" \
  --output "/path/to/report.docx" \
  --start-date "2026年01月04日" \
  --end-date "2026年01月12日"
```

### 参数说明

| 参数 | 说明 | 必需 |
|------|------|------|
| `--data-file` | 主数据 Excel 文件路径 | 是 |
| `--official-file` | 官方媒体报道 Excel 文件路径 | 否 |
| `--output` | 输出 Word 文件路径 | 是 |
| `--start-date` | 监测开始日期（格式：YYYY年MM月DD日） | 是 |
| `--end-date` | 监测结束日期（格式：YYYY年MM月DD日） | 是 |

## 目录结构

```
weekly-sentiment-report/
├── SKILL.md                          # Skill 主文件
├── assets/
│   └── template.docx                # Word 模板
├── scripts/
│   ├── generate_report.py           # 生成脚本
│   ├── config.json                  # 配置文件
│   └── enriched_summaries.json      # 提取的观点摘要
└── references/
    └── usage_guide.md               # 使用指南
```

## 配置说明

`scripts/config.json` 包含以下配置：

- **权威媒体列表**：24家权威媒体
- **转载网站列表**：需要排除的转载网站
- **公告关键词**：用于识别公告类内容
- **盈米关键词**：用于筛选盈米相关摘要
- **竞品/合作伙伴/行业工作表列表**：需要读取的工作表

## 数据格式

### 主数据文件

Excel 文件需包含以下工作表：
- 盈米相关：主品牌-盈米基金、关键人物、主产品-且慢等
- 竞品：E大、基金豆、翼支付、创必承等
- 合作伙伴：中航证券、粤开证券
- 银行券商：易方达、国联证券、中金财富等
- 行业监管：监管政策法规、基金处罚违规

每个工作表至少包含：
- 序号、监测专题名称、标题（支持 HYPERLINK）、发布时间、倾向性、来源网站
- 第 24 列（X 列）为摘要内容

### 官方媒体报道文件

包含以下列：
- 序号、媒体、发布时间、选题、标题、记者、署名、链接

## 输出格式

所有模块统一格式为：
```
序号、媒体、发布时间、新闻标题
摘要（如有盈米观点）
原文链接
```

## 去重逻辑

- **标题去重**：相同标题视为重复
- **摘要去重**：相同摘要视为重复
- **时间优先**：保留发布时间最早的版本
- **盈米优先**：竞品和合作伙伴中排除与盈米新闻重复的内容

## 许可证

MIT License

## 作者

盈米基金

## 更新日志

### v2.0 (2026-01-15)
- ✅ 增强去重逻辑（标题或摘要相同都去重）
- ✅ 添加原文观点提取功能
- ✅ 统一所有模块格式
- ✅ 优化筛选逻辑

### v1.0 (2026-01-15)
- 🎉 初始版本
- ✅ 基本的报告生成功能
- ✅ 权威媒体筛选
- ✅ 自动去重
