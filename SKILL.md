---
name: pbu-dashboard
description: Generate a comprehensive, interactive HTML recruitment dashboard from PBU Excel data. This skill should be used when the user provides an Excel file containing PBU recruitment data and wants a dashboard with KPI cards, funnel charts, department analysis, recruiter performance, and deadline alerts. It handles the full pipeline: Excel extraction → data generation → HTML dashboard → GitHub Pages deployment → permanent link.
agent_created: true
---

# PBU Recruitment Dashboard Generator

Generates a complete, interactive HTML recruitment dashboard from the PBU global recruitment Excel tracker. Produces a self-contained static site with embedded data and Chart.js visualizations, deployed to GitHub Pages for a permanent shareable link.

## When to Use

Trigger when the user:
- Provides a PBU recruitment Excel file (e.g., `2026 PBU全球招聘需求及周汇报监控表.xlsx`)
- Asks to "生成看板", "创建看板", "更新看板", "generate dashboard", or similar
- Wants to view PBU recruitment data in a dashboard format

## Quick Start — Full Pipeline

To generate the dashboard from an Excel file and deploy to GitHub Pages:

### Step 1: Generate Data and HTML

```bash
python3 scripts/generate.py <excel_path> <output_dir>
```

This produces two files in `output_dir`:
- `dd_v5.js` — Aggregated recruitment data (JSON)
- `index.html` — Complete dashboard with all visualizations

### Step 2: Deploy to GitHub Pages

```bash
# Set these variables first (token from user config):
# GITHUB_TOKEN="ghp_xxx"  USER="YourUsername"  REPO="pbu-recruitment-dashboard"

cd "$OUTDIR" && git init && git config user.email "dashboard@pb.com" && git config user.name "PBU Dashboard"
git remote add origin "https://${GITHUB_TOKEN}@github.com/${USER}/${REPO}.git"

# Create repo if not exists
curl -s -o /dev/null -X POST -H "Authorization: token ${GITHUB_TOKEN}" -H "Content-Type: application/json" \
  -d '{"name":"'"${REPO}"'","private":false,"has_pages":true}' \
  "https://api.github.com/user/repos"

git add -A && git commit -m "Update $(date +%Y-%m-%d)" && git branch -M main && git push -u origin main --force

# Enable GitHub Pages
curl -s -X POST -H "Authorization: token ${GITHUB_TOKEN}" -H "Accept: application/vnd.github+json" \
  "https://api.github.com/repos/${USER}/${REPO}/pages" \
  -d '{"source":{"branch":"main","path":"/"}}' > /dev/null
```

### Step 3: Return the Link

The permanent dashboard URL is:
```
https://rogerschiang.github.io/pbu-recruitment-dashboard/
```

**Important**: GitHub Pages takes 1-2 minutes for the first deployment. After pushing, wait and verify `HTTP 200` before returning the link.

## Dashboard Features

5 interactive tabs:

1. **全年招聘漏斗 & 明细** — Overall funnel chart, conversion rate analysis, detailed position table with filters (search, status, region, category). Positions sorted: 正常招聘 → 已发offer待入职 → 已入职 → 暂停招聘
2. **在招岗位** — In-progress positions funnel, department distribution chart, detail table with conversion rates
3. **部门维度** 🇬🇧🇩🇰... — Department CV-Offer comparison chart, stacked status distribution, detail table with conversion rates. Overseas regions show flag emojis
4. **负责人维度** — Recruiter performance cards + detail table with conversion rates
5. **超周期预警** — Deadline alert: positions overdue or due within 14 days

## Data Sources

The script reads from the `全球招聘数据监控` sheet only. Key columns:
- Col 0 (序号), 2 (岗位类别), 4 (三级组织), 5 (四级组织), 8 (职位名称), 14 (招聘负责人), 16 (岗位状态), 17 (区域)
- Col 35 (已招聘天数), 36 (标准周期), 37 (结束时间), 38 (超周期天数)
- Col 46-51 (推荐简历→Offer数), Col 52-57 (各阶段通过率)

## Department Logic

- If 三级组织 contains "欧洲区" → use 四级组织 as department
- If 四级组织 contains "跨综" → department = "跨综"
- Otherwise → department = 三级组织

## Important Rules

- **Status sorting**: Always sort tables 正常招聘 → 已发offer待入职 → 已入职 → 暂停招聘. Use `in` operator (NOT `||`) to avoid falsy 0 bug: `(a.岗位状态 in statusOrder ? statusOrder[a.岗位状态] : 9)`
- **Flag emojis**: Only show country flags in the 部门维度 tab. Other tabs use plain department names
- **Chart datalabels**: Only c-rate chart shows datalabels. All other charts use `datalabels:{display:false}`
- **Department order**: `PBU HRBP Dept. → 销售客服 → 交付运营 → 经营管理 → PBU技术部 → 四海捷运 → 跨综 → Japan → Australia → North America → Europe (NL→UK→DK→FR→BE→IT→HU→ES)`
- **Data cutoff**: Dashboard shows the Excel's current date; update `__DATE__` placeholder in template

## Files

- `scripts/generate.py` — Main data extraction and generation script
- `assets/template.html` — Dashboard HTML template with `__DATE__` placeholder
- No references directory needed — all logic is in the scripts