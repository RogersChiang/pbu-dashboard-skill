#!/usr/bin/env python3
"""PBU Recruitment Dashboard Generator.
Usage: python3 generate.py <excel_path> <output_dir>
Generates dd_v5.js and index.html in output_dir, ready for GitHub Pages deploy."""

import sys, os, json, math, datetime, re

try:
    import pandas as pd
except ImportError:
    os.system(f"{sys.executable} -m pip install pandas openpyxl -q")
    import pandas as pd

def get_val(df, row, col):
    v = df.iloc[row, col]
    if pd.isna(v): return None
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime('%Y-%m-%d')
    if isinstance(v, (int, float)):
        if math.isnan(v): return None
        if v == int(v): return int(v)
        return round(v, 4)
    return str(v).strip()

def get_department(l3, l4):
    if '欧洲区' in (l3 or ''): return l4
    if '跨综' in (l4 or ''): return '跨综'
    return l3

def main(excel_path, output_dir):
    df = pd.read_excel(excel_path, sheet_name='全球招聘数据监控', header=None)
    DATA_START = 5

    positions_all = []
    for r in range(DATA_START, len(df)):
        seq = get_val(df, r, 0)
        if seq is None or not isinstance(seq, (int, float)): continue
        status = str(get_val(df, r, 16) or '')
        l3 = str(get_val(df, r, 4) or '')
        l4 = str(get_val(df, r, 5) or '')
        p = {
            '序号': int(seq), '岗位类别': str(get_val(df, r, 2) or ''),
            '三级组织': l3, '四级组织': l4, '部门': get_department(l3, l4),
            '职位名称': str(get_val(df, r, 8) or ''),
            '职族': str(get_val(df, r, 9) or ''),
            '需求职级': str(get_val(df, r, 10) or ''),
            '新增替补': str(get_val(df, r, 11) or ''),
            '需求编号': str(get_val(df, r, 13) or ''),
            '招聘负责人': str(get_val(df, r, 14) or ''),
            '岗位状态': status,
            '区域': str(get_val(df, r, 17) or '').replace('(必填)', ''),
            'base地': str(get_val(df, r, 18) or ''),
            '推荐简历数': get_val(df, r, 46) or 0,
            '简历通过数': get_val(df, r, 47) or 0,
            '一面通过数': get_val(df, r, 48) or 0,
            '二面通过数': get_val(df, r, 49) or 0,
            '终面通过数': get_val(df, r, 50) or 0,
            'Offer数': get_val(df, r, 51) or 0,
            '简历筛选通过率': get_val(df, r, 52),
            '一面通过率': get_val(df, r, 53),
            '二面通过率': get_val(df, r, 54),
            '终面通过率': get_val(df, r, 55),
            'Offer接受率': get_val(df, r, 56),
            '入职到岗率': get_val(df, r, 57),
            '已招聘天数': get_val(df, r, 35),
            '岗位标准招聘周期': get_val(df, r, 36),
            '招聘周期结束时间': get_val(df, r, 37),
            '超周期天数': get_val(df, r, 38),
        }
        positions_all.append(p)

    # positions_detail: exclude 提前启动
    positions_detail = [p for p in positions_all if p['岗位状态'] != '提前启动，暂未提单']
    # positions: exclude 暂停+提前
    positions = [p for p in positions_all if p['岗位状态'] not in ('暂停招聘', '提前启动，暂未提单')]
    # in_progress
    in_progress = [p for p in positions_all if p['岗位状态'] in ('正常招聘', '已发offer待入职')]
    # alert: 超周期>=0 OR -14<=超<0
    alert_positions = []
    for p in in_progress:
        ov = p.get('超周期天数')
        if ov is not None and (ov >= 0 or (ov < 0 and ov >= -14)):
            alert_positions.append(p)

    status_counts = {}
    for p in positions_all:
        st = p['岗位状态']; status_counts[st] = status_counts.get(st, 0) + 1

    total_real = len(positions)
    keys = ['推荐简历数', '简历通过数', '一面通过数', '二面通过数', '终面通过数', 'Offer数']
    fnl_total = {k: sum(p[k] for p in positions) for k in keys}
    fnl_ip = {k: sum(p[k] for p in in_progress) for k in keys}

    cv = fnl_total['推荐简历数']; rv = fnl_total['简历通过数']
    m2 = fnl_total['二面通过数']; fn = fnl_total['终面通过数']
    of = fnl_total['Offer数']
    rates = {
        '简历筛选通过率': round(rv / cv * 100, 1) if cv else 0,
        '一二面通过率': round(m2 / rv * 100, 1) if rv else 0,
        '终面通过率': round(fn / rv * 100, 1) if rv else 0,
        'Offer接受率': round(of / fn * 100, 1) if fn else 0,
    }

    ip_dept = {}
    for p in in_progress: d = p['部门']; ip_dept[d] = ip_dept.get(d, 0) + 1

    # --- Read 个人招聘看板 for per-recruiter data ---
    person_stats = {}
    # Manual name cross-reference: global-sheet name → person-sheet name
    name_xref = {'Shiyu Qi': '齐诗雨', 'Weiqi Zhang': 'Weiqi Zhang', 'Yijie He': 'Yijie', 'Yuqing Liu': 'Yuqing'}
    try:
        df_person = pd.read_excel(excel_path, sheet_name='个人招聘看板', header=None)
        for r in range(4, len(df_person)):
            name = get_val(df_person, r, 0)
            if not name or name == 'nan': continue
            name = str(name).strip()
            if not person_stats.get(name):
                person_stats[name] = {'推荐简历数': 0, '简历通过数': 0, '一面通过数': 0, '二面通过数': 0, '终面通过数': 0, 'Offer数': 0, '岗位数': 0, '已入职': 0, '已发offer待入职': 0, '招聘中': 0}
            person_stats[name]['推荐简历数'] += get_val(df_person, r, 12) or 0
            person_stats[name]['简历通过数'] += get_val(df_person, r, 13) or 0
            person_stats[name]['一面通过数'] += get_val(df_person, r, 14) or 0
            person_stats[name]['二面通过数'] += get_val(df_person, r, 15) or 0
            person_stats[name]['终面通过数'] += get_val(df_person, r, 16) or 0
            person_stats[name]['Offer数'] += get_val(df_person, r, 17) or 0

        # Merge position status counts: map global names → person names
        for p in positions_all:
            gname = p.get('招聘负责人', '')
            if not gname: continue
            # Resolve name: check xref first, then direct match
            ps_name = name_xref.get(gname, gname)
            # If not in person_stats, try substring match
            if ps_name not in person_stats:
                for pn in person_stats:
                    if gname.lower().replace(' ','') in pn.lower() or pn.lower() in gname.lower().replace(' ',''):
                        ps_name = pn
                        break
            if ps_name not in person_stats:
                person_stats[ps_name] = {'推荐简历数': 0, '简历通过数': 0, '一面通过数': 0, '二面通过数': 0, '终面通过数': 0, 'Offer数': 0, '岗位数': 0, '已入职': 0, '已发offer待入职': 0, '招聘中': 0}
            person_stats[ps_name]['岗位数'] += 1
            st = p['岗位状态']
            if st == '已入职': person_stats[ps_name]['已入职'] += 1
            elif st == '已发offer待入职': person_stats[ps_name]['已发offer待入职'] += 1
            elif st == '正常招聘': person_stats[ps_name]['招聘中'] += 1

        print(f"个人招聘看板: {len(person_stats)} recruiters extracted")
    except Exception as e:
        print(f"WARNING: 个人招聘看板 read failed: {e}")

    now = datetime.datetime.now()
    data_date = now.strftime('%Y-%m-%d')

    D = {
        'data_date': data_date, 'status_counts': status_counts,
        'total_real': total_real, 'positions': positions,
        'positions_all': positions_all, 'positions_detail': positions_detail,
        'in_progress': in_progress, 'alert_positions': alert_positions,
        'fnl_total': fnl_total, 'fnl_ip': fnl_ip, 'rates': rates, 'ip_dept': ip_dept,
        'person_data': person_stats,
    }

    os.makedirs(output_dir, exist_ok=True)

    # Write dd_v5.js
    with open(os.path.join(output_dir, 'dd_v5.js'), 'w') as f:
        f.write('var D = '); json.dump(D, f, ensure_ascii=False); f.write(';')
    print(f"dd_v5.js written ({len(json.dumps(D, ensure_ascii=False))} chars)")

    # Read template and generate index.html
    skill_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    template_path = os.path.join(skill_dir, 'assets', 'template.html')
    if not os.path.exists(template_path):
        template_path = os.path.join(skill_dir, 'references', 'template.html')
    if not os.path.exists(template_path):
        # Fallback: look in current script dir
        template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'assets', 'template.html')

    if os.path.exists(template_path):
        with open(template_path) as tf:
            html = tf.read()
        html = html.replace('__DATE__', data_date)
        with open(os.path.join(output_dir, 'index.html'), 'w') as f:
            f.write(html)
        print(f"index.html written ({len(html)} bytes)")
    else:
        print(f"WARNING: template.html not found at {template_path}, skipping HTML generation")

    print(f"\n{'='*60}")
    print(f"Output: {output_dir}/")
    print(f"Files: dd_v5.js, index.html")
    print(f"Positions: {len(positions_all)} total, {len(positions_detail)} display, {len(alert_positions)} alerts")
    print(f"Ready for GitHub Pages deploy.")

if __name__ == '__main__':
    if len(sys.argv) < 3:
        print("Usage: python3 generate.py <excel_path> <output_dir>")
        sys.exit(1)
    main(sys.argv[1], sys.argv[2])
