# -*- coding: utf-8 -*-
"""
Power BI 完整看板 - 單一 Python 視覺效果呈現整個畫面
需拖入「值」：發生時間、營運機構、事故事件分類、原因第一階、原因第二階、原因第三階、件數(或項次)
"""
import pandas as pd
import matplotlib
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.gridspec import GridSpec
matplotlib.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'SimHei', 'sans-serif']
matplotlib.rcParams['axes.unicode_minus'] = False

try:
    import pandas as pd
    df = dataset.copy()
    df.columns = [str(c[0]) if isinstance(c, (tuple, list)) and len(c) > 0 else str(c) for c in df.columns]
    
    # 辨識欄位
    cols = list(df.columns)
    def find_col(kw):
        for c in cols:
            if any(k in str(c) for k in kw):
                return c
        return cols[0] if cols else None
    
    time_col = find_col(['發生時間', '時間'])
    op_col = find_col(['營運', '機構'])
    cls_col = find_col(['事故事件分類', '分類'])
    l1_col = find_col(['原因第一階', '第一階'])
    l2_col = find_col(['原因第二階', '第二階'])
    l3_col = find_col(['原因第三階', '第三階'])
    val_col = find_col(['件數', '項次', 'count'])
    if val_col is None:
        df['_cnt'] = 1
        val_col = '_cnt'
    
    df[val_col] = pd.to_numeric(df[val_col], errors='coerce').fillna(1)
    if time_col and time_col in df.columns:
        df['年月'] = pd.to_datetime(df[time_col], errors='coerce').dt.strftime('%Y/%m')
        df = df[df['年月'] != 'NaT']
    if '年月' not in df.columns:
        df['年月'] = 'N/A'
    
    OP_MAP = {'TRA': '臺鐵', 'THSR': '高鐵', 'AFR': '林鐵', 'TSR': '糖鐵'}
    if op_col and op_col in df.columns:
        df['機構'] = df[op_col].map(lambda x: OP_MAP.get(str(x).upper(), str(x)))
    elif '機構' not in df.columns:
        df['機構'] = 'N/A'
    
    CLS_COLORS = {'重大行車事故': '#E02F44', '一般行車事故': '#FF9830', '異常事件': '#32CEE6'}
    OP_COLORS = {'臺鐵': '#42A5F5', '高鐵': '#FF7043', '林鐵': '#66BB6A', '糖鐵': '#D4E157'}
    
    fig = plt.figure(figsize=(16, 12))
    fig.patch.set_facecolor('#111217')
    gs = GridSpec(3, 3, figure=fig, hspace=0.35, wspace=0.3)
    
    # 1. 頂部指標卡 (文字)
    ax0 = fig.add_subplot(gs[0, :])
    ax0.axis('off')
    stats = {'total': df[val_col].sum(), 'major': 0, 'general': 0, 'anomaly': 0}
    if cls_col and cls_col in df.columns:
        stats['major'] = df[df[cls_col] == '重大行車事故'][val_col].sum()
        stats['general'] = df[df[cls_col] == '一般行車事故'][val_col].sum()
        stats['anomaly'] = df[df[cls_col] == '異常事件'][val_col].sum()
    ax0.text(0.12, 0.5, f"{stats['total']:,.0f}", fontsize=28, fontweight='bold', color='#FFF', ha='center')
    ax0.text(0.12, 0.15, '累計件數', fontsize=10, color='#6C7A92', ha='center')
    ax0.text(0.37, 0.5, f"{stats['major']:,.0f}", fontsize=28, fontweight='bold', color='#E02F44', ha='center')
    ax0.text(0.37, 0.15, '重大行車事故', fontsize=10, color='#6C7A92', ha='center')
    ax0.text(0.62, 0.5, f"{stats['general']:,.0f}", fontsize=28, fontweight='bold', color='#FF9830', ha='center')
    ax0.text(0.62, 0.15, '一般行車事故', fontsize=10, color='#6C7A92', ha='center')
    ax0.text(0.87, 0.5, f"{stats['anomaly']:,.0f}", fontsize=28, fontweight='bold', color='#32CEE6', ha='center')
    ax0.text(0.87, 0.15, '異常事件', fontsize=10, color='#6C7A92', ha='center')
    
    # 2. 趨勢圖
    ax1 = fig.add_subplot(gs[1, :2])
    ax1.set_facecolor('#181B1F')
    ax1.tick_params(colors='#8E9BAE')
    for spine in ax1.spines.values():
        spine.set_color('#2B2D37')
    if '年月' in df.columns and cls_col and cls_col in df.columns:
        trend = df.groupby(['年月', cls_col])[val_col].sum().unstack(fill_value=0)
        for cls in ['重大行車事故', '一般行車事故', '異常事件']:
            if cls in trend.columns:
                ax1.fill_between(range(len(trend)), trend[cls].values, alpha=0.3, color=CLS_COLORS.get(cls, '#888'))
                ax1.plot(trend[cls].values, color=CLS_COLORS.get(cls, '#888'), label=cls[:2])
        ax1.set_xticks(range(len(trend)))
        ax1.set_xticklabels(trend.index, rotation=45, ha='right', color='#8E9BAE')
        ax1.legend(loc='upper right', facecolor='#181B1F', labelcolor='#8E9BAE', fontsize=9)
    else:
        t = df.groupby('年月')[val_col].sum()
        if len(t) > 0:
            ax1.bar(range(len(t)), t.values, color='#32CEE6', alpha=0.7)
            ax1.set_xticks(range(len(t)))
            ax1.set_xticklabels(t.index, rotation=45, ha='right', color='#8E9BAE')
    ax1.set_title('事故發生趨勢', color='#8E9BAE', fontsize=12)
    
    # 3. 營運機構圓餅
    ax2 = fig.add_subplot(gs[1, 2])
    ax2.set_facecolor('#181B1F')
    if op_col or '機構' in df.columns:
        op_data = df.groupby('機構')[val_col].sum()
        colors = [OP_COLORS.get(str(k), '#888') for k in op_data.index]
        wedges, texts, autotexts = ax2.pie(op_data, labels=op_data.index, autopct='%1.0f%%', colors=colors, startangle=90)
        for t in texts + autotexts:
            t.set_color('#C8D4E3')
    ax2.set_title('營運機構佔比', color='#8E9BAE', fontsize=12)
    
    # 4. 事件分類圓餅
    ax3 = fig.add_subplot(gs[2, 0])
    ax3.set_facecolor('#181B1F')
    if cls_col:
        cls_data = df.groupby(cls_col)[val_col].sum()
        colors = [CLS_COLORS.get(str(k), '#888') for k in cls_data.index]
        wedges, texts, autotexts = ax3.pie(cls_data, labels=cls_data.index, autopct='%1.0f%%', colors=colors, startangle=90)
        for t in texts + autotexts:
            t.set_color('#C8D4E3')
    ax3.set_title('事件分類佔比', color='#8E9BAE', fontsize=12)
    
    # 5. 肇因橫條圖
    ax4 = fig.add_subplot(gs[2, 1])
    ax4.set_facecolor('#181B1F')
    ax4.tick_params(colors='#8E9BAE')
    if l1_col:
        cause_data = df.groupby(l1_col)[val_col].sum().sort_values(ascending=True)
        colors = ['#475569'] * len(cause_data)
        ax4.barh(range(len(cause_data)), cause_data.values, color=colors)
        ax4.set_yticks(range(len(cause_data)))
        ax4.set_yticklabels(cause_data.index, fontsize=9, color='#C8D4E3')
    ax4.set_title('肇因深度鑽研 (一階)', color='#8E9BAE', fontsize=12)
    
    # 6. 熱區矩陣表格
    ax5 = fig.add_subplot(gs[2, 2])
    ax5.axis('off')
    ax5.set_facecolor('#111217')
    if l1_col and (op_col or '機構' in df.columns):
        mat = df.pivot_table(index=l1_col, columns='機構', values=val_col, aggfunc='sum', fill_value=0)
        mat = mat.reindex(mat.sum(axis=1).sort_values(ascending=False).index)
        vmax = max(mat.values.max(), 1)
        cell_text = [['(一階肇因) ↓ vs 機構 →'] + list(mat.columns)]
        cell_colors = [['#181B1F'] * (len(mat.columns) + 1)]
        cell_fg = [['#8E9BAE'] * (len(mat.columns) + 1)]
        for i, r in enumerate(mat.index):
            row_vals, row_bg, row_fg = [str(r)], ['#181B1F'], ['#C8D4E3']
            for j, c in enumerate(mat.columns):
                v = mat.loc[r, c]
                intensity = v / vmax if vmax > 0 else 0
                if v == 0:
                    row_bg.append('#181B1F')
                    row_fg.append('#2B2D37')
                else:
                    row_bg.append((224/255, 47/255, 68/255, 0.1 + intensity * 0.9))
                    row_fg.append('#FFFFFF' if intensity > 0.4 else '#E8F0FE')
                row_vals.append(str(int(v)) if v > 0 else '-')
            cell_text.append(row_vals)
            cell_colors.append(row_bg)
            cell_fg.append(row_fg)
        table = ax5.table(cellText=cell_text, cellLoc='center', loc='center', colWidths=[0.35] + [0.1] * len(mat.columns))
        table.set_fontsize(9)
        table.scale(1, 1.8)
        for (row, col), cell in table.get_celld().items():
            if row < len(cell_colors) and col < len(cell_colors[row]):
                cell.set_facecolor(cell_colors[row][col])
                cell.set_text_props(color=cell_fg[row][col], fontsize=8)
            cell.set_edgecolor('#2B2D37')
    ax5.set_title('營運機構 vs 肇因 熱區矩陣', color='#8E9BAE', fontsize=12)
    
    plt.suptitle('🚆 鐵道事故事件監理中心', fontsize=16, color='#E8F0FE', y=0.98)
    plt.tight_layout(rect=[0, 0, 1, 0.96])
except Exception as e:
    fig, ax = plt.subplots(figsize=(8, 4))
    fig.patch.set_facecolor('#111217')
    ax.set_facecolor('#111217')
    ax.text(0.5, 0.5, '錯誤: ' + str(e)[:120], ha='center', va='center', fontsize=10, color='#E8F0FE', wrap=True)
    ax.axis('off')
plt.show()
