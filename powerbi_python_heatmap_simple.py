# -*- coding: utf-8 -*-
# 請完整複製此腳本貼入 Power BI Python 視覺效果
# 拖入「值」：原因第一階、營運機構名稱、件數
# 呈現：與網頁相同之「表格型」熱區矩陣（列=肇因、欄=機構、儲存格=件數+紅漸層）
import pandas as pd
import matplotlib
import matplotlib.pyplot as plt
matplotlib.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'SimHei', 'sans-serif']
matplotlib.rcParams['axes.unicode_minus'] = False

try:
    import pandas as pd
    df = dataset
    cols = list(df.columns)
    if len(cols) < 3:
        raise ValueError("請拖入至少 3 個欄位：原因第一階、營運機構名稱、件數")
    
    def name(c):
        return str(c[0]) if isinstance(c, (tuple, list)) and len(c) > 0 else str(c)
    
    cause_col = op_col = val_col = None
    for c in cols:
        n = name(c)
        if '原因' in n and ('第一階' in n or '第二階' in n or '第三階' in n):
            cause_col = c
        elif '營運' in n or '機構' in n:
            op_col = c
        elif '件數' in n or '項次' in n or 'count' in n.lower():
            val_col = c
    
    cause_col = cause_col or cols[0]
    op_col = op_col or cols[1]
    val_col = val_col or cols[2]
    
    matrix = df.pivot_table(index=cause_col, columns=op_col, values=val_col, aggfunc='sum', fill_value=0)
    matrix = matrix.apply(pd.to_numeric, errors='coerce').fillna(0)
    matrix = matrix.reindex(matrix.sum(axis=1).sort_values(ascending=False).index)
    
    vmax = max(float(matrix.values.max()), 1) if matrix.size > 0 else 1
    vals = matrix.values
    rows = list(matrix.index)
    cols_list = list(matrix.columns)
    
    # 建立表格（與網頁相同：表格式、列=肇因、欄=機構）
    cell_text = []
    cell_colors = []
    cell_fg = []
    
    # 表頭列
    header = ['(一階肇因) ↓ vs 機構 →'] + [str(c) for c in cols_list]
    cell_text.append(header)
    cell_colors.append(['#181B1F'] * len(header))
    cell_fg.append(['#8E9BAE'] * len(header))
    
    # 資料列（與網頁相同：每格獨立背景色）
    for i, r in enumerate(rows):
        row_vals = [str(r)]
        row_bg = ['#181B1F']
        row_fg = ['#C8D4E3']
        for j in range(len(cols_list)):
            v = vals[i, j]
            intensity = (v / vmax) if vmax > 0 else 0
            if v == 0:
                bg = '#181B1F'
                fg = '#2B2D37'
            else:
                alpha = 0.1 + intensity * 0.9
                bg = (224/255, 47/255, 68/255, alpha)
                fg = '#FFFFFF' if intensity > 0.4 else '#E8F0FE'
            row_vals.append(str(int(v)) if v > 0 else '-')
            row_bg.append(bg)
            row_fg.append(fg)
        cell_text.append(row_vals)
        cell_colors.append(row_bg)
        cell_fg.append(row_fg)
    
    fig, ax = plt.subplots(figsize=(12, max(6, len(rows) * 0.5 + 1)))
    ax.axis('off')
    fig.patch.set_facecolor('#111217')
    ax.set_facecolor('#111217')
    
    table = ax.table(cellText=cell_text, cellLoc='center', loc='center',
                     colWidths=[0.4] + [0.12] * len(cols_list))
    table.auto_set_font_size(False)
    table.set_fontsize(11)
    table.scale(1, 2.2)
    
    for (row, col), cell in table.get_celld().items():
        if row < len(cell_colors) and col < len(cell_colors[row]):
            cell.set_facecolor(cell_colors[row][col])
            cell.set_text_props(color=cell_fg[row][col])
        cell.set_edgecolor('#2B2D37')
        if row == 0:
            cell.set_text_props(fontweight='bold', color='#8E9BAE')
        elif col == 0:
            cell.set_text_props(ha='left', color='#C8D4E3')
    
    plt.tight_layout()
except Exception as e:
    fig, ax = plt.subplots(figsize=(8, 4))
    fig.patch.set_facecolor('#111217')
    ax.set_facecolor('#111217')
    ax.text(0.5, 0.5, '錯誤: ' + str(e)[:80], ha='center', va='center', fontsize=10, color='#E8F0FE')
    ax.axis('off')
plt.show()
