# -*- coding: utf-8 -*-
"""
Power BI Python 視覺效果 - 熱區矩陣（連動 1、2、3 階鑽研）
需拖入「值」：原因第一階、原因第二階、原因第三階、營運機構名稱、件數
"""

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
matplotlib.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'SimHei', 'sans-serif']
matplotlib.rcParams['axes.unicode_minus'] = False

# 取得欄位（Power BI 可能傳入 tuple 欄名，統一轉字串比對）
def col_name(c):
    return c[0] if isinstance(c, (list, tuple)) and len(c) > 0 else str(c)

cols = list(dataset.columns)
col_names = [col_name(c) for c in cols]

cause_col_list = []
op_col = None
val_col = None
for i, name in enumerate(col_names):
    if '原因第三階' in name or '第三階' in name:
        cause_col_list.append((3, cols[i]))
    elif '原因第二階' in name or '第二階' in name:
        cause_col_list.append((2, cols[i]))
    elif '原因第一階' in name or '第一階' in name:
        cause_col_list.append((1, cols[i]))
    elif '營運' in name or '機構' in name:
        op_col = cols[i]
    elif '項次' in name or '件數' in name or 'count' in name.lower() or 'value' in name.lower():
        val_col = cols[i]

cause_col_list.sort(key=lambda x: x[0])
cause_cols_only = [x[1] for x in cause_col_list]
op_col = op_col if op_col is not None else (cols[1] if len(cols) > 1 else cols[0])
if val_col is None:
    dataset = dataset.copy()
    dataset['_cnt'] = 1
    val_col = '_cnt'

# 依篩選結果選擇階層
cause_col = None
for j in range(len(cause_cols_only) - 1, -1, -1):
    col = cause_cols_only[j]
    if col in dataset.columns and dataset[col].nunique() > 1:
        cause_col = col
        break
if cause_col is None and cause_cols_only:
    cause_col = cause_cols_only[-1]
elif cause_col is None and cols:
    cause_col = cols[0]

matrix = dataset.pivot_table(index=cause_col, columns=op_col, values=val_col, aggfunc='sum', fill_value=0)
matrix = matrix.apply(pd.to_numeric, errors='coerce').fillna(0)
if matrix.empty or matrix.size == 0:
    fig, ax = plt.subplots(figsize=(8, 4))
    ax.text(0.5, 0.5, '無資料可顯示\n請點選其他圖表或清除篩選', ha='center', va='center', fontsize=14)
    ax.axis('off')
else:
    matrix = matrix.reindex(matrix.sum(axis=1).sort_values(ascending=False).index)
    vmax = max(float(matrix.values.max()), 1) if matrix.size > 0 else 1
    vals = matrix.values
    rows = list(matrix.index)
    cols_list = list(matrix.columns)
    row_label = '一階肇因' if '第一階' in str(cause_col) else ('二階肇因' if '第二階' in str(cause_col) else '三階肇因')
    cell_text = [['(' + row_label + ') ↓ vs 機構 →'] + [str(c) for c in cols_list]]
    cell_colors = [['#181B1F'] * (len(cols_list) + 1)]
    cell_fg = [['#8E9BAE'] * (len(cols_list) + 1)]
    for i, r in enumerate(rows):
        row_vals, row_bg, row_fg = [str(r)], ['#181B1F'], ['#C8D4E3']
        for j in range(len(cols_list)):
            v = vals[i, j]
            intensity = (v / vmax) if vmax > 0 else 0
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
    fig, ax = plt.subplots(figsize=(12, max(6, len(rows) * 0.5 + 1)))
    ax.axis('off')
    fig.patch.set_facecolor('#111217')
    ax.set_facecolor('#111217')
    table = ax.table(cellText=cell_text, cellLoc='center', loc='center', colWidths=[0.4] + [0.12] * len(cols_list))
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
plt.show()
