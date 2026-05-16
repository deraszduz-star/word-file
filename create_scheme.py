"""
Скрипт для создания схемы взаимодействия участников строительного проекта.
Формат: горизонтальный 16:9, академический стиль, шрифт Liberation Serif (аналог Times New Roman).
Чистый академический дизайн с блоками, стрелками и подписями.
"""

import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
import matplotlib.font_manager as fm
import numpy as np

# --- Настройка шрифта ---
font_path = '/usr/share/fonts/liberation/LiberationSerif-Regular.ttf'
font_bold_path = '/usr/share/fonts/liberation/LiberationSerif-Bold.ttf'
fm.fontManager.addfont(font_path)
fm.fontManager.addfont(font_bold_path)
font_prop = fm.FontProperties(fname=font_path)
font_bold_prop = fm.FontProperties(fname=font_bold_path)
plt.rcParams['font.family'] = font_prop.get_name()

# --- Параметры фигуры (16:9) ---
fig_width = 16
fig_height = 9
fig, ax = plt.subplots(1, 1, figsize=(fig_width, fig_height), dpi=200)
ax.set_xlim(0, 16)
ax.set_ylim(0, 9)
ax.set_aspect('equal')
ax.axis('off')
fig.patch.set_facecolor('white')
ax.set_facecolor('white')

# --- Центральный блок ---
center_x, center_y = 8, 4.5
center_w, center_h = 3.4, 1.4

center_box = FancyBboxPatch(
    (center_x - center_w/2, center_y - center_h/2),
    center_w, center_h,
    boxstyle="round,pad=0.12",
    linewidth=2.5, edgecolor='black', facecolor='#f5f5f5'
)
ax.add_patch(center_box)
ax.text(center_x, center_y, 'СТРОИТЕЛЬНАЯ\nКОМПАНИЯ',
        ha='center', va='center', fontsize=12, fontproperties=font_bold_prop,
        linespacing=1.4)

# --- Внешние участники ---
# Позиции по кругу: сверху, лево-верх, лево-низ, низ, право-низ, право-верх
radius_x = 5.8
radius_y = 3.4

participants = [
    {
        'name': 'Заказчик',
        'angle': 90,
        'label': 'требования, цели, бюджет',
    },
    {
        'name': 'Проектная\nорганизация',
        'angle': 150,
        'label': 'проектная документация',
    },
    {
        'name': 'Субподрядчики',
        'angle': 210,
        'label': 'выполнение отдельных\nвидов работ',
    },
    {
        'name': 'Органы строительного\nконтроля',
        'angle': 270,
        'label': 'проверка качества\nи соответствия',
    },
    {
        'name': 'Поставщики и\nлогистические партнёры',
        'angle': 330,
        'label': 'материалы\nи оборудование',
    },
    {
        'name': 'Органы власти и\nразрешительные структуры',
        'angle': 30,
        'label': 'согласования и\nразрешительная документация',
    },
]

# --- Отрисовка участников ---
for p in participants:
    angle_rad = np.radians(p['angle'])
    px = center_x + radius_x * np.cos(angle_rad)
    py = center_y + radius_y * np.sin(angle_rad)

    # Размер блока участника
    box_w = 3.0
    box_h = 1.1

    # Блок участника
    part_box = FancyBboxPatch(
        (px - box_w/2, py - box_h/2),
        box_w, box_h,
        boxstyle="round,pad=0.1",
        linewidth=1.8, edgecolor='black', facecolor='white'
    )
    ax.add_patch(part_box)

    # Текст участника
    ax.text(px, py, p['name'],
            ha='center', va='center', fontsize=9.5, fontproperties=font_bold_prop,
            linespacing=1.2)

    # --- Стрелка от участника к центру ---
    # Вычисляем направление
    dx = center_x - px
    dy = center_y - py
    dist = np.sqrt(dx**2 + dy**2)

    # Нормализованное направление
    nx = dx / dist
    ny = dy / dist

    # Точка старта (от края блока участника)
    # Определяем точку пересечения с границей блока участника
    # Упрощённо: сдвигаемся от центра участника в сторону центра на половину блока
    start_offset = 0.7
    sx = px + nx * start_offset
    sy = py + ny * start_offset

    # Точка конца (к краю центрального блока)
    end_offset = 1.0
    ex = center_x - nx * end_offset
    ey = center_y - ny * end_offset

    arrow = FancyArrowPatch(
        (sx, sy), (ex, ey),
        arrowstyle='->,head_width=5,head_length=5',
        linewidth=1.2, color='black',
        connectionstyle='arc3,rad=0'
    )
    ax.add_patch(arrow)

    # --- Подпись связи (рядом со стрелкой) ---
    # Позиция подписи — смещена от середины стрелки перпендикулярно
    mid_x = (sx + ex) / 2
    mid_y = (sy + ey) / 2

    # Смещение подписи перпендикулярно стрелке
    perp_x = -ny * 0.4
    perp_y = nx * 0.4

    ax.text(mid_x + perp_x, mid_y + perp_y, p['label'],
            ha='center', va='center', fontsize=7.5, fontproperties=font_prop,
            style='italic', color='#333333', linespacing=1.1)

# --- Сохранение ---
output_path = '/projects/sandbox/word-file/scheme_participants.png'
plt.tight_layout(pad=0.5)
plt.savefig(output_path, dpi=200, bbox_inches='tight',
            facecolor='white', edgecolor='none')
plt.close()
print(f'Схема сохранена: {output_path}')
