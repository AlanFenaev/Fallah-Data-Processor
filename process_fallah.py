import re
from pathlib import Path
from collections import defaultdict

import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.chart import ScatterChart, Reference, Series
from openpyxl.chart.marker import Marker
from openpyxl.styles import Font, Alignment

# ------- НАСТРОЙКИ -------
BASE_DIR = Path(r"C:\путь\до\Fallah")           # ← поменяйте на свой путь!
SUB_DIM  = "08_EXCEL_DIMENSIONAL__РАЗМЕРНЫЕ"
LOAD_PREFIX = "01_НАГРУЗКА_ПРОГИБ"

OUT_XLSX = BASE_DIR / "СВОД_результат.xlsx"
OUT_PNG_DIR = BASE_DIR / "Графики"              # сюда сохраним PNG по группам

# Коэффициенты
K_Q = 6.25
K_W = (1.0 / 1000.0 * 18315.0) / 6250.0
# -------------------------


# ---------- Парсер имени папки ----------
def parse_folder_name(name: str):
    """
    Возвращает (boundary_condition, load_type).
    Например 'FULL_..._SSCC_..._UNIFORM-PRESSURE_...' -> ('SSCC', 'UNIFORM-PRESSURE').
    Если что-то не найдено — вернёт None в соответствующем поле.
    """
    upper = name.upper()

    # Тип защемления: ищем отдельным «словом», чтобы не поймать случайно SSSS внутри SSSSSS
    bc_match = re.search(r'(?:^|[_\-\s])(SSSS|CCCC|SSCC)(?:$|[_\-\s])', upper)
    bc = bc_match.group(1) if bc_match else "НЕИЗВЕСТНО"

    # Тип нагрузки: POINT-FORCE или UNIFORM-PRESSURE (варианты через - или _)
    load_match = re.search(r'(POINT[\-_]FORCE|UNIFORM[\-_]PRESSURE)', upper)
    if load_match:
        load = load_match.group(1).replace("_", "-")
    else:
        load = "НЕИЗВЕСТНО"

    return bc, load


def find_load_file(folder: Path):
    for f in folder.iterdir():
        if f.is_file() and f.suffix.lower() in (".xlsx", ".xlsm") \
           and f.stem.upper().startswith(LOAD_PREFIX.upper()):
            return f
    return None


def read_load_deflection(xlsx: Path) -> pd.DataFrame:
    df = pd.read_excel(xlsx, usecols="E,F", header=0)
    df.columns = ["q", "w"]
    df = df.dropna().apply(pd.to_numeric, errors="coerce").dropna()
    return df


# ---------- 1. Собираем данные ----------
raw_data = {}   # {имя_папки: df}

print(f"Сканирую папку: {BASE_DIR}")
if not BASE_DIR.exists():
    raise SystemExit(f"Ошибка: папка не найдена: {BASE_DIR}")

for sub in sorted(BASE_DIR.iterdir()):
    if not sub.is_dir():
        continue
    if sub.name == OUT_PNG_DIR.name:
        continue

    dim_folder = sub / SUB_DIM
    if not dim_folder.is_dir():
        continue

    xlsx = find_load_file(dim_folder)
    if xlsx is None:
        print(f"[!] Не найден файл '{LOAD_PREFIX}*' в {dim_folder}")
        continue

    try:
        df = read_load_deflection(xlsx)
        if df.empty:
            print(f"[!] Файл {xlsx.name} пуст")
            continue
        df["q_star"] = df["q"] / K_Q
        df["w_star"] = (df["w"] / 1000.0 * 18315.0) / 6250.0
        raw_data[sub.name] = df
        bc, load = parse_folder_name(sub.name)
        print(f"[+] {sub.name[:60]}...  BC={bc}  Load={load}  ({len(df)} строк)")
    except Exception as e:
        print(f"[!] Ошибка при чтении {xlsx.name}: {e}")

if not raw_data:
    raise SystemExit("Не найдено ни одной папки с данными. Проверьте BASE_DIR и имена.")


# ---------- 2. Группируем по (защемление, нагрузка) ----------
groups = defaultdict(dict)   # {(bc, load): {имя_папки: df}}
for name, df in raw_data.items():
    bc, load = parse_folder_name(name)
    groups[(bc, load)][name] = df

print("\nПолучилось групп:")
for (bc, load), members in sorted(groups.items()):
    print(f"  • {bc} / {load}: {len(members)} кривых")


# ---------- 3. Пишем Excel: отдельный лист на каждую группу ----------
wb = Workbook()
# удаляем дефолтный лист
wb.remove(wb.active)

# Сводный лист со всеми данными — на всякий случай
ws_all = wb.create_sheet("Все данные")
ws_all.append(["Папка", "BC", "Нагрузка", "q (исх.)", "w (исх.)", "q*", "w*"])
for c in ws_all[1]:
    c.font = Font(bold=True)
for name, df in raw_data.items():
    bc, load = parse_folder_name(name)
    for _, rec in df.iterrows():
        ws_all.append([name, bc, load, float(rec["q"]), float(rec["w"]),
                       float(rec["q_star"]), float(rec["w_star"])])
for col, w in zip("ABCDEFG", (45, 10, 22, 12, 12, 12, 12)):
    ws_all.column_dimensions[col].width = w

# Отдельный лист на каждую группу — со своими формулами и диаграммой
for (bc, load), members in sorted(groups.items()):
    # Имя листа в Excel ограничено 31 символом
    sheet_title = f"{bc}_{load}"[:31]
    ws = wb.create_sheet(sheet_title)

    ws.append(["Папка", "q (исх.)", "w (исх.)", "q* = q/6.25", "w* = (w/1000*18315)/6250"])
    for c in ws[1]:
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center")

    row_ranges = {}
    r = 2
    for name, df in members.items():
        start = r
        for _, rec in df.iterrows():
            ws.cell(row=r, column=1, value=name)
            ws.cell(row=r, column=2, value=float(rec["q"]))
            ws.cell(row=r, column=3, value=float(rec["w"]))
            ws.cell(row=r, column=4, value=f"=B{r}/{K_Q}")
            ws.cell(row=r, column=5, value=f"=(C{r}/1000*18315)/6250")
            r += 1
        row_ranges[name] = (start, r - 1)

    for col, w in zip("ABCDE", (45, 14, 14, 16, 26)):
        ws.column_dimensions[col].width = w

    # Диаграмма для этой группы
    chart = ScatterChart()
    chart.title = f"q* vs w*  |  {bc} / {load}"
    chart.style = 13
    chart.x_axis.title = "w*"
    chart.y_axis.title = "q*"
    chart.x_axis.delete = False
    chart.y_axis.delete = False

    for name, (start, end) in row_ranges.items():
        xref = Reference(ws, min_col=5, min_row=start, max_row=end)
        yref = Reference(ws, min_col=4, min_row=start, max_row=end)
        s = Series(yref, xref, title=name[:30])
        s.marker = Marker(symbol="circle", size=6)
        s.graphicalProperties.line.noFill = True
        chart.series.append(s)

    ws.add_chart(chart, "G2")

wb.save(OUT_XLSX)
print(f"\n[=] Таблица сохранена: {OUT_XLSX}")


# ---------- 4. Отдельные PNG на каждую группу ----------
OUT_PNG_DIR.mkdir(exist_ok=True)

for (bc, load), members in sorted(groups.items()):
    plt.figure(figsize=(10, 7))
    for name, df in members.items():
        df_sorted = df.sort_values("w_star")
        # В легенду — укороченное имя: только то, что меняется между папками
        # (например, MVI1, MVI2, ...). Можно поменять под себя.
        short = re.sub(r"^FULL_\d+_\d+_", "", name)   # убираем "FULL_20261005_124824_"
        plt.plot(df_sorted["w_star"], df_sorted["q_star"],
                 marker="o", linestyle="-", label=short[:40])

    plt.xlabel("w* (безразмерный прогиб)")
    plt.ylabel("q* (безразмерное давление)")
    plt.title(f"q* от w*   |   {bc}   |   {load}")
    plt.grid(True, alpha=0.3)
    plt.legend(fontsize=7, loc="best")
    plt.tight_layout()

    fname = OUT_PNG_DIR / f"{bc}__{load}.png"
    plt.savefig(fname, dpi=150)
    plt.close()
    print(f"[=] Сохранён график: {fname}")

print("\nГотово.")
