from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.chart import ScatterChart, Reference, Series
from openpyxl.chart.marker import Marker
from openpyxl.styles import Font, Alignment

# ------- НАСТРОЙКИ -------
# Укажите путь к папке Fallah на вашем компьютере
# Для Windows: r"C:\Users\ВашеИмя\Documents\Fallah"
# Для macOS/Linux: "/Users/ВашеИмя/Documents/Fallah"
BASE_DIR = Path(r"C:\путь\до\Fallah")           

SUB_DIM  = "08_EXCEL_DIMENSIONAL__РАЗМЕРНЫЕ"
LOAD_PREFIX = "01_НАГРУЗКА_ПРОГИБ"              # Префикс имени файла
OUT_XLSX = BASE_DIR / "СВОД_результат.xlsx"     # Куда писать таблицу
OUT_PNG  = BASE_DIR / "СВОД_график.png"         # Картинка графика

# Коэффициенты из вашего задания
K_Q = 6.25
K_W = (1.0 / 1000.0 * 18315.0) / 6250.0
# q* = q / 6.25
# w* = (w/1000 * 18315) / 6250
# -------------------------

def find_load_file(folder: Path) -> Path | None:
    """Ищет .xlsx, имя которого начинается с LOAD_PREFIX."""
    for f in folder.iterdir():
        if f.is_file() and f.suffix.lower() in (".xlsx", ".xlsm") \
           and f.stem.upper().startswith(LOAD_PREFIX.upper()):
            return f
    return None

def read_load_deflection(xlsx: Path) -> pd.DataFrame:
    """Читает столбцы E и F, начиная со 2-й строки (первая - заголовки)."""
    # Читаем с header=0 (первая строка - заголовки)
    df = pd.read_excel(xlsx, usecols="E,F", header=0)
    
    # Переименовываем колонки для удобства
    df.columns = ["q", "w"]
    
    # Удаляем пустые строки и приводим к числам
    df = df.dropna().apply(pd.to_numeric, errors='coerce').dropna()
    
    return df

# ---------- 1. Собираем данные ----------
data = {}   # {имя_папки: df с колонками q, w, q*, w*}

print(f"Сканирую папку: {BASE_DIR}")

if not BASE_DIR.exists():
    raise SystemExit(f"Ошибка: Папка не найдена: {BASE_DIR}")

for sub in sorted(BASE_DIR.iterdir()):
    if not sub.is_dir():
        continue
        
    dim_folder = sub / SUB_DIM
    if not dim_folder.is_dir():
        print(f"[-] Пропуск {sub.name}: нет папки {SUB_DIM}")
        continue
        
    xlsx = find_load_file(dim_folder)
    if xlsx is None:
        print(f"[!] Не найден файл '{LOAD_PREFIX}*' в {dim_folder}")
        continue

    try:
        df = read_load_deflection(xlsx)
        if df.empty:
            print(f"[!] Файл {xlsx.name} пуст или данные не распознаны")
            continue
            
        df["q_star"] = df["q"] / K_Q
        df["w_star"] = (df["w"] / 1000.0 * 18315.0) / 6250.0
        data[sub.name] = df
        print(f"[+] {sub.name}: {len(df)} строк из {xlsx.name}")
    except Exception as e:
        print(f"[!] Ошибка при чтении {xlsx.name}: {e}")

if not data:
    raise SystemExit("Не найдено ни одной папки с данными. Проверьте BASE_DIR и имена.")

# ---------- 2. Пишем сводную таблицу в Excel с формулами ----------
wb = Workbook()
ws = wb.active
ws.title = "Данные"

# Заголовки
headers = ["Папка", "q (исх.)", "w (исх.)", "q* = q/6.25", "w* = (w/1000*18315)/6250"]
ws.append(headers)
for cell in ws[1]:
    cell.font = Font(bold=True)
    cell.alignment = Alignment(horizontal="center")

# Запоминаем границы строк по каждой папке (для диаграммы)
row_ranges = {}    # name -> (start_row, end_row)
r = 2

for name, df in data.items():
    start = r
    for _, rec in df.iterrows():
        ws.cell(row=r, column=1, value=name)
        ws.cell(row=r, column=2, value=float(rec["q"]))
        ws.cell(row=r, column=3, value=float(rec["w"]))
        # Формулы Excel
        ws.cell(row=r, column=4, value=f"=B{r}/{K_Q}")
        ws.cell(row=r, column=5, value=f"=(C{r}/1000*18315)/6250")
        r += 1
    row_ranges[name] = (start, r - 1)

# Автоширина столбцов
for col, w in zip("ABCDE", (45, 14, 14, 16, 26)):
    ws.column_dimensions[col].width = w

# ---------- 3. Строим диаграмму внутри Excel ----------
chart = ScatterChart()
chart.title = "q* vs w*"
chart.style = 13
chart.x_axis.title = "w*"
chart.y_axis.title = "q*"
chart.x_axis.delete = False
chart.y_axis.delete = False

for name, (start, end) in row_ranges.items():
    xref = Reference(ws, min_col=5, min_row=start, max_row=end)  # w*
    yref = Reference(ws, min_col=4, min_row=start, max_row=end)  # q*
    s = Series(yref, xref, title=name)
    s.marker = Marker(symbol="circle", size=6)
    s.graphicalProperties.line.noFill = True   # только маркеры
    chart.series.append(s)

ws.add_chart(chart, "G2")
wb.save(OUT_XLSX)
print(f"[=] Таблица сохранена: {OUT_XLSX}")

# ---------- 4. Строим график через matplotlib (PNG) ----------
plt.figure(figsize=(10, 7))
for name, df in data.items():
    # Сортируем по w*, чтобы линия не "прыгала"
    df_sorted = df.sort_values("w_star")
    plt.plot(df_sorted["w_star"], df_sorted["q_star"], marker="o", linestyle="-", label=name)

plt.xlabel("w* (безразмерный прогиб)")
plt.ylabel("q* (безразмерное давление)")
plt.title("График q* от w*")
plt.grid(True, alpha=0.3)
plt.legend(fontsize=8, loc='best')
plt.tight_layout()
plt.savefig(OUT_PNG, dpi=150)
print(f"[=] График сохранён: {OUT_PNG}")
# plt.show()  # Раскомментируйте, если хотите, чтобы график открылся сразу
