# Обработчик данных Fallah

Обработка выходных данных решателя на python для сравнения со статьёй N. Fallah 
"A novel finite volume based formulation for the elasto-plastic
analysis of plates"

Скрипт собирает данные из папок названных `08_EXCEL_DIMENSIONAL__РАЗМЕРНЫЕ`, которые лежат в папке "Fallah" 
считает безразмерные коэффициенты q* и w* (заданные в кПа и мм - соответсвенно) и строит сводный график.

## Установка
1. Установите Python 3.9+
2. Склонируйте репозиторий:
   ```bash
   git clone https://github.com/ваш_логин/Fallah-Data-Processor.git
   cd Fallah-Data-Processor
