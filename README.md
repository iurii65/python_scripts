Данный репозиторий содержит скрипты, написанные в рамках рабочих задач и обучения. Ниже приведено краткое описание каждого из них.

Скрипт №1: best_worst_pharmacies

**Входные данные:** 3 файла с разрешением XLSX.

**Назначение:** Отбор недавно открывшихся объектов на основе анализа выполнения плана продаж.

**Функционал:** cчитывает данные из Excel-файлов.

Классифицирует объекты на две группы:
"Немного не хватило" — объекты, которые незначительно не выполнили план.
"Провальные" — объекты, выполнившие план менее чем на определённый процент.

Для объектов с историей в 3 месяца отбирает топ-3 по результатам выполнения плана

**Результат:** Excel-файл с двумя группами объектов (2 месяца и 21 день) и топ-3 объектами из файла за 3 месяца.

**Дата создания:** Сентябрь 2024.

Скрипт №2: script_1C_table
Назначение: Преобразование выгрузки из 1С с группировками в плоскую таблицу для удобства анализа.
Входные данные: выгрузка из 1С с разрешением XLSX.
Функционал:
>> Автоматизирует процесс преобразования данных из 1С в плоскую таблицу.
>> Решает задачу регулярной обработки данных в одном и том же формате.
Особенности: В целях конфиденциальности данные в примере заменены на случайные значения.
Результат: файл с разрешением XLSX с инфоормацией из выгрузки 1С в виде плоской таблица
Дата создания: Август 2024.

Скрипт №3: create_report_total_and_company
Назначение: Сравнение данных из двух баз 1С и создание отчётов.
Входные данные:
Функционал:
Анализирует заранее выгруженные файлы из 1С.
Результат: два файла с ра:
>> Сводный отчёт по всем файлам.
>> Отчёт по каждому файлу отдельно.
Особенности: В целях конфиденциальности наименования столбцов удалены, а переменные переименованы.
Дата создания: Февраль 2025.

Скрипт №4: oop_python_training_script
Назначение: Обучение базовым принципам объектно-ориентированного программирования (ООП) в Python.

Функционал: 
Создание класса-контейнера, который включает вызов двух других классов.
При создании объекта класса-контейнера автоматически создаются объекты других классов.

Особенности: Один из нескольких скриптов, написанных в рамках обучения ООП.

Дата создания: Январь 2025.
