Данный репозиторий содержит скрипты, написанные в рамках рабочих задач и обучения. Ниже приведено краткое описание каждого из них.

### Скрипт №1: create_report_new_object

**Входные данные:** 3 файла с разрешением XLSX.

**Назначение:** Отбор недавно открывшихся объектов на основе анализа выполнения плана продаж.

**Функционал:** Cчитывает данные из Excel-файлов. Классифицирует объекты на две группы:
1. "Немного не хватило" — объекты, которые незначительно не выполнили план.
2. "Провальные" — объекты, выполнившие план менее чем на определённый процент.

Для объектов, открывшихся 3 месяца назад отбирает топ-3 по результатам выполнения плана

**Результат:** Excel-файл с двумя группами объектов (2 месяца и 21 день) и топ-3 объектами из файла за 3 месяца.

**Дата создания:** Сентябрь 2024.
<br/><br/><br/>
### Скрипт №2: script_1C_table

**Входные данные:** выгрузка из 1С с разрешением XLSX.

**Назначение:** Преобразование выгрузки из 1С с группировками в плоскую таблицу для удобства анализа.

**Функционал:** Автоматизирует процесс преобразования данных из 1С в плоскую таблицу, которые необходимы на регулярной основе.

**Особенности:** В целях конфиденциальности данные в примере заменены на случайные значения.

**Результат:** файл с разрешением XLSX с информацией из выгрузки 1С в виде плоской таблица

**Дата создания:** Август 2024.
<br/><br/><br/>
### Скрипт №3: create_report_total_and_company

**Входные данные:** Выгрузки в формате CSV, по 2 файла на компанию, таким образом кол-фо файлов — 2n

**Назначение:** Сравнение данных из двух баз 1С и создание отчётов.

**Функционал:** Анализирует выгрузки из 1С:УТ в формате CSV.
Результат: два файла:
1. Сводный отчёт по всем файлам—компаниям.
2. Отчёт по каждой компании отдельно.

**Особенности:** В целях конфиденциальности наименования столбцов удалены, а переменные переименованы.

**Дата создания:** Февраль 2025.
<br/><br/><br/>
### Скрипт №4: oop_python_training_script

**Назначение:** Обучение базовым принципам объектно-ориентированного программирования (ООП) в Python.

**Тема задание:** Создание класса-контейнера, который включает вызов двух других классов. При создании объекта класса-контейнера автоматически создаются объекты других классов.

**Функционал:** Расчет необходимого кол-ва рулонов обоев для комнаты с учетом окон/дверей, а также длины и ширины рулона обоев

**Особенности:** Один из нескольких скриптов, написанных в рамках обучения ООП.

**Дата создания:** Январь 2025.
