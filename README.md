B2B_IC_MAKS_autotest_UI

UI Autotest CASCO - B2B insurance company MAKS. 

Structure: 
1) MAKS_B2B.py – Main file.
2) _Данные_СК_МАКС.xlsx – Import. Personal data involved in calculation.
3) conf_folders_maks.py – Configuration. Output by folder.
4) Chromedriver – Selenium driver. You must download current version!
5) Расчет_B2B_МАКС – Export of received data.

Task: Automatically import flow of personal data of insurer and vehicle into B2B. Subsequent receipt of insurance prize and a full breakdown of calculation parameters. 

Purpose: Checking relevance of CASCO tariffs, tracking changes on the site. 

Implementation: It is executed in the following sequence: 
1) Import of personal data from the .xlsx file (_Данные_СК_МАКС.xlsx);
2) Authorization, filling out forms, making calculations;
3) Saving the HTML page (necessary to track changes over time);
4) Parsing of all fields;
5) Export of all possible parameters and cases of calculation in .xlsx with formatting (convenient for analysis by users);
6) Creation of screenshots of all calculations, it is necessary to track errors and verify the correctness of the input / output of information.

Conclusion: This script allows you to get all possible variations from only one data in 2 minutes, increasing the device performance and adding more data, the implementation process is possible in large volumes. 

For the script to work, you must have an account with working access in B2B SK MAKS (Login / Password).

-------------------------------------------------------------------------------------------------------------------------------------
-------------------------------------------------------------------------------------------------------------------------------------

UI Автотест КАСКО - B2B страховой компании МАКС.

Структура: 
1) MAKS_B2B.py — Главный файл.
2) _Данные_СК_МАКС.xlsx — Импорт. Персональные данные участвующие в расчете.
3) conf_folders_maks.py — Конфигурация. Вывода по папкам.
4) Chromedriver — Драйвер Selenium. Необходимо скачать актуальную версию!
5) Расчет_B2B_МАКС — Экспорт полученных данных.

Задача: В автоматическом режиме импортировать поток персональных данных страховщика и транспортного средства в B2B. Последующее получение страховой премии и полной разверстки параметров расчета.

Цель: Проверка актуальности тарифов КАСКО, отслеживание изменений на сайте.

Реализация: Выполнена в следующей последовательности:
1) Импорт персональных данных из файла .xlsx (_Данные_СК_МАКС.xlsx);
2) Авторизация, заполнение форм, произведение расчетов;
3) Сохранение HTML страницы (необходимо для отслеживания изменений с течением времени);
4) Парсинг всех полей;
5) Экспорт всех возможных параметров и случаев расчета в .xlsx с форматированием (удобно для анализа пользователями);
6) Создание скриншотов всех расчетов, необходимо для отслеживания ошибок и удостоверении корректности ввода/вывода информации.

Вывод: Данный скрипт позволяет за 2 минуты получить все возможные вариации только по одним данным, повысив производительность устройства и добавив больше данных, процесс реализации возможен в больших объемах.

Для работы скрипта необходимо иметь учетную запись с рабочим доступ в B2B СК МАКС (Логин/Пароль).
