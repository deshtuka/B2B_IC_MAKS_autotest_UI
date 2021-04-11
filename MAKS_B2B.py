from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd
from collections import Counter
from xlsxwriter.utility import xl_rowcol_to_cell
from threading import Thread
from datetime import date
from lxml import html
import time
import os
import re
from conf_folders_maks import cf_trunk, cf_list_branches, cf_list_twig, cf_screen

# Параметры и премия
GL_good_1 = []
GL_good_2 = []
GL_good_3 = []
GL_prize = []


# Основное тело всего кода
def work(spisok, number):
    print(f'Поток #{number} - Запущен!')

    # Создание папок для скринов, согласно потоку => _Результаты/Скрины B2B/number
    creating_folders_screen(number)

    # Авторизация
    browser = authorization(user=spisok[0], password=spisok[1])

    # Переход к калькулятору КАСКО
    go_to_kasko_calculator(browser, user=spisok[0], password=spisok[1])

    # Ввод всех данных для расчета
    import_car_and_driver(browser, spisok, number)

    # Потоки - указываются данные под тип расчета
    type_of_calculation(browser, number)

    # Премия - Предварительный
    payment_1(browser, number)

    # Сканирование - Параметров страницы 1 [Быстрый расчет КАСКО]
    good_1 = scanner_param_1(browser, number)

    # Сканирование - Параметров страниц: 2 [ТС] и 3 [Параметры договора] - через selenium
    good_2 = scanner_param_2(browser)
    good_3 = scanner_param_3(browser)

    # Сканирование - Премий и Программ
    prize = scanner_prize(browser, number)

    # Закрытие браузера
    exit_code(browser)

    # Глобализация всех переменных
    global GL_good_1, GL_good_2, GL_good_3, GL_prize

    # Записываем все параметры по расчету
    GL_good_1 = GL_good_1 + [good_1]
    GL_good_2 = GL_good_2 + [good_2]
    GL_good_3 = GL_good_3 + [good_3]
    GL_prize = GL_prize + [prize]


# Переход к калькулятору КАСКО
def go_to_kasko_calculator(browser, user, password):
    try:
        print(f'Go to Калькулятор КАСКО')
        time.sleep(3)

        # Получить текущий url и убрать из него Логин и Пароль (Это особенность b2b)
        url_old = browser.current_url
        url_new = ''.join(url_old.split(f'{user}:{password}@'))

        browser.get(url_new)

        # Раскрыть список "Текущие страховые договоры"
        button = browser.find_element_by_xpath('//*[@id="INSURANCE_CONTRACT"]/i')
        time.sleep(0.5)
        button.click()

        # "Журнал ОСАГО 2.0 + КАСКО 2.0"
        button = browser.find_element_by_xpath('//*[@id="INSURANCE_CONTRACT_KBM_JOURNAL_KO_anchor"]')
        time.sleep(0.5)
        button.click()

        time.sleep(0.5)
        loading_spinner(browser)
        browser.find_element_by_tag_name('body').send_keys(Keys.CONTROL, 't')
        try:
            ale = browser.switch_to.alert
            ale.accept()
        except:
            pass

        # Переключение на фрейм, т.к. там находится КАСКО
        switch_frame(browser, xpath='//*[@id="btn_calc_kasko"]')

        # Калькулятор "КАСКО"
        try:
            browser.find_element_by_xpath(xpath_link).click()
        except:
            # Если НЕ кликается, то возможно появилась ошибка 501 - повторный клик после
            try:
                browser.find_element_by_class_name('ZebraDialog_Button_0').click()
            except:
                pass
            else:
                browser.find_element_by_xpath(xpath_link).click()

        time.sleep(1)

        # Переключаемся на вторую вкладку
        try:
            # Сохраняем Дискрептор текущего окна
            main_page = browser.current_window_handle
            # Ищем другой отличный дискрептор и изменяем
            for handle in browser.window_handles:
                if handle != main_page:
                    # Переключаем программу на найденную другу страницу
                    browser.switch_to.window(handle)
        except Exception as ex:
            print(f'Ошибка при переключение вкладок! Сообщение: {ex}')

        time.sleep(1)

        # "Осуществить новый расчет" - Клик
        try:
            xpath_button = '//*[@data-model="dataModelProgTempl_2"]//button'
            browser.find_element_by_xpath(xpath_button).click()
        except Exception as ex:
            browser.save_screenshot(f'{cf_screen}Ошибка - Осуществить новый расчет - Не кликается.png')
            print(f'Ошибка: "Осуществить новый расчет" - Не кликается! Сообщение: {ex}')
            try:
                browser.find_element_by_class_name('ZebraDialog_Button_0').click()
            except:
                pass
            else:
                browser.find_element_by_xpath(xpath_button).click()
    except:
        browser.save_screenshot(f'{cf_screen}Ошибка - не дошел до калькулятора КАСКО.png')


# Переключатель фреймов
def switch_frame(browser, xpath):
    seq = browser.find_elements_by_tag_name('iframe')
    check = False

    # Цикл переключения фреймов
    for index in range(len(seq)):
        # Закрыть старый фрейм
        browser.switch_to.default_content()
        # Берем каждый фрейм
        iframe = browser.find_elements_by_tag_name('iframe')[index]
        # Переключаемся на него
        browser.switch_to.frame(iframe)
        # Проверяем, есть ли наш элемент в фрейме
        try:
            browser.find_element_by_xpath(xpath)
        except Exception as ex:
            print(f'Поиск фрейма [{index}/{len(seq)}] -> Элемент - Не найден! Сообщение: {ex}')
        else:
            print(f'Поиск фрейма [{index}/{len(seq)}] -> Элемент - Найден!')
            check = True
            break

    return check


# Импорт всех данных
def import_car_and_driver(browser, spisok, number):
    print(f'--- Ввод данных ---')
    # browser.save_screenshot(f'_Ввод_данных #{number} - {cf_list_twig[number-1]}.png')
    time.sleep(0.5)

    # Регион преимущественного использования
    try:
        xpath_input = '//*[@id="regionListCompl"]'
        base_input = browser.find_element_by_xpath(xpath_input)
        base_input.clear()
        base_input.send_keys(spisok[3])
        base_input.send_keys(Keys.ENTER)
    except Exception as ex:
        browser.save_screenshot(f'{cf_screen}{number}/Ошибка - #{number} - Регион преимущественного использования!.png')
        print(f'Ошибка: [{number}] - Регион преимущественного использования!\n  Сообщение: {ex}')

    time.sleep(1)
    # ОБЯЗАТЕЛЬНОЕ ПОЛЕ - VIN или Гос.рег знак
    if spisok[36] != '' or str(spisok[36]).lower() != 'нет':
        # VIN
        id_checkbox = 'VinSearch'
        browser.find_element_by_id(id_checkbox).click()

        id_input = 'avtocodVIN'
        browser.find_element_by_id(id_input).send_keys(spisok[36])

    elif spisok[37] != '' or str(spisok[37]).lower() != 'нет':
        # Гос. номер
        id_checkbox = 'RegNumSearch'
        browser.find_element_by_id(id_checkbox).click()

        id_input = 'avtocodVIN'
        browser.find_element_by_id(id_input).send_keys(spisok[37])

    # -> Далее
    try:
        browser.find_element_by_id(id_checkbox).click()
        time.sleep(1)
        browser.find_element_by_xpath('//*[@data-bind="click:avtocodNextEvent"]').click()
    except Exception as ex:
        browser.save_screenshot(f'{cf_screen}{number}/Ошибка - далее у vin гос_номера.png')
        print(f'Ошибка - далее у vin гос_номера! Сообщение: {ex}')

    time.sleep(1)

    # Идет запрос в АВТОКОД
    loading_spinner(browser)

    # Если данные из АВТОКОДА не подтянулись - появляется окно "ПРЕДУПРЕЖДЕНИЕ"
    try:
        error_text = browser.find_element_by_xpath('//*[@class="ZebraDialog"]//strong').get_attribute('innerText')
        browser.find_element_by_xpath('//*[@class="ZebraDialog_Button_0"]').click()
        print(f'Данные не прошли в АВТОКОД! Сообщение: {error_text}')
    except Exception as ex:
        print(f'Ошибка: окно "ПРЕДУПРЕЖДЕНИЕ" - Не кликнулось!\n  Сообщение: {ex}')
    else:
        # Очистить поле и кликнуть Далее
        try:
            error_input = browser.find_element_by_id(id_input)
            while error_input.get_attribute('value') != '':
                ActionChains(browser).send_keys_to_element(error_input, Keys.BACKSPACE).perform()
            browser.find_element_by_id(id_checkbox).click()
            browser.find_element_by_xpath('//*[@data-bind="click:avtocodNextEvent"]').click()
        except Exception as ex:
            browser.save_screenshot(f'{cf_screen}{number}/Ошибка - При очистке поля VIN или Гос.номер.png')
            print(f'Ошибка: При очистке поля VIN/Гос.номер! Сообщение: {ex}')

    # Проверка на ввод Марки ТС...
    block = check_block(browser, xpath='//*[@id="markList"]')

    if block == 'block':
        try:
            list_brand = browser.find_elements_by_xpath('//*[@id="ul_markList"]//a')
            check = 0
            list_name_brand = []
            for brand in list_brand:
                name = brand.get_attribute('innerText')
                list_name_brand.append(name)
                if str(name).lower() == str(spisok[27]).lower():
                    print(f'Марка ТС [{name}] - Найдена!')
                    brand.click()
                    break
                check += 1
                if check == len(list_brand):
                    browser.save_screenshot(f'{cf_screen}Марка ТС - Не найдена.png')
                    print(f'Марка ТС [{spisok[27]}] - НЕ НАЙДЕНА В БАЗЕ!\nСписок МАРОК: {list_name_brand}\n')

        except Exception as ex:
            print(f'Ошибка: Марка ТС [{str(spisok[27]).lower()}] не найдена!\n  Сообщение: {ex}')

        else:
            block = check_block(browser, xpath='//*[@id="modelList"]')

            if block == 'block':
                # Модель ТС
                try:
                    list_model = browser.find_elements_by_xpath('//*[@id="ul_modelList"]//a')
                    check = 0
                    list_name_model = []
                    for model in list_model:
                        name = model.get_attribute('innerText')
                        list_name_model.append(name)
                        if str(name).lower() == str(spisok[28]).lower():
                            print(f'Модель ТС [{name}] - Найдена!')
                            model.click()
                            break
                        check += 1
                        if check == len(list_model):
                            browser.save_screenshot(f'{cf_screen}Модель ТС - Не найдена.png')
                            print(f'Модель ТС [{spisok[28]}] - НЕ НАЙДЕНА В БАЗЕ!\n'
                                  f'Список МОДЕЛЕЙ {spisok[27]}: {list_name_brand}\n')

                except Exception as ex:
                    print(f'Ошибка: Модель ТС [{spisok[28]}] не найдена!\n  Сообщение: {ex}')

                else:
                    block = check_block(browser, xpath='//*[@id="modelYearIssue"]')

                    if block == 'block':
                        # Год выпуска
                        try:
                            list_year = browser.find_elements_by_xpath('//*[@data-bind="foreach: yearIssueList"]//a')
                            check = 0
                            list_name_year = []
                            for year in list_year:
                                name = year.get_attribute('innerText')
                                list_name_year.append(name)
                                if str(name).lower() == str(spisok[29]).lower():
                                    print(f'Год ТС [{name}] - Найден!')
                                    year.click()
                                    break
                                check += 1
                                if check == len(list_year):
                                    browser.save_screenshot(f'{cf_screen}Год ТС - Не найден.png')
                                    print(f'Год ТС [{spisok[29]}] - НЕ НАЙДЕН В БАЗЕ!\n'
                                          f'Доступный список у [{spisok[27]} {spisok[28]}]: {list_name_year}\n')

                        except Exception as ex:
                            print(f'Ошибка: Год ТС [{spisok[29]}] не найден!\n  Сообщение: {ex}')

    browser.save_screenshot(f'Скрин после - марки-модели-года ТС.png')

    time.sleep(1)

    # Стоимость ТС
    xpath_input = '//*[@id="amountSelectVal_1"]'
    base_input = browser.find_element_by_xpath(xpath_input)

    print(f'Предложенная "Стоимость ТС" = {base_input.get_attribute("value")}')

    while base_input.get_attribute('value') != '':
        ActionChains(browser).send_keys_to_element(base_input, Keys.BACKSPACE).perform()
    ActionChains(browser).send_keys_to_element(base_input, spisok[35]).perform()

    # -> Далее
    browser.find_element_by_xpath('//*[@data-bind="click: modelAmountSelectEvent"]').click()

    # Водители / Мультидрайв
    if number != 5:
        # Водитель(я)
        try:
            # Проверка - сколько всего водителей
            n = 0
            try:
                if spisok[44].lower() == 'да':
                    if spisok[46] != '' and spisok[47] != '':
                        if spisok[5] != spisok[46] and spisok[6] != spisok[47]:
                            # Значит ДВА водителя
                            n = 2
                    else:
                        n = 0
                elif spisok[44].lower() == 'нет':
                    n = 41
            except Exception as ex:
                print(f'Ошибка: Не верно указан водитель(-я) в Excel файле!\n  Сообщение: {ex}')

            # Функция добавления Водителя(-ей)
            def add_driver(browser_def, excel_list, num_thread, driver_ratio):
                try:
                    # Добавить водителя - Клик
                    browser_def.find_element_by_xpath('//*[@data-bind="click: addDriver"]').click()
                    time.sleep(0.5)
                    # Список Инпутов
                    list_input = browser_def.find_elements_by_xpath('//*[@id="ageExperienceSelect"]//input')
                    # Водитель - Дата рождения
                    list_input[0].send_keys(str(excel_list[8 + driver_ratio].strftime("%d.%m.%Y")))
                    # Водитель - Стаж
                    list_input[1].send_keys(str(excel_list[14 + driver_ratio].strftime("%d.%m.%Y")))
                except Exception as ex_error:
                    browser_def.save_screenshot(
                        f'{cf_screen}{num_thread}/Ошибка - #{num_thread} - Добавление водителя.png')
                    print(f'Ошибка: [#{num_thread}] - Добавление водителя!\n  Сообщение: {ex_error}')
                else:
                    # Скрин водителя(-ей) - проверка наличия файла, если файл есть, то делается скрин второго водителя
                    screen_driver_1 = f'{cf_screen}{num_thread}/{cf_list_twig[num_thread - 1]} - Данные водителя #1.png'
                    screen_driver_2 = f'{cf_screen}{num_thread}/{cf_list_twig[num_thread - 1]} - Данные водителя #2.png'

                    if os.path.isfile(screen_driver_1):
                        # Если скрин #1 - Существует! - Сделать скрин #2
                        browser_def.save_screenshot(screen_driver_2)
                    else:
                        # Иначе сделать скрин #1
                        browser_def.save_screenshot(screen_driver_1)

                    # -> Далее [Добавить водителя]
                    browser_def.find_element_by_xpath('//*[@data-bind="click: saveDriver"]').click()

            # Первый водитель - Страховщик
            if n == 0:
                add_driver(browser, spisok, number, driver_ratio=0)
            # Второй водитель - Водитель
            elif n == 41:
                add_driver(browser, spisok, number, driver_ratio=41)
            # ДВА водителя - Страховщик + Водитель
            elif n == 2:
                add_driver(browser, spisok, number, driver_ratio=0)
                add_driver(browser, spisok, number, driver_ratio=41)

        except Exception as ex:
            print(f'Ошибка: Водитель!\n  Сообщение: {ex}')
    else:
        # Мультидрайв
        try:
            time.sleep(1)
            print(f'--- [{cf_list_twig[number - 1]}] -> Ввод данных ---')
            xpath_checkbox = '//*[@id="face_age"]//tbody/tr[4]//input'
            browser.find_element_by_xpath(xpath_checkbox).click()
        except Exception as ex:
            browser.save_screenshot(f'{cf_screen}{number}/Ошибка - {cf_list_twig[number - 1]} - Не кликается!')
            print(f'Ошибка: [#{number}] - Мультидрайв - Не кликается!\n  Сообщение: {ex}')
        else:
            browser.save_screenshot(f'{cf_screen}{number}/{cf_list_twig[number - 1]} - Данные водителя.png')

    # -> Далее [Водители / Мультидрайв]
    browser.find_element_by_xpath('//*[@data-bind="click: ageNextEvent"]').click()

    # Появляется СПИНЕР --- Проверка на отображение
    loading_spinner(browser)


# Потоки - указываются данные под тип расчета
def type_of_calculation(browser, number):
    if number != 1 and number != 5:
        print(f'--- [{cf_list_twig[number - 1]}] -> Ввод данных ---')

        # ТС - Кредитное
        if number == 2:
            xpath_button = '//span[contains(text(), "Агент-банк")]/parent::li'
            browser.find_element_by_xpath(xpath_button).click()

            # Агент - банк -> True
            xpath_checkbox = '//*[@data-bind="checked: isPartBank"]'
            browser.find_element_by_xpath(xpath_checkbox).click()
            time.sleep(0.5)

            # Название банка
            xpath_input = '//*[@id="bankPart1"]'
            browser.find_element_by_xpath(xpath_input).send_keys('АЛЬФА-БАНК')

            # Прогрузка
            loading_spinner(browser)

            # Выбор первого из выпадающего списка
            browser.find_element_by_xpath('//*[@id="ui-id-11"]/li[1]').click()

            # Далее - Клик
            browser.find_element_by_xpath('//*[@data-bind="click: bankPartNextEvent"]').click()

        # ТС - Переход из другой СК
        elif number == 3:
            try:
                xpath_button = '//span[contains(text(), "Безубыточный переход из другой СК")]/parent::li'
                browser.find_element_by_xpath(xpath_button).click()
                time.sleep(1)

                xpath_check = '//span[contains(text(), "Безубыточный переход из другой СК")]/../span[2]'
                check_text = browser.find_element_by_xpath(xpath_check).get_attribute('innerText')

                # ТС Новое - Переход запрещен!
                if str(check_text).lower() == 'нет':
                    print(f'Переход из другой СК невозможен - ТС новое!')
                    browser.find_element_by_class_name('ZebraDialog_Button_0').click()

            # Ошибка [Переход из другой СК невозможен - ТС новое]
            except Exception as ex:
                browser.save_screenshot(f'{cf_screen}{number}/Переход из другой СК - Не кликается.png')
                print(f'Ошибка: [ТС - Переход из другой СК] - Не кликается!\n  Сообщение: {ex}')

        # ТС - Франшиза
        elif number == 4:
            xpath_button = '//span[contains(text(), "Франшиза")]/parent::li'
            browser.find_element_by_xpath(xpath_button).click()

            # Применить безусловную франшизу - Клик
            xpath_checkbox = '//*[@id="deductible"]//input[@type="checkbox"]'
            browser.find_element_by_xpath(xpath_checkbox).click()
            time.sleep(1)

            # 20 000 - Клик
            try:
                xpath_label = '//*[@id="deductible"]//a[contains(text(), "20 000")]'
                browser.find_element_by_xpath(xpath_label).click()
            except Exception as ex:
                browser.save_screenshot(f'{cf_screen}Ошибка - Франшиза в размере = 20 000 - Не кликнулась.png')
                print(f'Ошибка: Франшиза в размере = 20 000 - Не кликнулась!\n  Сообщение: {ex}')
            time.sleep(0.5)

            # Далее - Клик
            xpath_button = '//*[@data-bind="click: deductibleNextEvent"]'
            browser.find_element_by_xpath(xpath_button).click()

        # Рассчитать премию - Клик
        try:
            browser.find_element_by_xpath('//input[@value="Рассчитать премию"]').click()
        except Exception as ex:
            browser.save_screenshot(
                f'{cf_screen}Ошибка - Рассчитать премию {cf_list_twig[number - 1]} - Не кликается.png')
            print(
                f'Ошибка: Рассчитать премию [#{number} = {cf_list_twig[number - 1]}] - Не кликается!\n Сообщение: {ex}')
        else:
            loading_spinner(browser)


# Проверка на прогрузку СПИНЕРА - расчет
def loading_spinner(browser):
    xpath_js = '//*[@aria-describedby="splash-dialog"]'
    try:
        while True:
            container = browser.find_element_by_xpath(xpath_js)
            block = browser.execute_script('return arguments[0].style.display;', container)
            print(f'Загрузка... -> block = {block}')
            if block == 'none':
                break
            time.sleep(0.7)
    except Exception as ex:
        print(f'Ошибка сприннер: {ex}')


# Проверка отображения блока - нужно для Марки/Модели/Года ТС
def check_block(browser, xpath):
    block = ''
    browser.implicitly_wait(3)
    xpath_window = xpath
    try:
        check = 0
        while True:
            container = browser.find_element_by_xpath(xpath_window)
            block = browser.execute_script("return arguments[0].style.display;", container)
            if block == 'block' or check == 10:
                break
            time.sleep(1)
            check += 1
    except Exception as ex:
        print(f'Ошибка: при определение отображения блока!\n  Сообщение: {ex}')
    finally:
        browser.implicitly_wait(15)
        return block


# Получение премии
def payment_1(browser, number):
    print(f'=== Получение премии ===')
    try:
        xpath_name_1 = '//*[@id="frmProgram"]//tr[3]//a[2]'
        xpath_value_1 = '//*[@id="frmProgram"]//tr[3]//td[2]'

        name_1 = browser.find_element_by_xpath(xpath_name_1).get_attribute('innerText')
        value_1 = browser.find_element_by_xpath(xpath_value_1).get_attribute('innerText')

    except Exception as ex:
        browser.save_screenshot(f'{cf_screen}{number}/Ошибка - {cf_list_twig[number - 1]} - Премия.png')
        print(f'Ошибка премии! Сообщение: {ex}')

    else:
        print(f'\n@-@ Премия @-@\n{name_1} = {value_1}')
        browser.save_screenshot(f'{cf_screen}{number}/{cf_list_twig[number - 1]} - Премия.png')

    if number == 1:
        save_page(browser, 1)


# Сканирование параметров
def scanner_param_1(browser, number):
    xpath_head = [
        '//*[@id="contract_ts_info"]',
        '//*[@id="contract_condition"]',
        '//*[@id="contract_ways_reduce_cost"]',
        '//*[@id="contract_additional"]',
        '//*[@id="send_sms_email"]',
        '//*[@id="add_prams"]'
    ]

    id_body = [
        'selectableListTC',
        'selectableListCondition',
        'selectableListReduce',
        'selectableListAdditional'
    ]

    website = browser.page_source.encode('utf-8')
    tree = html.fromstring(website)

    # Объявление переменных
    good, txt_head, txt_key, txt_value = ('' for i in range(4))

    try:
        txt_tab_1 = browser.find_element_by_xpath('//*[@id="tabs-contract"]/ul/li[2]/a').get_attribute('innerText')
    except Exception as ex:
        browser.save_screenshot(f'{cf_screen}Ошибка - Парсинг - Вкладка #1 - Быстрый расчет КАСКО - Не найдена.png')
        print(f'Ошибка при парсинге: Вкладка #1 - "Быстрый расчет КАСКО" - Не найдена!\n  Сообщение: {ex}')
    else:
        print(f'{txt_tab_1}')
        good = [[txt_tab_1, cf_list_twig[number - 1]]]

        # Парсинг
        for number in range(0, len(xpath_head)):

            # Заголовок
            head_object = tree.xpath(f'{xpath_head[number]}')
            for head in head_object:
                txt_head = head.text
            print(f'\n{txt_head}')
            good += [[txt_head, '']]

            # Параметры
            for n in range(1, 16):
                try:
                    param_key = tree.xpath(f'//*[@id="{id_body[number]}"]/li[{n}]/span[1]')
                    param_value = tree.xpath(f'//*[@id="{id_body[number]}"]/li[{n}]/span[2]')
                    for key in param_key:
                        # Создание уникальности для панды меняем ТЕКСТ
                        txt_key = f'{key.text}_1/{n}'

                    for value in param_value:
                        txt_value = value.text
                        txt_value = '' if txt_value is None else txt_value

                    if txt_key != '':
                        print(f'{txt_key}: {txt_value}')
                        good += [[txt_key, txt_value]]

                    txt_key, txt_value = ('' for i in range(2))
                except:
                    pass

    return good


# Сканирование параметров страница 2 [ТС] - через selenium
def scanner_param_2(browser):
    xpath_text_free = [
        '//*[@id="tabs-contract"]/ul/li[3]/a',
        '//*[@id="formTC"]/fieldset[1]/legend',
        '//*[@id="formTC"]/fieldset[1]/table/tbody/tr[2]/td[1]',
        'VIN Отсутствует',
        '//*[@id="formTC"]/fieldset[1]/table/tbody/tr[3]/td[1]',
        '//*[@id="formTC"]/fieldset[1]/table/tbody/tr[4]/td[1]',
        '//*[@id="formTC"]/fieldset[1]/table/tbody/tr[2]/td[3]',
        '//*[@id="formTC"]/fieldset[1]/table/tbody/tr[3]/td[3]',
        '//*[@id="formTC"]/fieldset[1]/table/tbody/tr[4]/td[3]',
        '//*[@id="formTC"]/fieldset[1]/table/tbody/tr[2]/td[5]',
        '//*[@id="formTC"]/fieldset[1]/table/tbody/tr[3]/td[5]',
        '//*[@id="formTC"]/fieldset[1]/table/tbody/tr[4]/td[5]',
        '//*[@id="formTC"]/fieldset[2]/legend',
        '//*[@id="docTSType1"]/option[1]',
        '//*[@id="formTC"]/fieldset[3]/legend',
        'Тип',
        '//*[@id="formTC"]/fieldset[3]/table/tbody/tr[2]/td[1]',
        '//*[@id="formTC"]/fieldset[4]/legend',
        '//*[@id="formTC"]/fieldset[4]/table[1]/tbody/tr/td[1]/span',
        'Тип',
        '//*[@id="formTC"]/fieldset[4]/table[2]/tbody/tr[1]/td[1]',
        '//*[@id="formTC"]/fieldset[6]/legend',
        '//*[@id="formTC"]/fieldset[6]/table/tbody/tr/td[1]',
        '//*[@id="formTC"]/fieldset[6]/table/tbody/tr/td[2]',
        '//*[@id="formTC"]/fieldset[7]/legend',
        '//*[@id="formTC"]/fieldset[7]/table/thead/tr/th[1]',
        '//*[@id="formTC"]/fieldset[7]/table/thead/tr/th[2]',
        '//*[@id="formTC"]/fieldset[7]/table/thead/tr/th[3]',
        '//*[@id="formTC"]/fieldset[7]/table/thead/tr/th[4]',
        '//*[@id="formTC"]/fieldset[7]/table/thead/tr/th[5]',
        '//*[@id="formTC"]/fieldset[8]/legend',
        'Тип',
        '//*[@id="formTC"]/fieldset[9]/legend',
        '//*[@id="formTC"]/fieldset[9]/table/tbody/tr[1]/td[1]/span',
        '//*[@id="formTC"]/fieldset[9]/table/tbody/tr[2]/td[1]/span',
        '//*[@id="formTC"]/fieldset[9]/table/tbody/tr[1]/td[2]/span',
        'С автозапуском',
        '//*[@id="formTC"]/fieldset[9]/table/tbody/tr[3]/td[1]/span',
        '//*[@id="formTC"]/fieldset[9]/table/tbody/tr[1]/td[3]/span',
        '//*[@id="formTC"]/fieldset[9]/table/tbody/tr[2]/td[3]/span',
        '//*[@id="formTC"]/fieldset[9]/table/tbody/tr[4]/td[1]/span',
        '//*[@id="formTC"]/fieldset[9]/table/tbody/tr[4]/td[2]/span'
    ]

    xpath_input_free = [
        '',
        '//*[@id="formTC"]/fieldset[1]/input',
        '//*[@id="VINNum1"]',
        '//*[@id="formTC"]/fieldset[1]/table/tbody/tr[1]/td[2]/input',
        '//*[@id="regNum1"]',
        '//*[@id="transportBuyDate1"]',
        '//*[@id="formTC"]/fieldset[1]/table/tbody/tr[2]/td[4]/div/input',
        '//*[@id="formTC"]/fieldset[1]/table/tbody/tr[3]/td[4]/div/input',
        '//*[@id="formTC"]/fieldset[1]/table/tbody/tr[4]/td[4]/div/input',
        '//*[@id="carModelModification"]',
        '//*[@id="formTC"]/fieldset[1]/table/tbody/tr[3]/td[6]/input',
        '//*[@id="ptson"]',
        '',
        '//*[@id="documentTCList1"]/li/input',
        '',
        '//*[@id="formTC"]/fieldset[3]/table/tbody/tr[1]/td[1]/input[1]',
        '//*[@id="docOwnershipConditions3"]/option[1]',
        '',
        '//*[@id="formTC"]/fieldset[4]/table[1]/tbody/tr/td[1]/input',
        '//*[@id="formTC"]/fieldset[4]/table[1]/tbody/tr/td[2]/div/input',
        '//*[@id="docOwnershipConditions1"]/option[1]',
        '',
        '//*[@id="formTC"]/fieldset[6]/table/tbody/tr/td[1]/input',
        '//*[@id="formTC"]/fieldset[6]/table/tbody/tr/td[2]/input',
        '',
        '//*[@id="formTC"]/fieldset[7]/table/tbody/tr/td[1]/input',
        '//*[@id="formTC"]/fieldset[7]/table/tbody/tr/td[2]/input',
        '//*[@id="formTC"]/fieldset[7]/table/tbody/tr/td[3]/input',
        '//*[@id="formTC"]/fieldset[7]/table/tbody/tr/td[4]/input',
        '//*[@id="formTC"]/fieldset[7]/table/tbody/tr/td[5]/input',
        '',
        '//*[@id="equipment_list1"]',
        '',
        '//*[@id="formTC"]/fieldset[9]/table/tbody/tr[1]/td[1]/input[1]',
        '//*[@id="formTC"]/fieldset[9]/table/tbody/tr[2]/td[1]/input[1]',
        '//*[@id="formTC"]/fieldset[9]/table/tbody/tr[2]/td[2]/div/input',
        '//*[@id="formTC"]/fieldset[9]/table/tbody/tr[2]/td[1]/input[3]',
        '//*[@id="car_night_keep_cond"]/option[1]',
        '//*[@id="formTC"]/fieldset[9]/table/tbody/tr[1]/td[4]/div/input',
        '//*[@id="formTC"]/fieldset[9]/table/tbody/tr[2]/td[4]/div/input',
        '//*[@id="formTC"]/fieldset[9]/table/tbody/tr[4]/td[1]/div/input',
        '//*[@id="dateInspectionTC1"]'
    ]

    good = []

    try:
        browser.find_element_by_xpath('//*[@id="tabs-contract"]/ul/li[3]/a').click()
    except Exception as ex:
        browser.save_screenshot(f'{cf_screen}Ошибка - Парсинг - Вкладка #2 - Транспортное средство - Не кликается.png')
        print(f'Ошибка при парсинге: Вкладка [#2] "Транспортное средство" - Не кликается!\n  Сообщение: {ex}')
    else:
        browser.implicitly_wait(0.1)

        for i in range(0, len(xpath_text_free)):
            # (2) Если ТЕКСТ - xpath, а ИНПУТ - не xpath
            if xpath_text_free[i][:2] == '//' and xpath_input_free[i][:2] != '//':
                try:
                    # Значение + уникальность для панды
                    text = browser.find_element_by_xpath(xpath_text_free[i]).get_attribute('innerText') + f'_2/2/{i}'

                    text_input = xpath_input_free[i]
                except:
                    pass
                else:
                    good = good + [[text, text_input]]

            # (3) Если ТЕКСТ - не xpath, а ИНПУТ - xpath
            elif xpath_text_free[i][:2] != '//' and xpath_input_free[i][:2] == '//':
                try:
                    # Значение + уникальность для панды
                    text = xpath_text_free[i] + f'_2/3/{i}'

                    # Берем данные по признаку - CHECKED - Да/Нет
                    if i == 3 or i == 36:
                        text_input = 'Да' if browser.find_element_by_xpath(xpath_input_free[i]).get_attribute(
                            'checked') == 'true' else 'Нет'

                    # Берем данные по признаку - CHECKED - Физическое лицо/Юридическое лицо
                    elif i == 15 or i == 19:
                        text_input = 'Физическое лицо' if browser.find_element_by_xpath(
                            xpath_input_free[i]).get_attribute('checked') == 'true' else 'Юридическое лицо'

                    # Берем данные по признаку - VALUE
                    elif i == 31:
                        text_input = browser.find_element_by_xpath(xpath_input_free[i]).get_attribute('value')
                    else:
                        text_input = browser.find_element_by_xpath(xpath_input_free[i]).get_attribute('innerText')

                except:
                    pass
                else:
                    good = good + [[text, text_input]]

            # (4) Иначе всё остальное
            else:
                try:
                    text = browser.find_element_by_xpath(xpath_text_free[i]).get_attribute('innerText')
                    text = re.sub(r'[\t|\n|\r]', '', text)  # Убрать переносы строки
                    text = text.strip(' ') + f'_2/4/{i}'  # Убрать пробелы До/После + Добавление уникальности для панды

                    # Берем данные по признаку - CHECKED
                    if i == 18 or i == 22 or i == 23 or i == 33 or i == 34:
                        text_input = 'Да' if browser.find_element_by_xpath(xpath_input_free[i]).get_attribute(
                            'checked') == 'true' else 'Нет'
                    # Берем данные по признаку - VALUE
                    elif i == 1 or i == 2 or i == 4 or i == 5 or i == 6 or i == 7 or i == 8 or i == 9 or i == 10 or \
                            i == 11 or i == 13 or i == 25 or i == 26 or i == 27 or i == 28 or i == 29 or i == 35 or \
                            i == 38 or i == 39 or i == 40 or i == 41:
                        text_input = browser.find_element_by_xpath(xpath_input_free[i]).get_attribute('value')

                    else:
                        text_input = browser.find_element_by_xpath(xpath_input_free[i]).get_attribute('innerText')

                except:
                    pass
                else:
                    good = good + [[text, text_input]]

        browser.implicitly_wait(15)

    finally:
        return good


# Сканирование параметров страница 3 [Параметры договора] - через selenium
def scanner_param_3(browser):
    xpath_text_free = [
        '//*[@id="tabs-contract"]/ul/li[4]/a',
        '//*[@id="dog_params_form"]/fieldset[1]//div[1]/div[1]//label',
        '//*[@id="dog_params_form"]/fieldset[1]//div[1]/div[2]//label',
        '//*[@id="dog_params_form"]/fieldset[1]//div[1]/div[3]//label',
        '//*[@id="dog_params_form"]/fieldset[1]//div[1]/div[4]//label',
        '//*[@id="dog_params_form"]/fieldset[1]//div[1]/div[5]//label',
        '//*[@id="dog_params_form"]/fieldset[1]//div[1]/div[6]//label',
        '//*[@id="dog_params_form"]/fieldset[1]//div[1]/div[7]//label',
        '//*[@id="dog_params_form"]/fieldset[1]//div[1]/div[8]//label',
        '//*[@id="dog_params_form"]/fieldset[1]//div[1]/div[9]//label',
        '//*[@id="dog_params_form"]/fieldset[1]//div[1]/div[10]//label',
        '//*[@id="dog_params_form"]/fieldset[1]//div[1]/div[11]//label',
        '//*[@id="dog_params_form"]/fieldset[2]/div[1]/div/div/div[1]',
        '//*[@id="dog_params_form"]/fieldset[2]/div[1]/div/div/div[2]/span',
        '//*[@id="dog_params_form"]/fieldset[2]/div[2]/div[1]/div/div[1]',
        '//*[@id="dog_params_form"]/fieldset[2]/div[2]/div[1]/div/div[2]/span',
        '//*[@id="dog_params_form"]/fieldset[2]/div[2]/div[2]/div/div[1]/span',
        '//*[@id="dog_params_form"]/fieldset[3]/legend',
        '//*[@id="dog_params_form"]/fieldset[3]/div',
        '//*[@id="dog_params_form"]/fieldset[3]/span/div',
        '//*[@id="dog_params_form"]/fieldset[3]/select/option[1]',
        '//*[@id="dog_params_form"]/fieldset[4]/legend',
        '//*[@id="dog_params_form"]/fieldset[4]/div[1]/fieldset/legend'
    ]

    xpath_input_free = [
        '',
        '//*[@id="dog_params_form"]/fieldset[1]//div[1]/div[1]//input',
        '//*[@id="dog_params_form"]/fieldset[1]//div[1]/div[2]//input',
        '//*[@id="dog_params_form"]/fieldset[1]//div[1]/div[3]//input',
        '//*[@id="KASKO_dog_type"]/option[2]',
        '//*[@id="dog_params_form"]/fieldset[1]//div[1]/div[5]//input',
        '//*[@id="dog_params_form"]/fieldset[1]//div[1]/div[6]//input',
        '//*[@id="dog_params_form"]/fieldset[1]//div[1]/div[7]//input',
        '//*[@id="KASKO_Currency"]/option[2]',
        '//*[@id="KASKO_Pay_Period"]/option[1]',
        '//*[@id="KASKO_Pay_Type"]/option[1]',
        '//*[@id="dog_params_form"]/fieldset[1]//div[1]/div[111]//input',
        '//*[@id="dog_params_form"]/fieldset[2]/div[1]/div/div/div[1]/input',
        '//*[@id="signatoryListCompl"]',
        '//*[@id="dog_params_form"]/fieldset[2]/div[2]/div[1]/div/div[1]/input',
        '//*[@id="tenderListComp"]',
        '//*[@id="Lot"]',
        '',
        '//*[@id="dog_params_form"]/fieldset[3]/div/input',
        '//*[@id="EA7number"]',
        '',
        '',
        '//*[@id="dog_params_form"]/fieldset[4]/div[1]/fieldset/div/div'
    ]

    good = []

    try:
        browser.find_element_by_xpath('//*[@id="tabs-contract"]/ul/li[4]/a').click()
    except Exception as ex:
        browser.save_screenshot(f'{cf_screen}Ошибка - Парсинг - Вкладка #3 - Параметры договора - Не кликается.png')
        print(f'Ошибка при парсинге: Вкладка [#3] "Параметры договора" - Не кликается!\n  Сообщение: {ex}')
    else:
        browser.implicitly_wait(0.1)

        for i in range(0, len(xpath_text_free)):
            # (2) Если ТЕКСТ - xpath, а ИНПУТ - не xpath
            if xpath_text_free[i][:2] == '//' and xpath_input_free[i][:2] != '//':
                try:
                    # Значение + уникальность для панды
                    text = browser.find_element_by_xpath(xpath_text_free[i]).get_attribute('innerText') + f'_3/2/{i}'

                    text_input = xpath_input_free[i]
                except:
                    pass
                else:
                    good = good + [[text, text_input]]

            # (4) Иначе всё остальное
            else:
                try:
                    text = browser.find_element_by_xpath(xpath_text_free[i]).get_attribute('innerText')
                    text = re.sub(r"[\t|\n|\r]", '', text)  # Убрать переносы строки
                    text = text.strip(' ') + f'_3/4/{i}'  # Убрать пробелы До/После + Уникальность для Панды

                    # Берем данные по признаку - CHECKED
                    if i == 12 or i == 14 or i == 18:
                        text_input = 'Да' if browser.find_element_by_xpath(xpath_input_free[i]).get_attribute(
                            'checked') == 'true' else 'Нет'

                    # Берем данные по признаку - VALUE
                    elif i == 1 or i == 2 or i == 3 or i == 5 or i == 6 or i == 7 or i == 11 or i == 13 or i == 15 \
                            or i == 16 or i == 19:
                        text_input = browser.find_element_by_xpath(xpath_input_free[i]).get_attribute('value')

                    else:
                        text_input = browser.find_element_by_xpath(xpath_input_free[i]).get_attribute('innerText')
                except:
                    pass
                else:
                    good = good + [[text, text_input]]

        browser.implicitly_wait(15)

    finally:
        return good


# Сканер премий и программ по которым нету расчета
def scanner_prize(browser, number):
    """
    Парсинг идёт поблочно, следующим образом:
    1) Главный расчет - базовая программа
    2) Программы по которым имеется расчет
    3) Тарифы и коэффициенты по договору [ОСАГО]
    4) Программы по которым НЕТ расчета
    """

    good = [[]]

    try:
        browser.find_element_by_xpath('//*[@id="tabs-contract"]/ul/li[2]/a').click()
    except Exception as ex:
        browser.save_screenshot(f'{cf_screen}Ошибка - Вкладка #1 - Быстрый расчет КАСКО - Не кликается.png')
        print(f'Ошибка: Вкладка [#1] "Быстрый расчет КАСКО" - Не кликается!\n  Сообщение: {ex}')
    else:

        # Клик по "Программы не соответствующие параметрам расчёта" - HTML страница загоняется в переменную для парсинга
        try:
            browser.find_element_by_xpath('//*[@data-bind="click: showNotCalcProg"]').click()
        except Exception as ex:
            browser.save_screenshot(f'{cf_screen}Ошибка - Программы не соответствующие параметрам - Не кликается.png')
            print(f'Ошибка: "Программы не соответствующие параметрам расчёта" - Не кликается!\nСообщение: {ex}\n')
            website = browser.page_source.encode('utf-8')
        else:
            time.sleep(0.5)
            loading_spinner(browser)
            website = browser.page_source.encode('utf-8')
            browser.find_element_by_id('notCalcProg_close').click()

        tree = html.fromstring(website)

        # Объявление переменных
        txt_key, txt_value, txt_good_key, txt_good_value, txt_bad_key, txt_bad_value, \
        txt_contract_key, txt_contract_value = ('' for i in range(8))

        id_good_programs = ['calcProgramList']
        id_bad_programs = ['dialog_not_calc_prog']

        good = [['Опция', cf_list_twig[number - 1]]]

        # 1) Основная программа и премия
        print(f'\nОсновная программа:')
        try:
            main_key = tree.xpath('//*[@data-bind="foreach: baseProgramListKb"]//a[2]')
            main_value = tree.xpath('//*[@data-bind="foreach: baseProgramListKb"]//td[2]')

            for key in main_key:
                txt_key = key.text
            for value in main_value:
                txt_value = value.text

            if txt_key != '':
                print(f'{txt_key}: {txt_value}')
                good = good + [[txt_key, txt_value]]
        except:
            pass

        # 2) Вывод Программ по которым есть премия!
        print(f'\nПрограммы:')
        good = good + [['Программы', '']]
        for n in range(1, 25):
            try:
                good_key = tree.xpath(f'//*[@id="{id_good_programs[0]}"]//tr[{n}]/td[1]/a[2]')
                good_value = tree.xpath(f'//*[@id="{id_good_programs[0]}"]//tr[{n}]/td[2]')

                for key in good_key:
                    txt_good_key = key.text
                for value in good_value:
                    txt_good_value = value.text

                if txt_good_key != '':
                    print(f'{txt_good_key}: {txt_good_value}')
                    good = good + [[txt_good_key, txt_good_value]]

                txt_good_key, txt_good_value = ('' for i in range(2))
            except:
                pass

        # 3) Тарифы и коэффициенты по договору [ОСАГО]
        print(f'\nТариф и коэффициенты по договору:')
        good = good + [['Тариф и коэффициенты по договору', '']]
        for n in range(1, 10):
            try:
                contract_key = tree.xpath(f'//*[@id="calcPremOSAGO"]/tbody/tr[{n}]/td[1]')
                contract_value = tree.xpath(f'//*[@id="calcPremOSAGO"]/tbody/tr[{n}]/td[2]')

                for key in contract_key:
                    txt_contract_key = key.text
                    txt_contract_key = re.sub(r'[\t|\n|\r]', '', txt_contract_key)  # Убрать переносы строки
                    txt_contract_key = txt_contract_key.strip(' ')  # Убрать пробелы До/После
                for value in contract_value:
                    txt_contract_value = value.text

                if txt_contract_key != '':
                    txt_contract_key = '' if txt_contract_key is None else txt_contract_key
                    print(f'{txt_contract_key}: {txt_contract_value}')
                    good += [[txt_contract_key, txt_contract_value]]

                txt_contract_key, txt_contract_value = ('' for i in range(2))
            except:
                pass

        # 4) Вывод Программ по которым НЕ возможно получить премии!
        print(f'\nПрограммы не соответствующие параметрам расчета:')
        good = good + [['Программы не соответствующие параметрам расчета', '']]
        for n in range(1, 25):
            try:
                bad_key = tree.xpath(f'//*[@id="{id_bad_programs[0]}"]//tr[{n}]/td[1]/span')
                bad_value = tree.xpath(f'//*[@id="{id_bad_programs[0]}"]//tr[{n}]/td[2]/span')

                for key in bad_key:
                    txt_bad_key = key.text
                for value in bad_value:
                    txt_bad_value = value.text

                if txt_bad_key != '':
                    print(f'{txt_bad_key}: {txt_bad_value}')
                    good = good + [[txt_bad_key, txt_bad_value]]

                txt_bad_key, txt_bad_value = ('' for i in range(2))
            except:
                pass

    finally:
        return good


# Авторизация на сайте
def authorization(user, password):
    url_with_auth = f'https://{user}:{password}@cis.makc.ru/pls/fobos/unicus_maks.pkg_contract_ko.maks_main'

    options = webdriver.ChromeOptions()
    # options.add_argument('--start-maximized')
    options.add_argument("--headless")
    options.add_argument("--window-size=1700x1200")

    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) \
    Chrome/89.0.4389.90 Safari/537.36 OPR/74.0.3911.218")
    options.add_argument('--disable-blink-features=AutomationControlled')

    options.add_experimental_option(
        'prefs',
        {
            # 'profile.default_content_setting_values.notifications': 2,
            # 'profile.managed_default_content_settings.cookies': 2,
            # 'profile.managed_default_content_settings.popups': 2,
            'profile.managed_default_content_settings.geolocation': 2,
            # 'profile.managed_default_content_settings.plugins': 2,
            # 'profile.managed_default_content_settings.fullscreen': 2,
            # 'profile.managed_default_content_settings.javascript': 2,
            # 'profile.managed_default_content_settings.images': 2,
            # 'profile.managed_default_content_settings.mixed_script': 2,
            # 'profile.managed_default_content_settings.media_stream': 2,
            # 'profile.managed_default_content_settings.stylesheets': 2
        }
    )

    browser = webdriver.Chrome(options=options)

    # Запуск
    browser.get(url_with_auth)

    # Неявные ожидания - до 15 секунд
    browser.implicitly_wait(15)

    return browser


# Сохранение страницы - Бэкап
def save_page(browser, i):
    print('\n----- Сохранение страницы в HTML -----')
    website = browser.page_source.encode('utf-8')

    page = f'{cf_trunk}/{cf_list_branches[0]}/'

    # Если файл Страница_1_Дата.html уже есть, перезаписать
    try:
        f = open(f'{page}Страница_{i}_{date.today().strftime("%d.%m.%Y")}.html')
        f.close()
    except FileNotFoundError:
        print(f'Файл {page}' + str(f'Страница_{i}_{date.today().strftime("%d.%m.%Y")}.html не существует! - Создаем'))
    else:
        print(f'Файл {page}' + str(f'Страница_{i}_{date.today().strftime("%d.%m.%Y")}.html существует! - Перезапись'))
    finally:
        with open(f'{page}Страница_{i}_{date.today().strftime("%d.%m.%Y")}.html', 'wb') as output_file:
            output_file.write(website)


# Закрытие браузера
def exit_code(browser):
    browser.close()
    browser.quit()


# ------ Ввод ВСЕХ данных из Excel ------
def excel_import():
    print(f'----- Ввод данных из Экселя -----')

    # Количество строк данных в Excel - Всегда +1
    strok = 67 + 1

    xl = pd.read_excel('_Данные_СК_МАКС.xlsx', sheet_name='Лист1', usecols='B:D', header=0, nrows=strok)
    list_import_excel = []
    for i in range(0, strok):
        try:
            if xl.iat[i, 0] == '' or pd.isnull(xl.iat[i, 0]):
                xl.iat[i, 0] = ''
        except IndexError as ix:
            print(f'IndexError = {ix}')
        else:
            list_import_excel.append(xl.iat[i, 0])

    return list_import_excel


# Вывод всех данных в Эксель
def excel_export():
    data_today = date.today().strftime('%d.%m.%Y')
    time_now = time.strftime("%H.%M", time.localtime())

    file_excel = f'{cf_trunk}/Расчет B2B МАКС [{data_today} - {time_now}].xlsx'

    # Вывод в Excel расчета
    try:
        # Параметры - Запускать Заголовочник <*= Стереть!
        good_1 = splitting_by_titles(GL_good_1)
        good_2 = splitting_by_titles(GL_good_2)
        good_3 = splitting_by_titles(GL_good_3)

        # Премии
        prize = splitting_by_titles(GL_prize)

        # Соединение ПАРАМЕТРОВ и ПРЕМИИ
        dict_1 = dict()

        # ПАРАМЕТРЫ
        for it in range(len(good_1.values)):
            dict_1[it] = good_1.values[it]

        # Добавление -> Параметров
        for i in range(2, 3 + 1):
            total = len(dict_1)
            for it in range(len(eval(f'good_{i}.values'))):
                dict_1[it + total + 1] = (eval(f'good_{i}.values[{it}]'))

        # Добавление -> Премии
        total = len(dict_1)
        for it in range(len(prize.values)):
            dict_1[it + total + 1] = (eval(f'prize.values[{it}]'))

        df_2 = pd.DataFrame(data=dict_1)

        df_2 = df_2.T

        df_2[0] = df_2[0].apply(lambda x: panda_trash(x))

    except Exception as ex:
        print(f'Ошибка: Вывод данных в Excel!\n  Сообщение: {ex}')

    else:
        with pd.ExcelWriter(file_excel) as writer:
            df_2.to_excel(writer, sheet_name='Пред|Итоговый', index=False, header=False)

            sheet_2 = writer.sheets['Пред|Итоговый']
            workbook = writer.book

            first_fmt = workbook.add_format({'font_name': 'Times New Roman',
                                             'font_size': 9,  # Размер шрифта
                                             'text_wrap': True,  # Перенос текста
                                             'valign': 'left',  # Выравнивание текста по горизонтали
                                             'align': 'vcenter',  # Выравнивание текста по вертикали
                                             })

            second_fmt = workbook.add_format({'font_name': 'Times New Roman',
                                              'font_size': 9,  # Размер шрифта
                                              'text_wrap': True,  # Перенос текста
                                              'valign': 'center',  # Выравнивание текста по горизонтали
                                              'align': 'vcenter',  # Выравнивание текста по вертикали
                                              })

            format1 = workbook.add_format({'bold': 1,
                                           'bg_color': '#1DACD6'})

            # //////////////////// СТИЛЬ ЗАГОЛОВКОВ -> ПАРАМЕТРОВ ////////////////////
            fmt_head_1 = workbook.add_format({'bold': 1,
                                              'bg_color': '#C86B85',
                                              'font_name': 'Times New Roman',
                                              'font_size': 9,  # Размер шрифта
                                              'valign': 'left',  # Выравнивание текста по горизонтали
                                              'align': 'vcenter',  # Выравнивание текста по вертикали
                                              })

            fmt_head_2 = workbook.add_format({'bold': 1,
                                              'bg_color': '#FFD3B6',
                                              'font_name': 'Times New Roman',
                                              'font_size': 9,  # Размер шрифта
                                              'valign': 'center',  # Выравнивание текста по горизонтали
                                              'align': 'vcenter',  # Выравнивание текста по вертикали
                                              })

            # //////////////////// СТИЛЬ ЗАГОЛОВКОВ -> ПРЕМИЙ ////////////////////
            fmt_prime_1 = workbook.add_format({'bold': 1,
                                               'bg_color': '#1FAB89',
                                               'font_name': 'Times New Roman',
                                               'font_size': 9,  # Размер шрифта
                                               'valign': 'left',  # Выравнивание текста по горизонтали
                                               'align': 'vcenter',  # Выравнивание текста по вертикали
                                               })

            fmt_prime_2 = workbook.add_format({'bold': 1,
                                               'bg_color': '#62D2A2',
                                               'font_name': 'Times New Roman',
                                               'font_size': 9,  # Размер шрифта
                                               'valign': 'center',  # Выравнивание текста по горизонтали
                                               'align': 'vcenter',  # Выравнивание текста по вертикали
                                               })

            # Заголовки - которые нужно закрасить
            # Параметры
            heading_1_param = [
                'Быстрый расчет КАСКО',
                'Транспортное средство',
                'Параметры договора'
            ]
            heading_2_param = [
                'Транспортное средство №1',
                'Условия договора',
                'Способы уменьшения стоимости',
                'Дополнительные параметры',
                'Контактная информация',
                'Документ транспортного средства',
                'Страхователь',
                'Выгодоприобретатель',
                'Собственник',
                'Лица допущенные к управлению ТС',
                'Дополнительное оборудование',
                'Результат осмотра ТС',
                'Бланки',
                'Андеррайтинг, СЭиИЗ'
            ]

            # Премии
            heading_1_prime = ['Опция']
            heading_2_prime = [
                'Программы',
                'Тариф и коэффициенты по договору',
                'Программы не соответствующие параметрам расчета'
            ]

            # Для ПРЕДВАРИТ/ИТОГОВОГО РАСЧЕТА
            """
            df_2            - DataFrame
            sheet_1         - Название листа
            heading_1_param - Название параметров
            fmt_head_1      - Стиль
            """

            style_of_design(df_2, sheet_2, heading_1_param, fmt_head_1)
            style_of_design(df_2, sheet_2, heading_2_param, fmt_head_2)
            # Премии
            style_of_design(df_2, sheet_2, heading_1_prime, fmt_prime_1)
            style_of_design(df_2, sheet_2, heading_2_prime, fmt_prime_2)

            # Для ПРЕДВАРИТ/ИТОГОВОГО РАСЧЕТА
            # Закраска всех уникальных
            n = 1
            prob = []

            for i in df_2.itertuples():
                for j in range(2, len(i)):
                    if i[j] == i[j]:
                        prob.append(str(i[j]).lower())
                cnt = Counter(prob).most_common(1)
                prob.clear()
                # Если cnt[..][1] == 1 это значит все уникальные переменные, то их и красить
                for j in range(2, len(i)):
                    if str(i[j]).lower() != cnt[0][0] and cnt[0][1] != 1 and i[j] == i[j] and i[j] != '':
                        # Эту ячейку нужно закрасить
                        cell = xl_rowcol_to_cell(n - 1, j - 1)
                        sheet_2.conditional_format(cell, {'type': 'no_errors',
                                                          'format': format1})
                    # Все уникальные - закрасить
                    elif cnt[0][1] == 1 and i[j] == i[j] and i[j] != '':
                        cell = xl_rowcol_to_cell(n - 1, j - 1)
                        sheet_2.conditional_format(cell, {'type': 'no_errors',
                                                          'format': format1})
                # n - это порядковый номер строки
                n += 1

            # Первый столбец применяет формат
            sheet_2.set_column(0, 0, 27.86, first_fmt)

            # Весь лист применяет формат
            sheet_2.set_column(1, len(df_2.columns) - 1, 27.86, second_fmt)

            writer.save()


# Разбитие общего СПИСКА по ЗАГОЛОВКам
def splitting_by_titles(list_import):
    """
    list_import     - Список - Входные значения. Структура: [[['',''],['','']]]
    index           - Список - Содержит координаты заголовков
    list_processing - Список - Временный список, участвующий между Импортом и Экспортом
    list_export     - Список - Разбитый список по заголовкам (delimiter)
    """

    # Название ЗАГОЛОВКОВ
    delimiter = [
        'Быстрый расчет КАСКО',
        'Транспортное средство',
        'Параметры договора',
        'Транспортное средство №1',
        'Условия договора',
        'Способы уменьшения стоимости',
        'Дополнительные параметры',
        'Контактная информация',
        'Документ транспортного средства',
        'Страхователь',
        'Выгодоприобретатель',
        'Собственник',
        'Лица допущенные к управлению ТС',
        'Дополнительное оборудование',
        'Результат осмотра ТС',
        'Бланки',
        'Андеррайтинг, СЭиИЗ',
        'Опция',
        'Программы',
        'Тариф и коэффициенты по договору',
        'Программы не соответствующие параметрам расчета'
    ]

    index, list_processing, list_export = ([] for i in range(3))

    # Поиск Заголовков и определение их координат
    for x in range(0, len(list_import)):
        # Выводится - по потокам
        for y in range(0, len(list_import[x])):
            for one_delimiter in delimiter:
                if list_import[x][y][0] == one_delimiter:
                    # Заполняем список - координатами заголовка
                    index.append([x, y, 0])

    # Соотношение индексов заголовков с основным списком и его разбив (срез)
    for x in range(0, len(list_import)):
        for y in range(0, len(list_import[x])):
            # При переборе основного списка -> Смотрим все координаты и соотносим
            for number in range(0, len(index)):
                # Находим заголовки по полученным ранее координатам
                if x == int(index[number][0]) and y == int(index[number][1]):
                    # Срез с ЗАГОЛОВКА (включительно) и до следующего заголовка (НЕ включительно)
                    section = list_import[x][y:]
                    try:
                        next_y = int(index[number + 1][1])
                    except IndexError:
                        list_processing += [section]
                    else:
                        # Если следующий заголовок есть, то резать до него, иначе всё порезано
                        if next_y != 0:
                            step = next_y - int(index[number][1])
                            section = section[:step]

                        list_processing += [section]

    # Рассортировка по блокам - разбитого списка
    for number in range(0, len(delimiter)):
        for x in range(0, len(list_processing)):

            if list_processing[x][0][0] == delimiter[number]:
                # Создаем список БЛОКА -> отправляем на merge (pandas)
                list_export += [list_processing[x]]

        # Если Блок с названием Заголовка существует
        if len(list_export) != 0:

            # -> отправляем на merge (pandas) - после зачищаем
            panda_export = panda_param(list_export)

            # Словарь - Соединение/наращивание объединенных блоков
            good = panda_export if number == 0 else pd.concat([good, panda_export])

            list_export.clear()

    return good


# Функция, чтобы убрать мусор
def panda_trash(value):
    value = value.split('_')[0]
    return value


# Создание DataFrame из Списков
def panda_param(doc):
    good = pd.DataFrame(doc[0], columns=['name', '1'])

    try:
        for x in range(1, len(doc)):
            da = pd.DataFrame(doc[x], columns=['name', '{}'.format(x + 1)])
            good = pd.merge(good, da, on='name', how='outer')
    except ValueError:
        # Если какой-то из потоков не завершился, то выскакивает ошибка, что столбцов 0, а нужно 2
        pass

    return good


def style_of_design(df, sheet, heading, fmt_head):
    # df - DataFrame
    # sheet - Название листа
    # heading - Название параметров
    # fmt_head - Стиль

    check = False
    for i in range(0, df.shape[0]):
        if not check:
            for j in range(0, len(heading)):
                if df.iat[i, 0] == heading[j]:
                    sheet.conditional_format(i, 0, i, df.shape[1] - 1, {'type': 'no_errors',
                                                                        'format': fmt_head})
                    # Затереть старое значение (исключает закраску повторяющихся заголовков)
                    heading[j] = 'xxx'
        else:
            break


# Запуск ПОТОКОВ
def threads():
    # Выгрузка данных из Экселя
    import_data_list = excel_import()

    # Создание папок
    creating_folders()

    t1 = Thread(target=work, args=(import_data_list, 1))
    t1.start()

    t2 = Thread(target=work, args=(import_data_list, 2))
    t2.start()

    t3 = Thread(target=work, args=(import_data_list, 3))
    t3.start()

    t4 = Thread(target=work, args=(import_data_list, 4))
    t4.start()

    t5 = Thread(target=work, args=(import_data_list, 5))
    t5.start()

    # Когда все потоки завершились, то вывести все значения - загнать в EXCEL!
    t1.join()
    t2.join()
    t3.join()
    t4.join()
    t5.join()

    # Переименовать папки со скринами
    rename_folders()

    # Вывод в Excel файл
    excel_export()


# Создание папок
def creating_folders():
    try:
        os.mkdir(cf_trunk)
    except OSError:
        pass

    for branch in cf_list_branches:
        try:
            os.makedirs(f'{cf_trunk}/{branch}')
        except OSError:
            pass


# Создание папок для скринов, согласно потоку => _Результаты/Скрины B2B/number
def creating_folders_screen(num_thread):
    try:
        os.makedirs(f'{cf_trunk}/{cf_list_branches[1]}/{num_thread}')
    except OSError:
        pass


# Переименовываем папки с выводами
def rename_folders():
    # Переименовывание папок со скринами
    for number in range(0, len(cf_list_twig)):
        try:
            os.rename(f'{cf_trunk}/{cf_list_branches[1]}/{number + 1}',
                      f'{cf_trunk}/{cf_list_branches[1]}/{cf_list_twig[number]}')
        except OSError:
            pass  # Файл не найден
        except IndexError as ix:
            print(f'Список не полный -> вышел за рамки! Ошибка: {ix}')

    data_today = date.today().strftime('%d.%m.%Y')
    time_now = time.strftime("%H.%M", time.localtime())

    # Добавление уникальности (Даты и времени расчета) для папок со скринами
    try:
        os.rename(f'{cf_trunk}/{cf_list_branches[1]}', f'{cf_trunk}/{cf_list_branches[1]} [{data_today} - {time_now}]')
    except OSError:
        # Файл не найден
        pass
    except IndexError as ix:
        print(f'Список не полный -> вышел за рамки! Ошибка: {ix}')


def main():
    start_time = time.time()

    threads()

    time_work = time.time() - start_time
    minutes = time_work // 60
    seconds = time_work - minutes * 60
    print(f'\n@-@-@ Время работы программы: {int(minutes)}m {int(seconds)}s @-@-@')


if __name__ == "__main__":
    main()
