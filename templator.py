import datetime
import calendar
from pathlib import Path
from docxtpl import DocxTemplate

import PySimpleGUI as sg
sg.theme('LightGrey1')


# path to the main template
document_path = Path(__file__).parent / "main_template.docx"

doc = DocxTemplate(document_path)

delivery = {
    1: "одного",
    2: "двох",
    3: "трьох",
    4: "чотирьох",
    5: "п'яти",
    6: "шести",
    7: "семи",
    8: "восьми",
    9: "дев'яти",
    10: "десяти",
    11: "одинадцяти",
    12: "дванадцяти",
    13: "тринадцяти",
    14: "чотирнадцяти",
    15: "п'ятнадцяти",
    16: "шістнадцяти",
    17: "сімнадцяти",
    18: "вісімнадцяти",
    19: "дев'ятнадцяти",
    20: "двадцяти",
    21: "двадцяти одного",
    22: "двадцяти двох",
    23: "двадцяти трьох",
    24: "двадцяти чотирьох",
    25: "двадцяти п'яти",
    26: "двадцяти шести",
    27: "двадцяти семи",
    28: "двадцяти восьми",
    29: "двадцяти дев'яти",
    30: "тридцяти",
    31: "тридцяти одного",
    32: "тридцяти двох",
    33: "тридцяти трьох",
    34: "тридцяти чотирьох",
    35: "тридцяти п'яти",
    36: "тридцяти шести",
    37: "тридцяти семи",
    38: "тридцяти восьми",
    39: "тридцяти дев'яти",
    40: "сорока",
    41: "сорока одного",
    42: "сорока двох",
    43: "сорока трьох",
    44: "сорока чотирьох",
    45: "сорока п'яти",
    46: "сорока шести",
    47: "сорока семи",
    48: "сорока восьми",
    49: "сорока дев'яти",
    50: "п'ятдесяти",
    51: "п'ятдесяти одного",
    52: "п'ятдесяти двох",
    53: "п'ятдесяти трьох",
    54: "п'ятдесяти чотирьох",
    55: "п'ятдесяти п'яти",
    56: "п'ятдесяти шести",
    57: "п'ятдесяти семи",
    58: "п'ятдесяти восьми",
    59: "п'ятдесяти дев'яти",
    60: "шістдесяти",
    61: "шістдесяти одного",
    62: "шістдесяти двох",
    63: "шістдесяти трьох",
    64: "шістдесяти чотирьох",
    65: "шістдесяти п'яти",
    66: "шістдесяти шести",
    67: "шістдесяти семи",
    68: "шістдесяти восьми",
    69: "шістдесяти дев'яти",
    70: "сімдесяти",
    71: "сімдесяти одного",
    72: "сімдесяти двох",
    73: "сімдесяти трьох",
    74: "сімдесяти чотирьох",
    75: "сімдесяти п'яти",
    76: "сімдесяти шести",
    77: "сімдесяти семи",
    78: "сімдесяти восьми",
    79: "сімдесяти дев'яти",
    80: "восьмидесяти",
    81: "восьмидесяти одного",
    82: "восьмидесяти двох",
    83: "восьмидесяти трьох",
    84: "восьмидесяти чотирьох",
    85: "восьмидесяти п'яти",
    86: "восьмидесяти шести",
    87: "восьмидесяти семи",
    88: "восьмидесяти восьми",
    89: "восьмидесяти дев'яти",
    90: "дев'яноста",
    91: "дев'яноста одного",
    92: "дев'яноста двох",
}


date = datetime.datetime.today()
months = {
    'January': 'січня',
    'February': 'лютого',
    'March': 'березня',
    'April': 'квітня',
    'May': 'травня',
    'June': 'червня',
    'July': 'липня',
    'August': 'серпня',
    'September': 'вересня',
    'October': 'жовтня',
    'November': 'листопада',
    'December': 'грудня'
}
month = months[calendar.month_name[date.month]]


# all fields with information in a pop-up window
layout = [[sg.Text('Номер договору:'), sg.Input(
    key='CONTRACT_NUMBER', do_not_clear=True, enable_events=True, size=(6, 1))],
    [sg.Text('Клієнт платник ПДВ?:'),
     sg.Combo(['так', 'ні', '2 група'], key='TAXPAYER', default_value='так')],
    [sg.Text('Який буде тип оплати?:'), sg.Combo(
        ['100% передоплата', '50% передоплата', 'післяплата'], key='PAID', default_value='100% передоплата')],
    [sg.Text('Назва контрагента:'), sg.Input(
        key='CLIENT', do_not_clear=True, enable_events=True)],
    [sg.Text('ПІБ директора контрагента:'), sg.Input(
        key='DIRECTOR', do_not_clear=True, enable_events=True)],
    [sg.Text('Поставка товару буде протягом:'),
     sg.Combo([i for i in range(1, 101)], key='DELIEVERY_DAYS', default_value=7), sg.Text('днів')],
    [sg.Text('Замовлення забирають з:'), sg.Combo(
        ['склад СКІФ-ІНВЕСТ', 'склад транспортної компанії в КР', 'склад покупця за адресою ...', 'склад транспортної компанії в місті покупця'], key='WAREHOUSE', default_value='склад транспортної компанії в місті покупця')],
    [sg.Text('EMAIL:'), sg.Input(
        key='EMAIL', do_not_clear=True, enable_events=True,)],
    [sg.Text('Юридична адреса контрагента:'), sg.Input(
        key='LEGAL_ADDRESS', do_not_clear=True, enable_events=True)],
    [sg.Text('Фактична адреса контрагента:'), sg.Input(
        key='ACTUAL_ADDRESS', do_not_clear=True, enable_events=True)],
    [sg.Text('Банк контрагента№1:'), sg.Input(
        key='BANK_NAME_1', do_not_clear=True, enable_events=True)],
    [sg.Text('Розрахунковий рахунок (IBAN):'),
     sg.Input(key='IBAN', do_not_clear=True, enable_events=True)],
    [sg.Text('Банк контрагента №2:'), sg.Input(
        key='BANK_NAME_2', do_not_clear=True, enable_events=True)],
    [sg.Text('Розрахунковий рахунок (IBAN №2):'),
     sg.Input(key='IBAN2', do_not_clear=True, enable_events=True)],
    [sg.Text('Телефон 1'), sg.Input(
        key='PHONE1', do_not_clear=True, enable_events=True, size=(16, 1)), sg.Text('Телефон 2'), sg.Input(
        key='PHONE2', do_not_clear=True, enable_events=True, size=(16, 1))],
    [sg.Text('Код банку (МФО):'), sg.Input(key='BANK_ID', do_not_clear=True, enable_events=True, size=(
        10, 1)), sg.Text('ЄДРПОУ:'), sg.Input(key='EDRPOU', do_not_clear=True, enable_events=True, size=(9, 1))],
    [sg.Text('ІПН:'), sg.Input(
        key='IPN', do_not_clear=True, enable_events=True)],
    [sg.Text('Свідоцтво платника ПДВ:'), sg.Input(
        key='TAXES_NUMBER', do_not_clear=True, enable_events=True)],
    [sg.Button('Згенерувати договір')]
]

window = sg.Window('Генератор договорів 0.9', layout,
                   element_justification='right')


fl_docenko = "Фізична особа – підприємець ДОЦЕНКО ЮРІЙ ЮРІЙОВИЧ  (надалі іменується «Постачальник») номер запису  про державну реєстрацію в Єдиному державному реєстрі юридичних осіб, фізичних осіб-підприємців та громадських формувань 22270000000063066 від 25.07.2016, з однієї сторони,"

fl_skif_rudich = "Товариство з обмеженою відповідальністю «СKIФ ІНВЕСТ» (надалі іменується «Постачальник»), в особі Директора Рудича Володимира Васильовича, що діє на підставі Статуту"

fl_skif_docenko = "Товариство з обмеженою відповідальністю «СKIФ ІНВЕСТ»  (надалі іменується «Постачальник») в особі Комерційного директора Доценко Юрія Юрійовича, що діє на підставі Доручення №6 від 01.02.2021"

rekviziti_docenko = "ФІЗИЧНА ОСОБА - ПІДПРИЄМЕЦЬ\n\
ДОЦЕНКО ЮРІЙ ЮРІЙОВИЧ\n\
Адреса: 50046, м. Кривий Ріг, вул. Всебратське-2, дом № 45, кв.39\n\
П/р UA923057500000026000053532165 у Банк АТ КБ \"ПРИВАТБАНК\"\n\
Код банку 305750\n\
Ідентифікаційний код 3190224255\n\
тел. (050) 109-49-14; (096) 409-49-14\n\
e-mail: dotsenkoyurij@gmail.com"

rekviziti_skif = 'ТОВ «СКІФ ІНВЕСТ»\n\
Адреса: 50005, Дніпропетровська обл.,\n\
м. Кривий Ріг, вул. Домобудівна, 25А\n\
IBAN: UA 16 320984 00000 26003210415714\n\
у АТ «ПРОКРЕДИТ БАНК»\n\
П/р UA413223130000026009000010816,\n\
Банк АТ \"Укрексімбанк\"\n\
Код банку 320984\n\
тел. 096-409-49-14, 050-109-49-14\n\
ЄДРПОУ 35601501\n\
ІПН 356015004822\n\
Свідоцтво платника ПДВ 200026344'

skif = 'ТОВ «СКІФ ІНВЕСТ»'
fop_docenko = 'ФОП ДОЦЕНКО Ю. Ю.'

skif_pdv = "На момент укладення цього Договору Постачальник є:\nПлатником податку на прибуток підприємств на загальних умовах згідно Податкового кодексу України. \nПлатником податку на додану вартість на загальних умовах згідно Податкового кодексу України."
fop_pdv = "На момент укладення цього Договору Постачальник є: платником єдиного податку 2 групи."

taxpayer = 'На момент укладення цього Договору Покупець є Платником податку на прибуток підприємств на загальних умовах згідно Податкового кодексу України.'
no_taxpayer = 'На момент укладення цього Договору Покупець не є Платником податку на прибуток підприємств на загальних умовах згідно Податкового кодексу України.'
second_group = "На момент укладення цього Договору Покупець є: платником єдиного податку 2 групи. "

director_docenko = "Доценко Ю.Ю."
director_rudich = "Рудич В.В."


while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    elif event == 'Згенерувати договір':
        if values['WAREHOUSE'] == 'склад СКІФ-ІНВЕСТ':
            values['WAREHOUSE'] = "склад Постачальника, за адресою:  Дніпропетровська обл., м. Кривий Ріг,  вул. Домобудівна 25а;"
        elif values['WAREHOUSE'] == 'склад транспортної компанії в КР':
            values['WAREHOUSE'] = "склад транспортної компанії (САТ, Нова Пошта) за адресою: Дніпропетровська обл. м. Кривий Ріг, якщо інше не вказано у Специфікаціях або Графіках поставки до даного Договору."
        elif values['WAREHOUSE'] == 'склад покупця за адресою ...':
            values['WAREHOUSE'] = "склад Покупця, за адресою: ___________"
        elif values['WAREHOUSE'] == 'склад транспортної компанії в місті покупця':
            values['WAREHOUSE'] = "склад транспортної компанії в місті Покупця (САТ, Нова Пошта)"

        prepaid = f"Поставка Товару (товарів) здійснюється  Постачальником протягом {values['DELIEVERY_DAYS']} ({delivery[values['DELIEVERY_DAYS']]}) робочих днів від дати попередньої оплати на умовах «Франко-завод» » EXW/«Фрахт/перевезення оплачені до» CPT, згідно правил «ІНКОТЕРМС» (в редакції 2010):"
        half_paid = f"Поставка Товару (товарів) здійснюється  Постачальником протягом {values['DELIEVERY_DAYS']} ({delivery[values['DELIEVERY_DAYS']]}) робочих днів від дати 50%  попередньої оплати на умовах «Франко-завод» » EXW/«Фрахт/перевезення оплачені до» CPT, згідно правил «ІНКОТЕРМС» (в редакції 2010):"
        post_paid = f"Поставка Товару (товарів) здійснюється  Постачальником протягом {values['DELIEVERY_DAYS']} ({delivery[values['DELIEVERY_DAYS']]}) робочих днів від дати підтвердження замовником продукції на умовах «Франко-завод» » EXW/«Фрахт/перевезення оплачені до» CPT, згідно правил «ІНКОТЕРМС» (в редакції 2010) з  подальшою  оплатою протягом _____ днів:"

        values['DAYS_TEXT'] = delivery[values['DELIEVERY_DAYS']]
        # по замовчуванню стоїть платник ПДВ і передоплата, постачальник - скіф
        values['DATE'] = date.strftime(f"%d {month} %Y") + ' р.'
        # values['DELIEVERY'] = f"{values['DELIEVERY_DAYS']} {delivery[values['DELIEVERY_DAYS']]}"

        if values['TAXPAYER'] == 'так':
            values['TAXPAYER'] = taxpayer
        elif values['TAXPAYER'] == 'ні':
            values['TAXPAYER'] = no_taxpayer
        elif values['TAXPAYER'] == '2 група':
            values['TAXPAYER'] = second_group

        if values['PAID'] == '100% передоплата':
            values['PAID'] = prepaid
        elif values['PAID'] == '50% передоплата':
            values['PAID'] = half_paid
        elif values['PAID'] == 'післяплата':
            values['PAID'] = post_paid

        if values['TAXPAYER'] == taxpayer and values['PAID'] == prepaid:
            values['WEPDV'] = skif_pdv
            values['FIRST_LINE'] = fl_skif_docenko
            values['REKVIZITI'] = rekviziti_skif
            values['BOSS'] = director_docenko
            values['POSTACHALNIK'] = skif
        elif values['TAXPAYER'] == taxpayer and values['PAID'] == half_paid:
            values['WEPDV'] = skif_pdv
            values['FIRST_LINE'] = fl_skif_rudich
            values['REKVIZITI'] = rekviziti_skif
            values['BOSS'] = director_rudich
            values['POSTACHALNIK'] = skif
        elif values['TAXPAYER'] == taxpayer and values['PAID'] == post_paid:
            values['WEPDV'] = skif_pdv
            values['FIRST_LINE'] = fl_skif_rudich
            values['REKVIZITI'] = rekviziti_skif
            values['BOSS'] = director_rudich
            values['POSTACHALNIK'] = skif
        elif values['TAXPAYER'] != taxpayer and values['PAID'] == prepaid:
            values['WEPDV'] = fop_pdv
            values['POSTACHALNIK'] = fop_docenko
            values['FIRST_LINE'] = fl_docenko
            values['REKVIZITI'] = rekviziti_docenko
            values['BOSS'] = director_docenko
        elif values['TAXPAYER'] != taxpayer and values['PAID'] == half_paid:
            values['WEPDV'] = fop_pdv
            values['POSTACHALNIK'] = fop_docenko
            values['FIRST_LINE'] = fl_docenko
            values['REKVIZITI'] = rekviziti_docenko
            values['BOSS'] = director_docenko
        elif values['TAXPAYER'] != taxpayer and values['PAID'] == post_paid:
            values['WEPDV'] = fop_pdv
            values['POSTACHALNIK'] = fop_docenko
            values['FIRST_LINE'] = fl_docenko
            values['REKVIZITI'] = rekviziti_docenko
            values['BOSS'] = director_docenko

        # render the document
        doc.render(values)
        output_path = Path(__file__).parent / \
            f"{values['CONTRACT_NUMBER']} - {values['CLIENT']}.docx"
        doc.save(output_path)
        sg.popup('Документ успішно збережено\nНехай щастить!')
