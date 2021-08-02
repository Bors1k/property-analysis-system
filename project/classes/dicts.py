# Для того, чтобы excel рассчитывал срок использования
lifetime = {
    'стол': 10,
    'шкаф': 10,
    'кресл': 10,
    'стул': 5,
    'тумб': 10,
    'диван':7,
    'мебел': 10,
    'вешалка': 10,
    'зеркало': 10,
    'кондиционер': 5,
    'холодильник': 10,
    'телевизор': 10,
    'графин': 3,
    'кувшин': 3,
    'стакан': 3,
    'портьеры': 5,
    'тюль': 5,
    'жалюз': 5,
    'ковро': 10,
    'карта': 5,
    'лампа': 5,
    'настольный набор': 5,
    'часы': 10,
    'кронштейн': 7,
    'цифрового тел': 3,
    'флаг': 7,
    'карниз': 5,
    'экран защитный': 5,
    'доска': 5,
    'чайник': 5,
    'печь': 5,
    'стеллаж': 10,
    'стелаж': 10,
    'стенд': 5,
    'комод': 10,
    'кофемашин': 5,
    'кофеварк': 5,
    'микроволн': 5
}

# для выборки
choose_position = {
    'стол': 'стол',
    'шкаф': 'шкаф',
    'кресл': 'кресло',
    'стул': 'стул',
    'тумб': 'тумба',
    'диван': 'диван',
    # 'мебел': 'мебель',
    'вешалка': 'вешалка',
    'зеркало': 'зеркало',
    'кондиционер': 'кондиционер',
    'холодильник': 'холодильник',
    'телевизор': 'телевизор',
    'графин': 'графин',
    'кувшин': 'кувшин',
    'стакан': 'стакан',
    'портьеры': 'портьеры',
    'тюль': 'тюль',
    'жалюз': 'жалюзи',
    'ковро': 'ковролин',
    'карта': 'карта',
    'лампа': 'лампа',
    'настольный набор': 'настольный набор',
    'часы': 'часы',
    'кронштейн': 'кронштейн',
    'цифрового тел': 'цифрового тел',
    'флаг': 'флаг',
    'карниз': 'карниз',
    'экран защитный': 'экран защитный',
    'доска': 'доска',
    'чайник': 'чайник',
    'печь': 'печь',
    'стеллаж': 'стеллаж',
    'стелаж': 'стелаж',
    'стенд': 'стенд',
    'комод': 'комод',
    'кофемашин':  'кофемашина',
    'кофеварк': 'кофеварка',
    'микроволн': 'микроволновка'
}

choose_position_header = {
    'Количество столов': 'стол',
    'Количество шкафов': 'шкаф',
    'Количество кресел': 'кресло',
    'Количество стульев': 'стул',
    'Количество тумб': 'тумба',
    'Количество диванов': 'диван',
    # 'Количество мебели': 'мебель',
    'Количество вешалок': 'вешалка',
    'Количество зеркал': 'зеркало',
    'Количество кондиционеров': 'кондиционер',
    'Количество холодильников': 'холодильник',
    'Количество телевизоров': 'телевизор',
    'Количество графинов': 'графин',
    'Количество кувшинов': 'кувшин',
    'Количество стаканов': 'стакан',
    'Количество портьер': 'портьеры',
    'Количество тюлей': 'тюль',
    'Количество жалюзей': 'жалюзи',
    'Количество ковров': 'ковер',
    'Количество карт': 'карта',
    'Количество ламп': 'лампа',
    'Количество настольных наборов': 'настольный набор',
    'Количество часов': 'часы',
    'Количество кронштейнов': 'кронштейн',
    'Количество цифровых тел.': 'цифрового тел',
    'Количество флагов': 'флаг',
    'Количество карнизов': 'карниз',
    'Количество экранов защитных': 'экран защитный',
    'Количество досок': 'доска',
    'Количество чайников': 'чайник',
    'Количество печей': 'печь',
    'Количество стеллажей': 'стеллаж',
    'Количество стелажей': 'стелаж',
    'Количество стендов': 'стенд',
    'Количество комодов': 'комод',
    'Количество кофемашин':  'кофемашина',
    'Количество кофеварк': 'кофеварка',
    'Количество СВЧ': 'микроволновка'
}

choose_position_header_evry_two = {
    'Количество столов': 'Из них столы с превышенным сроком',
    'Количество шкафов': 'Из них шкафы с превышенным сроком',
    'Количество кресел': 'Из них кресла с превышенным сроком',
    'Количество стульев': 'Из них стулья с превышенным сроком',
    'Количество тумб': 'Из них тумбы с превышенным сроком',
    'Количество диванов': 'Из них диваны с превышенным сроком',
    # 'Количество мебели': 'Из них мебель с превышенным сроком',
    'Количество вешалок': 'Из них вешалки с превышенным сроком',
    'Количество зеркал': 'Из них зеркала с превышенным сроком',
    'Количество кондиционеров': 'Из них кондиционеры с превышенным сроком',
    'Количество холодильников': 'Из них холодильники с превышенным сроком',
    'Количество телевизоров': 'Из них телевизоры с превышенным сроком',
    'Количество графинов': 'Из них графины с превышенным сроком',
    'Количество кувшинов': 'Из них кувшины с превышенным сроком',
    'Количество стаканов': 'Из них стаканы с превышенным сроком',
    'Количество портьер': 'Из них портьеры с превышенным сроком',
    'Количество тюлей': 'Из них тюль с превышенным сроком',
    'Количество жалюзей': 'Из них жалюзи с превышенным сроком',
    'Количество ковров': 'Из них ковров с превышенным сроком',
    'Количество карт': 'Из них карт с превышенным сроком',
    'Количество ламп': 'Из них ламп с превышенным сроком',
    'Количество настольных наборов': 'Из них настольных наборов с превышенным сроком',
    'Количество часов': 'Из них часов с превышенным сроком',
    'Количество кронштейнов': 'Из них кронштейнов с превышенным сроком',
    'Количество цифровых тел.': 'Из них цифровых тел. с превышенным сроком',
    'Количество флагов': 'Из них флагов с превышенным сроком',
    'Количество карнизов': 'Из них карнизов с превышенным сроком',
    'Количество экранов защитных': 'Из них экранов защитных с превышенным сроком',
    'Количество досок': 'Из них досок с превышенным сроком',
    'Количество чайников': 'Из них чайников с превышенным сроком',
    'Количество печей': 'Из них печей с превышенным сроком',
    'Количество стеллажей': 'Из них стеллажей с превышенным сроком',
    'Количество стелажей': 'Из них стеллажей с превышенным сроком',
    'Количество стендов': 'Из них стендов с превышенным сроком',
    'Количество комодов': 'Из них комодов с превышенным сроком',
    'Количество кофемашин':  'Из них кофемашин с превышенным сроком',
    'Количество кофеварк': 'Из них кофеварок с превышенным сроком',
    'Количество СВЧ': 'Из них СВЧ с превышенным сроком'
}


choose_otdel = [
    'Административно-финансовый отдел',
    'Административный отдел',
    'Административный отдел (Б)',
    'АХО (012 каб)',
    'АХО (недвежимость)',
    'АХО (Ногинск)',
    'АХО Кожевники',
    'АХО РФН',
    'Контрольно - ревизионный отдел в сфере деятельности силовых ведомств и судебной системы',
    'Контрольно-ревизионный отдел в социально-экономической сфере',
    'Московская область, г. Озеры, ул. Ленина, д. 22',
    'Операционный отдел',
    'Организационно - аналитический отдел',
    'Отдел № 1',
    'Отдел № 10',
    'Отдел № 12',
    'Отдел № 12 (г. Мытищи)',
    'Отдел № 13',
    'Отдел № 15',
    'Отдел № 15 (г. Истра)',
    'Отдел № 18',
    'Отдел № 19',
    'Отдел № 20',
    'Отдел № 22',
    'Отдел № 23',
    'Отдел № 24',
    'Отдел № 27',
    'Отдел № 27 (г. Домодедово)',
    'Отдел № 28',
    'Отдел № 29',
    'Отдел № 3',
    'Отдел № 32',
    'Отдел № 33',
    'Отдел № 34',
    'Отдел № 34 (г. Чехов)',
    'Отдел № 35',
    'Отдел № 36',
    'Отдел № 38',
    'Отдел № 4',
    'Отдел № 4 (Воскресенск)',
    'Отдел № 4 (г. Коломна)',
    'Отдел № 41',
    'Отдел № 42',
    'Отдел № 43',
    'Отдел № 43 (г. Павловский Посад)',
    'Отдел № 5',
    'Отдел № 8',
    'Отдел № 9',
    'Отдел №18',
    'Отдел №40',
    'Отдел бюджетного учета и отчетности по операциям бюджетов',
    'Отдел ведения федеральных реестров',
    'Отдел внутреннего контроля и аудита',
    'Отдел государственной гражданской службы и кадров',
    'Отдел доходов',
    'Отдел информационных систем',
    'Отдел казначейского сопровождения',
    'Отдел кассового обслуживания исполнения бюджетов',
    'Отдел мобилизационной подготовки и гражданской обороны',
    'Отдел обслуживания силовых ведомств',
    'Отдел по надзору за аудиторской деятельностью',
    'Отдел по централизованному ведению бухгалтерского учета',
    'Отдел по централизованному начислению заработной платы и иных выплат',
    'Отдел по централизованному начислению заработной платы и иных выплат (Кашира)',
    'Отдел расходов',
    'Отдел режима секретности и безопасности информации',
    'Отдел технологического обеспечения',
    'Отдел централизованной бухгалтерии',
    'Отдел централизованной отчетности и мониторинга',
    'Руководство',
    'Руководство ( руководитель УФК и его заместители )',
    'Склад',
    'Юридический отдел',
]




spravochnik = {
'Сводный' : 'Обеспеченность',
'Административно-финансовый отдел' : 'Аппарат Управления',
'Административный отдел' : 'Аппарат Управления',
'Административный отдел (Б)' : 'Аппарат Управления',
'АХО Кожевники' : 'Аппарат Управления',
'АХО РФН' : 'Аппарат Управления',
'Операционный отдел' : 'Аппарат Управления',
'Отдел бюджетного учета и отчетности по операциям бюджетов' : 'Аппарат Управления',
'Отдел ведения федеральных реестров' : 'Аппарат Управления',
'Отдел внутреннего контроля и аудита' : 'Аппарат Управления',
'Отдел государственной гражданской службы и кадров' : 'Аппарат Управления',
'Отдел доходов' : 'Аппарат Управления',
'Отдел информационных систем' : 'Аппарат Управления',
'Отдел казначейского сопровождения' : 'Аппарат Управления',
'Отдел кассового обслуживания исполнения бюджетов' : 'Аппарат Управления',
'Отдел мобилизационной подготовки и гражданской обороны' : 'Аппарат Управления',
'Отдел обслуживания силовых ведомств' : 'Аппарат Управления',
'Отдел по централизованному начислению заработной платы и иных выплат' : 'Аппарат Управления',
'Отдел по централизованному начислению заработной платы и иных выплат (Кашира)' : 'Аппарат Управления',
'Отдел расходов' : 'Аппарат Управления',
'Отдел режима секретности и безопасности информации' : 'Аппарат Управления',
'Отдел технологического обеспечения' : 'Аппарат Управления',
'Отдел централизованной бухгалтерии' : 'Аппарат Управления',
'Отдел централизованной отчетности и мониторинга' : 'Аппарат Управления',
'Руководство' : 'Аппарат Управления',
'Руководство ( руководитель УФК и его заместители )' : 'Аппарат Управления',
'Юридический отдел' : 'Аппарат Управления',
'Склад' : 'Склад',
'Отдел № 1' : 'Отдел № 1 (г. Балашиха)',
'Отдел № 3' : 'Отдел № 3 (г. Волоколамск)',
'Отдел № 4' : 'Отдел № 4 (г. Коломна)',
'Отдел № 4 (г. Коломна)' : 'Отдел № 4 (г. Коломна)',
'Отдел № 4 (Воскресенск)' : 'УРМ Отдела № 4 (г. Воскресенск)',
'Отдел № 5' : 'Отдел № 5 (г. Дмитров)',
'Отдел № 8' : 'Отдел № 8 (г. Егорьевск)',
'Отдел № 9' : 'Отдел № 9 (г. Зарайск)',
'Отдел № 10' : 'Отдел № 10 (г. Звенигород)',
'Отдел № 12' : 'Отдел № 12 (г. Королев)',
'Отдел № 12 (г. Мытищи)' : 'УРМ Отдела № 12 (г. Мытищи)',
'Отдел № 13' : 'Отдел № 13 (г. Клин)',
'Отдел № 15' : 'Отдел № 15 (г. Красногорск)',
'Отдел № 15 (г. Истра)' : 'УРМ Отдела № 15 (г. Истра)',
'Отдел № 18' : 'Отдел № 18 (г. Луховицы)',
'Отдел №18' : 'Отдел № 18 (г. Луховицы)',
'Отдел № 19' : 'Отдел № 19 (г. Люберцы)',
'Отдел № 20' : 'Отдел № 20 (г. Можайск)',
'Отдел № 22' : 'Отдел № 22 (г. Наро-Фоминск)',
'Отдел № 23' : 'Отдел № 23 (г. Одинцово)',
'Отдел № 24' : 'Отдел № 24 (г. Озеры)',
'Отдел № 27' : 'Отдел № 27 (г. Подольск)',
'Отдел № 27 (г. Домодедово)' : 'Отдел № 27 (г. Домодедово)',
'Отдел № 28' : 'Отдел № 28 (г. Пушкино)',
'Отдел № 29' : 'Отдел № 29 (г. Раменское)',
'Отдел № 32' : 'Отдел № 32 (п.г.т. Серебряные Пруды)',
'Отдел № 33' : 'Отдел № 33 (г. Сергиев Посад)',
'Отдел № 34' : 'Отдел № 34 (г. Серпухов)',
'Отдел № 34 (г. Чехов)' : 'УРМ Отдела № 34 (г. Чехов)',
'Отдел № 35' : 'Отдел № 35 (г. Солнечногорск)',
'Отдел № 36' : 'Отдел № 36 (г. Ступино)',
'Отдел № 38' : 'Отдел № 38 (г. Химки)',
'Отдел №40' : 'Отдел № 40 (г. Шатура)',
'Отдел № 41' : 'Отдел № 41 (п. Шаховская)',
'Отдел № 42' : 'Отдел № 42 (г. Щелково)',
'Отдел № 43' : 'Отдел № 43 (г. Ногинск)',
'Отдел № 43 (г. Павловский Посад)' : 'УРМ Отдела № 43 (г. Павловский-Посад)'
}