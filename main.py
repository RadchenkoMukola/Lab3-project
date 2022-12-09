from openpyxl import load_workbook
from openpyxl.styles import Color, PatternFill, Font, Border
from openpyxl.styles import colors
from openpyxl.cell import Cell
import PySimpleGUI as sg


def input_cells():
    return [
        [sg.Text(item, size=size5, pad=(0, 0)),
         sg.Input("",size=size2, enable_events=True, key=item, expand_x=True)]
            for item in input_items]


def fk_cells():
    return [
        [sg.Text(item, size=size2, pad=(0, 0)),
         sg.Input("",size = size3, enable_events=True, key=item, expand_x=True)]
            for item in fk_items]


def sn_cells():
    return [
        [sg.Text(item, size=size1, pad=(0, 0)),
         sg.Input("",size = size3, enable_events=True, key=item, expand_x=True)]
            for item in sn_items]


def cat_cells():
 return [
  [sg.Text(item, size=5, pad=(0, 0)),
   sg.Input("", size=size3, enable_events=True, key=item, expand_x=True)]
  for item in cat_items]


def gem_cells():
 return [
  [sg.Text(item, size=18, pad=(0, 0)),
   sg.Input("", size=3, enable_events=True, key=item, expand_x=True)]
  for item in gem_items]


def at_cells():
 return [
  [sg.Text(item, size=18, pad=(0, 0)),
   sg.Input("", size=size3, enable_events=True, key=item, expand_x=True)]
  for item in at_items]


def post_cells():
 return [
  [sg.Text(item, size=18, pad=(0, 0)),
   sg.Input("", size=size3, enable_events=True, key=item, expand_x=True)]
  for item in post_items]


def eco_cells():
 return [
  [sg.Text(item, size=20, pad=(0, 0)),
   sg.Input("", size=15, enable_events=True, key=item, expand_x=True)]
  for item in eco_items]


def kat_cells():
 return [
  [sg.Text(item, size=25, pad=(0, 0)),
   sg.Input("", size=3, enable_events=True, key=item, expand_x=True)]
  for item in kat_items]


def gas_cells():
 return [
  [sg.Text(item, size=15, pad=(0, 0)),
   sg.Input("", size=5, enable_events=True, key=item, expand_x=True)]
  for item in gas_items]


def sin_cells():
 return [
  [sg.Text(item, size=15, pad=(0, 0)),
   sg.Input("", size=3, enable_events=True, key=item, expand_x=True)]
  for item in sin_items]


def kt_cells():
 return [
  [sg.Text(item, size=size2, pad=(0, 0)),
   sg.Input("", size=15, enable_events=True, key=item, expand_x=True)]
  for item in kt_items]


def lik_cells():
 return [
  [sg.Text(item, size=35, pad=(0, 0)),
   sg.Input("", size=3, enable_events=True, key=item, expand_x=True)]
  for item in lik_items]


def prm_cells():
 return [
  [sg.Text(item, size=50, pad=(0, 0)),
   sg.Input("", size=30, enable_events=True, key=item, expand_x=True)]
  for item in prm_items]

size1,size2,size3,size4,size5 = 40,10,2,12,32

input_items  = ('Рік', 'Номер','Госпіталізовані -1 чи амбулаторні -0','№ ІХ','ПІБ','Пацієнт первинний -1, повторний -0','Дата народження','Вік','Місце проживання','Телефон','Лікар','Стать ч-1 ж-0','Зріст(см)','Вага п','Вага в','Скарги, міс','Діагностована ЛГ,міс')
fk_items = ('ЛГ ','1.1.','1.2.','1.3.','1.4.1.','1.4.2.','1.4.3.','1.4.4.','1.4.5.','1’','2','3','4.1.','4.2.','5','І ст.','ІІ ст.','ІІІ ст.','I ФК ВОЗ','IІ ФК ВОЗ','IІІ ФК ВОЗ','IV ФК ВОЗ',)
sn_items = ('СН І','СН ІІА','СН ІІБ','ФП','ТП','Асцит','Гідроперикард','Плеврит, гідроторакс','ХОЗЛ','СОАС','Індекс AHI','Мінімальна сатурація','ГПМК','Анемія','Залізодефіцитний стан','Гіпотиреоз 1, гіпертиреоз 2, еутиреоз 3','ВВС','Після операції з приводу вади','ДМПП','ДМШП','ВАП','ІНШІ ВВС','ЗКИД ЛП 0, ПЛ 1, перехрестний 2','Тромбоутворення в н/кінцівках 1-є, 0- відсутне','Кровотечі','Синкопе')
cat_items = ('САТ п','ДАТ п','ЧСС п','САТ в','ДАТ в','ЧСС в')
gem_items = ('Гемоглобін','Еритроцити','Гематокрит','MCV','MCH','Тромбоцити','Лейкоцити','ШОЄ','МНО','Заг.хол.','ТГ','К','білірубін','креатинін','кліренс креатиніну','Сечова кислота','АЛТ ','АСТ','глюкоза','HBsAg','Anti-HCV','ВИЧ','Ntpro-BNP (ДІЛА)','Ntpro-BNP (Інститут)','ТТГ','Феритин','Anti-ENA скринінг','Результат:0-нег, 1-поз')
at_items =('Цент АТ','СРПХ м','СРПХ е','R-CAVI','L-CAVI','R-ABI','L-ABI')
post_items = ('6ХХ м пост.','6ХХ бали','SpO2 до','SpO2 після','6ХХ м вип.','6ХХ бали','SpO2 до','SpO2 після')
eco_items = ('ЕхоКс','Аорта','ЛП','S ЛП','Інд.ЛП','S ПП','Інд.ПП','МШП','ЗСЛШ','КДР ЛШ','КСР ЛШ','КДО ЛШ','КСО ЛШ','УО ЛШ','ФВ','ММ ЛШ','E/A МК','DT МК','БС ПШ/ЛШ','Стінка ПШ','S ПШ діаст','S ПШ сист','Фракц.скор.ПШ','Позд.розПШ','Попер.розПШ','S’','TAPSE мм','Швидкість рег. на ТСК ','Т АСС ','ЛА','НПВ','Колаб{1>50,25<(0,5)<50,0<25}','Сист.ТЛА','Сер.ТЛА','ІЕ діастола','ІЕ систола','АК нед-ть','МК нед-ть','ТК нед-ть','ЛК нед-ть','ЙЛГ 0-низ, 1-сер, 2-вис','Висновок')
kat_items = ('КПС Інс-1,поза-2,не роб-0','Дата кат','Сист.Т.ЛА','Діас.Т.ЛА','сер ТЛА','Сист.Т.ПШ ПШ','Діас.Т.ПШ','сер ТПШ','Сист.Т. ПП','Діас.Т.ПП','сер ПП','ТЗЛА','ХОК','Серц.індекс','Удар.викид','Удар.індекс','PVR','PVR Wood','ЧСС','САТ','ДАТ','Сер АТ','SVR','TPR','ГДД','LVSW ','LVSWI','RVSW','RVSWI','Вазореак.тест','СТЛА п т','ДТЛА п т','Сер тиск ЛА п т','ТЗЛА','ХОК','Серц.інд','Удар.викид','Удар.інд','PVR','ЧСС','САТ','ДАТ','сер АТ','Результат','Проба з піднятими н/к','СТЛА','ДТЛА','СерТЛА','ТЗЛА','Кров ЛА')
gas_items = ('pH','pO2','pCO2','Hct','Na','K','Ca','THb','HCO3','SvO2','BE в','Р50 смеш вен','Кров артер','pH','pO2','pCO2','Hct','Na','K','Ca','THb','HCO3','SaO2','BE в','Р50 арт','Оцінка ризику')
sin_items = ('Спирометрія','ФЖЕЛ','ОФВ1','ОФВ1/ФЖЕЛ','ЖЕЛ','ОФВ1/ЖЕЛ','Хвилинне споживання кисню','ДО','ЛН 1-рестриктивний тип, 2- обструктивний','БодіПГ','TLC','ДифЗдЛегень','DLCO')
kt_items = ('КТ, МРТ ОГК','Дата','Висновок')
lik_items = ('Лік:1-д,1`-н,2-а,3-сіл,4-т,5-в,6-б,7-амб,8-р,9-с','Фурос.в/в','Фурос.в/м','Дофамін','Сілденафіл','Доза','Тадалафіл','Доза','Вентавіс','Доза мкг','АК 1-а, 2-д, 3-н, 4-вер','Доза','Фуросемід','Доза','Ант. Альд.  1 Верошпірон, 2 -еплеренон','Доза','Сер.глік.','Доза','Небіволол 1, карведілол 2','Доза','ЕРА:Б-1,а-2','Доза','ОАК:вар-1, кс-2, сінк-3,прад-4','Доза','Тріфас','Доза','Кордарон','Доза','АСК','Залізо 1 в/м, 2 per os, 3 в/в','Оксигенотерапія','Оптимальна Терапія')
prm_items = ('Примітки','ТЕА 0-реком,1-пров,2-відм.БАП 3-реком,4-пров, 5-відм 6-не реком','Дата','Рекомендации Кулика','Декомпенсація','Дата','Трансплантація 1-проведена, 0 -рекомендована','Дата','Смерть','Дата')
sg.theme('SandyBeach')
sg.SetOptions(font = ('Arial', 10, 'bold'))

layout = [
    [sg.Frame('Основні дані',input_cells(),vertical_alignment='top', expand_y=True, expand_x=True),sg.Column(fk_cells(),vertical_scroll_only=True,scrollable=True,  vertical_alignment='top', expand_y=True, expand_x=True),sg.Column(sn_cells(),vertical_scroll_only=True,scrollable=True,  vertical_alignment='top', expand_y=True, expand_x=True),sg.Column(gem_cells(),vertical_scroll_only=True,scrollable=True,  vertical_alignment='top', expand_y=True, expand_x=True),sg.Column(eco_cells(),vertical_scroll_only=True,scrollable=True,vertical_alignment='top', expand_y=True, expand_x=True),sg.Column(kat_cells(),vertical_scroll_only=True,scrollable=True,vertical_alignment='top', expand_y=True, expand_x=True),sg.Column(gas_cells(),vertical_scroll_only=True,scrollable=True,vertical_alignment='top', expand_y=True, expand_x=True)]
 ,[sg.Column(cat_cells(), vertical_scroll_only=True,scrollable=True,  vertical_alignment='top', expand_y=True, expand_x=True),sg.Column(at_cells(),vertical_scroll_only=True,scrollable=True,  vertical_alignment='top', expand_y=True, expand_x=True),sg.Column(post_cells(),vertical_scroll_only=True,scrollable=True,vertical_alignment='top', expand_y=True, expand_x=True),sg.Column(sin_cells(),vertical_scroll_only=True,scrollable=True,vertical_alignment='top', expand_y=True, expand_x=True),sg.Column(kt_cells(), vertical_scroll_only=True,scrollable=True,  vertical_alignment='top', expand_y=True, expand_x=True),sg.Column(lik_cells(), vertical_scroll_only=True,scrollable=True,  vertical_alignment='top', expand_y=True, expand_x=True),sg.Column(prm_cells(), vertical_scroll_only=True,scrollable=True,  vertical_alignment='top', expand_y=True, expand_x=True),[sg.Button('Додати',button_color='green', key='calcA'),sg.Button('Очистити',button_color='blue', key='clear')]]]


window = sg.Window('Exelfiller', layout,resizable=True)


while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED:
        break
    elif event == 'calcA':

     for i in values:
      if str(values[i]) == "0,5":
       values[i] = 0.5

     #1 'Рік', 'Номер','Госпіталізовані -1 чи амбулаторні -0','№ ІХ','ПІБ','Пацієнт первинний -1, повторний -0','Дата народження','Вік','Місце проживання','Телефон','Лікар','Стать ч-1 ж-0','Зріст(см)','Вага п','Вага в','Скарги, міс','Діагностована ЛГ,міс'
     Rik = None if str((values['Рік'])) == "" else float(values['Рік'])
     Number = None if str((values['Номер'])) == "" else float(values['Номер'])
     Gos = None if str((values['Госпіталізовані -1 чи амбулаторні -0'])) == ""  else float(values['Госпіталізовані -1 чи амбулаторні -0'])
     NoIX =None if str((values['№ ІХ'])) == "" else float(values['№ ІХ'])
     PIB = str(values['ПІБ'])
     Per = None if str((values['Пацієнт первинний -1, повторний -0'])) == "" else float(values['Пацієнт первинний -1, повторний -0'])
     Birth = str(values['Дата народження'])
     Vik = None if str((values['Вік'])) == "" else float(values['Вік'])
     Mis = str(values['Місце проживання'])
     Pho = str(values['Телефон'])
     Doc = str(values['Лікар'])
     Sex = None if str((values['Стать ч-1 ж-0'])) == "" else  float(values['Стать ч-1 ж-0'])
     weightp = None if str((values['Вага п'])) == "" else float(values['Вага п'])
     weightv = None if str((values['Вага в'])) == "" else float(values['Вага в'])
     height = None if str((values['Зріст(см)'])) == "" else float(values['Зріст(см)'])
     Scarm = None if str((values['Скарги, міс'])) == "" else float(values['Скарги, міс'])
     Diag = None if str((values['Діагностована ЛГ,міс'])) == "" else float(values['Діагностована ЛГ,міс'])
     #2 'ЛГ ','1.1.','1.2.','1.3.','1.4.1.','1.4.2.','1.4.3.','1.4.4.','1.4.5.','1’','2','3','4.1.','4.2.','5','І ст.','ІІ ст.','ІІІ ст.','I ФК ВОЗ','IІ ФК ВОЗ','IІІ ФК ВОЗ','IV ФК ВОЗ'

     Lg = None if str((values['ЛГ '])) == "" else float(values['ЛГ '])
     N11 = None if str((values['1.1.'])) == "" else float(values['1.1.'])
     N12 = None if str((values['1.2.'])) == "" else float(values['1.2.'])
     N13 = None if str((values['1.3.'])) == "" else float(values['1.3.'])
     N141 = None if str((values['1.4.1.'])) == "" else float(values['1.4.1.'])
     N142 = None if str((values['1.4.2.'])) == "" else float(values['1.4.2.'])
     N143 = None if str((values['1.4.3.'])) == "" else float(values['1.4.3.'])
     N144 = None if str((values['1.4.4.'])) == "" else float(values['1.4.4.'])
     N145 = None if str((values['1.4.5.'])) == "" else float(values['1.4.5.'])
     N1 = None if str((values['1’'])) == "" else float(values['1’'])
     N2 = None if str((values['2'])) == "" else float(values['2'])
     N3 = None if str((values['3'])) == "" else float(values['3'])
     N41 = None if str((values['4.1.'])) == "" else float(values['4.1.'])
     N42 = None if str((values['4.2.'])) == "" else float(values['4.2.'])
     N5 = None if str((values['5'])) == "" else float(values['5'])
     St1 = None if str((values['І ст.'])) == "" else float(values['І ст.'])
     St2 = None if str((values['ІІ ст.'])) == "" else float(values['ІІ ст.'])
     St3 = None if str((values['ІІІ ст.'])) == "" else float(values['ІІІ ст.'])
     Fk1 = None if str((values['I ФК ВОЗ'])) == "" else float(values['I ФК ВОЗ'])
     Fk2 = None if str((values['IІ ФК ВОЗ'])) == "" else float(values['IІ ФК ВОЗ'])
     Fk3 = None if str((values['IІІ ФК ВОЗ'])) == "" else float(values['IІІ ФК ВОЗ'])
     Fk4 = None if str((values['IV ФК ВОЗ'])) == "" else float(values['IV ФК ВОЗ'])

     # 3 'СН І','СН ІІА','СН ІІБ','ФП','ТП','Асцит','Гідроперикард','Плеврит, гідроторакс','ХОЗЛ','СОАС','Індекс AHI','Мінімальна сатурація','ГПМК','Анемія','Залізодефіцитний стан','Гіпотиреоз 1, гіпертиреоз 2, еутиреоз 3','ВВС','Після операції з приводу вади','ДМПП','ДМШП','ВАП','ІНШІ ВВС','ЗКИД ЛП 0, ПЛ 1, перехрестний 2','Тромбоутворення в н/кінцівках 1-є, 0- відсутне','Кровотечі','Синкопе'
     Sn1 = None if str((values['СН І'])) == "" else float(values['СН І'])
     Sn2a = None if str((values['СН ІІА'])) == "" else float(values['СН ІІА'])
     Sn2b = None if str((values['СН ІІБ'])) == "" else float(values['СН ІІБ'])
     Fp = None if str((values['ФП'])) == "" else float(values['ФП'])
     Tp = None if str((values['ТП'])) == "" else float(values['ТП'])
     As = None if str((values['Асцит'])) == "" else float(values['Асцит'])
     Gip = None if str((values['Гідроперикард'])) == "" else float(values['Гідроперикард'])
     Plv = None if str((values['Плеврит, гідроторакс'])) == "" else float(values['Плеврит, гідроторакс'])
     Hoz = None if str((values['ХОЗЛ'])) == "" else float(values['ХОЗЛ'])
     Soa = None if str((values['СОАС'])) == "" else float(values['СОАС'])
     ANI = None if str((values['Індекс AHI'])) == "" else float(values['Індекс AHI'])
     Mins = None if str((values['Мінімальна сатурація'])) == "" else float(values['Мінімальна сатурація'])
     Gpmk = None if str((values['ГПМК'])) == "" else float(values['ГПМК'])
     Anm = None if str((values['Анемія'])) == "" else float(values['Анемія'])
     Zald = None if str((values['Залізодефіцитний стан'])) == "" else float(values['Залізодефіцитний стан'])
     Gp = None if str((values['Гіпотиреоз 1, гіпертиреоз 2, еутиреоз 3'])) == "" else float(values['Гіпотиреоз 1, гіпертиреоз 2, еутиреоз 3'])
     BBC = None if str((values['ВВС'])) == "" else float(values['ВВС'])
     Pis = None if str((values['Після операції з приводу вади'])) == "" else float(values['Після операції з приводу вади'])
     DMP = None if str((values['ДМПП'])) == "" else float(values['ДМПП'])
     DSP = None if str((values['ДМШП'])) == "" else float(values['ДМШП'])
     Vap = None if str((values['ВАП'])) == "" else float(values['ВАП'])
     IBBC = None if str((values['ІНШІ ВВС'])) == "" else float(values['ІНШІ ВВС'])
     Zkd = None if str((values['ЗКИД ЛП 0, ПЛ 1, перехрестний 2'])) == "" else float(values['ЗКИД ЛП 0, ПЛ 1, перехрестний 2'])
     Tmb = None if str((values['Тромбоутворення в н/кінцівках 1-є, 0- відсутне'])) == "" else float(values['Тромбоутворення в н/кінцівках 1-є, 0- відсутне'])
     Krv = None if str((values['Кровотечі'])) == "" else float(values['Кровотечі'])
     Sin = None if str((values['Синкопе'])) == "" else float(values['Синкопе'])

     #4 'САТ п','ДАТ п','ЧСС п','САТ в','ДАТ в','ЧСС в'
     Catp = None if str((values['САТ п'])) == "" else float(values['САТ п'])
     Datp = None if str((values['ДАТ п'])) == "" else float(values['ДАТ п'])
     Csp = None if str((values['ЧСС п'])) == "" else float(values['ЧСС п'])
     Catv = None if str((values['САТ в'])) == "" else float(values['САТ в'])
     Datv = None if str((values['ДАТ в'])) == "" else float(values['ДАТ в'])
     Csv = None if str((values['ЧСС в'])) == "" else float(values['ЧСС в'])

     #5'Гемоглобін','Еритроцити','Гематокрит','MCV','MCH','Тромбоцити','Лейкоцити','ШОЄ','МНО','Заг.хол.','ТГ','К','білірубін','креатинін','кліренс креатиніну','Сечова кислота','АЛТ ','АСТ','глюкоза','HBsAg','Anti-HCV','ВИЧ','Ntpro-BNP (ДІЛА)','Ntpro-BNP (Інститут)','ТТГ','Феритин','Anti-ENA скринінг','Результат: 0- негативний, 1 - позитивний'
     Gem = None if str((values['Гемоглобін'])) == "" else str(values['Гемоглобін'])
     Ert = None if str((values['Еритроцити'])) == "" else str(values['Еритроцити'])
     MCV = None if str((values['MCV'])) == "" else str(values['MCV'])
     MCH = None if str((values['MCH'])) == "" else str(values['MCH'])
     Trobm = None if str((values['Тромбоцити'])) == "" else str(values['Тромбоцити'])
     Lek = None if str((values['Лейкоцити'])) == "" else str(values['Лейкоцити'])
     Shoe = None if str((values['ШОЄ'])) == "" else str(values['ШОЄ'])
     MNO = None if str((values['МНО'])) == "" else str(values['МНО'])
     Zag = None if str((values['Заг.хол.'])) == "" else str(values['Заг.хол.'])
     Tg = None if str((values['ТГ'])) == "" else str(values['ТГ'])
     K1 = None if str((values['К'])) == "" else str(values['К'])
     Bil = None if str((values['білірубін'])) == "" else str(values['білірубін'])
     Cre = None if str((values['креатинін'])) == "" else str(values['креатинін'])
     Cli = None if str((values['кліренс креатиніну'])) == "" else str(values['кліренс креатиніну'])
     Sech = None if str((values['Сечова кислота'])) == "" else str(values['Сечова кислота'])
     Alt = None if str((values['АЛТ '])) == "" else str(values['АЛТ '])
     Ast = None if str((values['АСТ'])) == "" else str(values['АСТ'])
     Glu = None if str((values['глюкоза'])) == "" else str(values['глюкоза'])
     HBsAg = None if str((values['HBsAg'])) == "" else str(values['HBsAg'])
     AHCV = None if str((values['Anti-HCV'])) == "" else str(values['Anti-HCV'])
     Vich = None if str((values['ВИЧ'])) == "" else str(values['ВИЧ'])
     Ntpro1 = None if str((values['Ntpro-BNP (ДІЛА)'])) == "" else str(values['Ntpro-BNP (ДІЛА)'])
     Ntpro2 = None if str((values['Ntpro-BNP (Інститут)'])) == "" else str(values['Ntpro-BNP (Інститут)'])
     TTG = None if str((values['ТТГ'])) == "" else str(values['ТТГ'])
     Fer = None if str((values['Феритин'])) == "" else str(values['Феритин'])
     AENA = None if str((values['Anti-ENA скринінг'])) == "" else str(values['Anti-ENA скринінг'])
     Rez = None if str((values['Результат:0-нег, 1-поз'])) == "" else str(values['Результат:0-нег, 1-поз'])

     #6'Цент АТ','СРПХ м','СРПХ е','R-CAVI','L-CAVI','R-ABI','L-ABI'
     AENA = None if str((values['Anti-ENA скринінг'])) == "" else str(values['Anti-ENA скринінг'])
     AENA = None if str((values['Anti-ENA скринінг'])) == "" else str(values['Anti-ENA скринінг'])
     AENA = None if str((values['Anti-ENA скринінг'])) == "" else str(values['Anti-ENA скринінг'])
     AENA = None if str((values['Anti-ENA скринінг'])) == "" else str(values['Anti-ENA скринінг'])


     #7'6ХХ м пост.','6ХХ бали','SpO2 до','SpO2 після','6ХХ м вип.','6ХХ бали','SpO2 до','SpO2 після'

     #8'ЕхоКс','Аорта','ЛП','S ЛП','Інд.ЛП','S ПП','Інд.ПП','МШП','ЗСЛШ','КДР ЛШ','КСР ЛШ','КДО ЛШ','КСО ЛШ','УО ЛШ','ФВ','ММ ЛШ','E/A МК','DT МК','БС ПШ/ЛШ','Стінка ПШ','S ПШ діаст','S ПШ сист','Фракц.скор.ПШ','Позд.розПШ','Попер.розПШ','S’','TAPSE мм','Швидкість рег. на ТСК ','Т АСС ','ЛА','НПВ','Колаб{1>50,25<(0,5)<50,0<25}','Сист.ТЛА','Сер.ТЛА','ІЕ діастола','ІЕ систола','АК нед-ть','МК нед-ть','ТК нед-ть','ЛК нед-ть','ЙЛГ 0-низ, 1-сер, 2-вис','Висновок'

     #9КПС Інс-1,поза-2,не роб-0','Дата кат','Сист.Т.ЛА','Діас.Т.ЛА','сер ТЛА','Сист.Т.ПШ ПШ','Діас.Т.ПШ','сер ТПШ','Сист.Т. ПП','Діас.Т.ПП','сер ПП','ТЗЛА','ХОК','Серц.індекс','Удар.викид','Удар.індекс','PVR','PVR Wood','ЧСС','САТ','ДАТ','Сер АТ','SVR','TPR','ГДД','LVSW ','LVSWI','RVSW','RVSWI','Вазореак.тест','СТЛА п т','ДТЛА п т','Сер тиск ЛА п т','ТЗЛА','ХОК','Серц.інд','Удар.викид','Удар.інд','PVR','ЧСС','САТ','ДАТ','сер АТ','Результат','Проба з піднятими н/к','СТЛА','ДТЛА','СерТЛА','ТЗЛА','Кров ЛА'

     #10'pH','pO2','pCO2','Hct','Na','K','Ca','THb','HCO3','SvO2','BE в','Р50 смеш вен','Кров артер','pH','pO2','pCO2','Hct','Na','K','Ca','THb','HCO3','SaO2','BE в','Р50 арт','Оцінка ризику'

     #11'Спирометрія','ФЖЕЛ','ОФВ1','ОФВ1/ФЖЕЛ','ЖЕЛ','ОФВ1/ЖЕЛ','Хвилинне споживання кисню','ДО','ЛН 1-рестриктивний тип, 2- обструктивний','БодіПГ','TLC','ДифЗдЛегень','DLCO'

     #12'КТ, МРТ ОГК','Дата','Висновок'

     #13'Лік:1-д,1`-н,2-а,3-сіл,4-т,5-в,6-б,7-амб,8-р,9-с','Фурос.в/в','Фурос.в/м','Дофамін','Сілденафіл','Доза','Тадалафіл','Доза','Вентавіс','Доза мкг','АК 1-а, 2-д, 3-н, 4-вер','Доза','Фуросемід','Доза','Ант. Альд.  1 Верошпірон, 2 -еплеренон','Доза','Сер.глік.','Доза','Небіволол 1, карведілол 2','Доза','ЕРА:Б-1,а-2','Доза','ОАК:вар-1, кс-2, сінк-3,прад-4','Доза','Тріфас','Доза','Кордарон','Доза','АСК','Залізо 1 в/м, 2 per os, 3 в/в','Оксигенотерапія','Оптимальна Терапія'

     #14'Примітки','ТЕА 0-реком,1-пров,2-відм.БАП 3-реком,4-пров, 5-відм 6-не реком','Дата','Рекомендации Кулика','Декомпенсація','Дата','Трансплантація 1-проведена, 0 -рекомендована','Дата','Смерть','Дата'

     workbook = load_workbook(filename="patients.xlsx")
     sheet = workbook.active

     if weightp != None and height != None:
        BMI = float(weightp) / ((float(height) / 100 * float(height) / 100))
     else: BMI = None

     data = [Rik,Number,Gos,NoIX,PIB,Per,Birth,Vik,Mis,Pho,Doc,Sex,height,weightp,weightv,BMI,Scarm,Diag,Lg,N11,N12,N13,N141,N142,N143,N144,N145,N1,N2,N3,N41,N42,N5,St1,St2,St3,Fk1,Fk2,Fk3,Fk4,Sn1,Sn2a,Sn2b,Fp,Tp,As,Gip,Plv,Hoz,Soa,ANI,Mins,Gpmk,Anm,Zald,Gp,BBC,Pis,DMP,DSP,Vap,IBBC,Zkd,Tmb,Krv,Sin,Catp,Datp,Csp,Catv,Datv,Csv,Gem,Ert,MCV,MCH,Trobm,Lek,Shoe,MNO,Zag,Tg,K1,Bil,Cre,Cli,Sech,Alt,Ast,Glu,HBsAg,AHCV,Vich,Ntpro1,Ntpro2,TTG,Fer,AENA,Rez]
     sheet.append(data)
     workbook.save(filename="patients.xlsx")

    elif event == 'clear':
        for item in values:
            window[item].update("")

window.close()
