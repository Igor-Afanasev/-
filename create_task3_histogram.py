import xlsxwriter

name='Zadanie_3_variant_14_gistogramma_s_nakopleniem.xlsx'
data = [
[1,305,172,248,715,713,270,316,341],
[2,450,153,572,610,810,348,728,534],
[3,720,168,352,345,201,439,461,495],
[4,498,174,145,504,105,140,826,189],
[5,307,125,620,312,50,170,514,294],
[6,415,147,230,601,340,95,718,116],
]
wb=xlsxwriter.Workbook(name)
wb.set_calc_mode('auto')
fmt_title=wb.add_format({'bold':True,'font_size':17,'font_color':'#17365D'})
fmt_hdr=wb.add_format({'bold':True,'bg_color':'#1F4E78','font_color':'white','align':'center','border':1})
fmt_cell=wb.add_format({'border':1,'align':'center','num_format':'0'})
fmt_input=wb.add_format({'border':1,'align':'center','bg_color':'#FFF2CC','num_format':'0'})
fmt_label=wb.add_format({'bold':True,'font_color':'#254669'})
fmt_note=wb.add_format({'font_color':'#555555','italic':True,'text_wrap':True})
fmt_total=wb.add_format({'bold':True,'num_format':'0','border':1})
src=wb.add_worksheet('Исходные данные')
src.set_column('A:A',12);src.set_column('B:I',13)
src.write('A1','Задание № 3. Выручка предприятия, млн руб.',fmt_title)
src.write('A3','Период',fmt_hdr)
headers=['1 и 11','2 и 9','3 и 13','4 и 10','5 и 16','6 и 12','7 и 14','8 и 15']
for j,h in enumerate(headers,1):src.write(2,j,h,fmt_hdr)
for i,row in enumerate(data,3):
 for j,v in enumerate(row):src.write_number(i,j,v,fmt_input if j==7 else fmt_cell)
src.write('A11','Жёлтым выделен вариант 14 (столбец H). Меняйте значения — график обновится.',fmt_note)
src.set_row(10,40);src.merge_range('A11:I11','Жёлтым выделен вариант 14 (столбец H). Меняйте значения — график обновится.',fmt_note)
src.freeze_panes(3,1)

ws=wb.add_worksheet('Гистограмма варианта 14')
ws.set_column('A:A',12);ws.set_column('B:E',19);ws.set_column('F:L',15)
ws.write('A1','Вариант 14 — изменения выручки по периодам',fmt_title)
ws.write('A2','Гистограмма с накоплением по образцу рисунка 37',fmt_note)
columns=['Период','Выручка','Невидимая основа','Рост','Снижение']
for j,h in enumerate(columns):ws.write(3,j,h,fmt_hdr)
revenues=[316,728,461,826,514,718]
for i,(period,r) in enumerate(zip(range(1,7),revenues),4):
 ws.write_number(i,0,period,fmt_cell)
 ws.write_formula(i,1,f"='Исходные данные'!H{i}",fmt_input,r)
 if i==4:
  ws.write_formula(i,2,f'=B{i+1}',fmt_cell,r)
  ws.write_formula(i,3,'=0',fmt_cell,0)
  ws.write_formula(i,4,'=0',fmt_cell,0)
 else:
  prev=revenues[i-5]
  ws.write_formula(i,2,f'=MIN(B{i},B{i+1})',fmt_cell,min(prev,r))
  ws.write_formula(i,3,f'=MAX(0,B{i+1}-B{i})',fmt_cell,max(0,r-prev))
  ws.write_formula(i,4,f'=MAX(0,B{i}-B{i+1})',fmt_cell,max(0,prev-r))
ws.write('A12','Принцип:',fmt_label)
ws.merge_range('B12:H12','Невидимая основа + рост/снижение дают плавающие столбцы, как на рисунке 37.',fmt_note)
ws.write('A13','Обновление:',fmt_label)
ws.merge_range('B13:H13','Изменяйте жёлтые значения на листе «Исходные данные» → диаграмма пересчитывается.',fmt_note)
ws.freeze_panes('B5')

chart=wb.add_chart({'type':'column','subtype':'stacked'})
def series(name,col,fill,none=False):
 conf={'name':name,'categories':['Гистограмма варианта 14',5,0,9,0],
       'values':['Гистограмма варианта 14',5,col,9,col],
       'gap':38,'border':{'none':True}}
 conf['fill']={'none':True} if none else {'color':fill}
 chart.add_series(conf)
series('Основа',2,'#FFFFFF',True)
series('Рост',3,'#4472C4')
series('Снижение',4,'#C0504D')
chart.set_title({'name':'Динамика выручки по периодам, млн руб.'})
chart.set_x_axis({'name':'Период','major_tick_mark':'none'})
chart.set_y_axis({'name':'Выручка, млн руб.','min':0,'max':900,'major_unit':100,'major_gridlines':{'visible':True}})
chart.set_legend({'none':True})
chart.set_size({'width':920,'height':560})
chart.set_chartarea({'border':{'none':True}})
chart.set_plotarea({'border':{'none':True}})
ws.insert_chart('G4',chart)

guide=wb.add_worksheet('Пояснения')
guide.set_column('A:A',22);guide.set_column('B:B',96)
guide.write('A1','Задание 3',fmt_title)
guide.write('A3','Вариант',fmt_label);guide.write('B3','14 (общая колонка 7 и 14)')
guide.write('A4','Тип',fmt_label);guide.write('B4','Гистограмма с накоплением (столбцы с невидимой базой)')
guide.write('A5','Изменения',fmt_label);guide.write('B5','Рост — синий, снижение — красный')
guide.write('A6','Редактирование',fmt_label);guide.write('B6','Введите новые значения выручки в столбец H на листе «Исходные данные»')
guide.write('A7','Проверка',fmt_label);guide.write('B7','Используется 6 периодов; первому периоду соответствует начальный уровень без столбца изменений')
guide.write('A9','Примечание',fmt_label);guide.write('B9','Рис. 37 показывает изменения от предыдущего уровня; диаграмма повторяет этот приём.')
wb.close()
print(name)
