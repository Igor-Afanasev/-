import requests,zipfile,io,json,uuid,copy,os
url='https://raw.githubusercontent.com/h5p/h5p-integration-test-suite/master/test-content/h5p-org-examples/course-presentation-08042017.h5p'
r=requests.get(url,timeout=90);r.raise_for_status()
zin=zipfile.ZipFile(io.BytesIO(r.content))
manifest=json.loads(zin.read('h5p.json'))
original=json.loads(zin.read('content/content.json'))
manifest['title']='Гидросфера — водная оболочка Земли. Интерактивная презентация'
manifest['language']='ru'
manifest['author']='Учебная презентация'
manifest['license']='CC BY'
manifest['meta']={'title':manifest['title'],'license':'CC BY'}
# Reuse the tested content-type examples from the source H5P package
original_slide=original['presentation']['slides']
text_el=next(el for el in original_slide[0]['elements'] if el.get('action',{}).get('library','').startswith('H5P.AdvancedText'))
quiz_el=original_slide[1]['elements'][0]
assert quiz_el['action']['library'].startswith('H5P.SingleChoiceSet')

questions=[
('Что называют гидросферой?',['Водную оболочку Земли','Воздушную оболочку Земли','Твёрдую оболочку Земли','Оболочку живых организмов'],'Гидросфера включает все воды Земли: океаны, моря, воды суши, лёд, подземные воды и атмосферную влагу.'),
('Какая часть гидросферы содержит наибольшее количество воды?',['Мировой океан','Реки','Озёра','Атмосфера'],'Основной объём земной воды сосредоточен в морях и океанах.'),
('Какова приблизительная доля солёной воды на Земле?',['Около 97%','Около 25%','Около 50%','Около 3%'],'Примерно 97,5% воды Земли солёная; пресной воды около 2,5%.'),
('Где находится большая часть пресной воды Земли?',['В ледниках и ледяных покровах','В реках','В атмосфере','В пресноводных озёрах'],'Большая часть пресной воды заключена в ледниковых покровах и ледниках.'),
('Что служит главным источником энергии круговорота воды?',['Солнечная энергия','Магнитное поле','Приливы','Энергия земного ядра'],'Солнечное тепло обеспечивает испарение воды и запускает её круговорот.'),
('Как называется переход воды из газообразного состояния в жидкое?',['Конденсация','Испарение','Таяние','Сублимация'],'При охлаждении водяной пар конденсируется, образуя капли облаков и тумана.'),
('Какая вода относится к подземным водам?',['Вода в порах и трещинах горных пород','Вода на поверхности океана','Вода в облаках','Лёд айсберга'],'Подземные воды расположены в верхних слоях земной коры в пустотах и трещинах пород.'),
('Какое течение считается тёплым в Атлантическом океане?',['Гольфстрим','Лабрадорское','Канарское','Бенгельское'],'Гольфстрим переносит тёплую воду в Северной Атлантике.'),
('Что происходит с уровнем Мирового океана при таянии материковых ледников?',['Он повышается','Он всегда понижается','Он не может меняться','Океан становится пресным'],'Таяние льда, находящегося на суше, добавляет воду в океан и способствует повышению уровня моря.'),
('Какое действие помогает сохранять водные ресурсы?',['Очистка сточных вод','Сброс мусора в реки','Неограниченный расход воды','Уничтожение прибрежных растений'],'Очистка сточных вод и рациональное потребление уменьшают загрязнение и потери воды.')
]
assert len(questions)==10

def text_element(html,x=4,y=5,w=92,h=38):
    el=copy.deepcopy(text_el)
    el['x']=x;el['y']=y;el['width']=w;el['height']=h
    el['action']['params']={'text':html}
    el['action']['subContentId']=str(uuid.uuid4())
    return el

def slide(elements):
    return {'elements':elements,'slideBackgroundSelector':{'fillSlideBackground':False},'keywords':[]}
slides=[]
slides.append(slide([
 text_element('<h2>Гидросфера — водная оболочка Земли</h2><p>Интерактивный урок по географии.</p><p>В презентации — 10 вопросов. После каждого ответа можно проверить результат, посмотреть правильный вариант и повторить попытку.</p>',4,5,92,58),
 text_element('<p><strong>Темы:</strong> состав гидросферы, запасы воды, Мировой океан, круговорот воды, подземные воды, течения и охрана водных ресурсов.</p>',4,69,90,20)
]))
for i,(question,answers,explanation) in enumerate(questions,1):
    qel=copy.deepcopy(quiz_el)
    qel['x']=5;qel['y']=42;qel['width']=90;qel['height']=52
    qel['action']['params']['choices']=[{'answers':answers,'question':'<p>'+question+'</p>','subContentId':str(uuid.uuid4())}]
    qel['action']['params']['behaviour'].update({'enableRetry':True,'enableSolutionsButton':True,'passPercentage':100,'soundEffectsEnabled':False})
    qel['action']['params']['l10n'].update({'showSolutionButtonLabel':'Показать ответ','retryButtonLabel':'Повторить','solutionViewTitle':'Правильный ответ','correctText':'Верно!','incorrectText':'Неверно','resultSlideTitle':'Правильных ответов: :numcorrect из :maxscore','slideOfTotal':'Слайд :num из :total'})
    qel['action']['subContentId']=str(uuid.uuid4())
    header=text_element(f'<h2>Вопрос {i} из 10</h2><p><strong>Гидросфера — водная оболочка Земли</strong></p>',5,3,89,23)
    hint=text_element('<p><em>Подсказка после ответа:</em> '+explanation+'</p>',5,27,89,15)
    slides.append(slide([header,hint,qel]))
slides.append(slide([text_element('<h2>Итоги урока</h2><p>Вы познакомились с составом гидросферы, распределением воды, её круговоротом и важностью охраны водных ресурсов.</p><p><strong>Повторите вопросы, на которые ответили неверно.</strong></p><p>Помните: вода необходима всем живым организмам, поэтому её нужно беречь.</p>',5,8,90,76)]))
original['presentation']['slides']=slides
original['presentation']['keywordListEnabled']=False
out='Gidrosfera_interaktivnaya_prezentaciya_10_voprosov.h5p'
with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED,compresslevel=7) as zout:
    for item in zin.infolist():
        if item.filename=='content/content.json':
            zout.writestr('content/content.json',json.dumps(original,ensure_ascii=False,separators=(',',':')))
        elif item.filename=='h5p.json':
            zout.writestr('h5p.json',json.dumps(manifest,ensure_ascii=False,separators=(',',':')))
        else:
            zout.writestr(item,zin.read(item.filename))
with zipfile.ZipFile(out) as check:
    c=json.loads(check.read('content/content.json'))
    assert len(c['presentation']['slides'])==12
    assert sum(1 for s in c['presentation']['slides'] for e in s['elements'] if 'SingleChoiceSet' in e.get('action',{}).get('library',''))==10
    assert check.testzip() is None
print('Created',out,'size:',os.stat(out).st_size,'slides',len(slides))
