from pathlib import Path
from xml.etree import ElementTree as ET
from collections import Counter
import re

source = Path('Moodle_Gidrosfera_bank_27_questions.xml')
root=ET.parse(source).getroot()
questions=[q for q in root.findall('question') if q.get('type')!='category']
special={'=','~','#','{','}',':'}
def esc(s):
    s=str(s or '').replace('\\','\\\\')
    for c in special: s=s.replace(c,'\\'+c)
    return s.replace('\n',' ').strip()
def plain(s):
    s=re.sub(r'<[^>]*>',' ',str(s or ''))
    return re.sub(r'\s+',' ',s).strip()
def question_title(q):
    return esc(q.findtext('./name/text') or 'Без названия')
def stem(q):
    return esc(plain(q.findtext('./questiontext/text')))
def answers(q):
    return [(a,float(a.get('fraction','0')),a.findtext('text') or '') for a in q.findall('answer')]
def block(title,statement,answer):
    return f'::{title}::{statement} {{\n{answer}\n}}\n\n'
pieces=['// Вопросы по теме «Гидросфера — водная оболочка Земли»\n',
        '// Формат Moodle GIFT, кодировка UTF-8\n',
        '$CATEGORY: Гидросфера — водная оболочка Земли\n\n']
n=0; skipped=[]
for q in questions:
    kind=q.get('type'); title=question_title(q); prompt=stem(q)
    if kind=='description':
        pieces.append('// '+prompt+'\n\n');continue
    if kind=='multichoice':
        arr=answers(q); ismulti=(q.findtext('single','true')=='false')
        lines=[]
        if ismulti:
            for a,frac,label in arr:
                score=('33.33333' if frac>0 else '-100')
                lines.append('~%'+score+'%'+esc(label))
        else:
            for a,frac,label in arr:
                lines.append(('=' if frac>0 else '~')+esc(label))
        pieces.append(block(title,prompt,'\n'.join(lines)))
    elif kind=='truefalse':
        a=answers(q)
        istrue=next((frac>0 for _,frac,label in a if label=='true'),False)
        pieces.append(block(title,prompt,'TRUE' if istrue else 'FALSE'))
    elif kind=='shortanswer':
        good=[esc(label) for _,frac,label in answers(q) if frac>0 and label!='*']
        pieces.append(block(title,prompt,'\n'.join('='+s for s in good)))
    elif kind=='numerical':
        a=next((a for a in q.findall('answer') if float(a.get('fraction','0'))>0),None)
        num=a.findtext('text');tol=a.findtext('tolerance','0')
        pieces.append(block(title,prompt,'#'+num+(':'+tol if float(tol)>0 else '')))
    elif kind=='matching':
        pairs=[]
        for sub in q.findall('subquestion'):
            left=esc(plain(sub.findtext('text')))
            right=esc(sub.findtext('./answer/text'))
            pairs.append('='+left+' -> '+right)
        pieces.append(block(title,prompt,'\n'.join(pairs)))
    elif kind=='essay':
        pieces.append(block(title,prompt,''))
    elif kind=='cloze':
        # The original Cloze stems embed Moodle XML cloze syntax. Convert to GIFT-compatible
        # distinct single-answer and short-answer questions, preserving underlying facts.
        if '19 ·' in title:
            pieces.append(block('19а · Круговорот · Испарение',
               'Как называется переход воды с поверхности океана в пар?', '=испарение'))
            pieces.append(block('19б · Круговорот · Конденсация',
               'Как называется превращение водяного пара в жидкие капли?', '=конденсация'))
            pieces.append(block('19в · Круговорот · Осадки',
               'Как одним словом называют дождь, снег и град?', '=осадки'))
        else:
            pieces.append(block('20а · Что такое гидросфера',
               'Какая оболочка Земли называется гидросферой?',
               '=Водная\n~Воздушная\n~Твёрдая'))
            pieces.append(block('20б · Главные запасы воды',
               'Где сосредоточена основная часть воды Земли?',
               '=В океанах\n~В реках\n~В озёрах'))
            pieces.append(block('20в · Покрытие поверхности',
               'Около скольких процентов поверхности Земли покрывают океаны и моря?', '#71:1'))
    elif kind=='ddwtos':
        if '21 ·' in title:
            pieces.append(block('21а · Перетаскивание преобразовано · Испарение',
                'Как называется превращение жидкой воды в водяной пар?', '=испарение'))
            pieces.append(block('21б · Перетаскивание преобразовано · Конденсация',
                'Как называется превращение водяного пара в жидкие капли?', '=конденсация'))
        else:
            pieces.append(block('22а · Воды суши · Река',
                'Как называется естественный текущий водоток?', '=река'))
            pieces.append(block('22б · Воды суши · Озеро',
                'Как называется природный водоём в углублении суши?', '=озеро'))
    elif kind=='gapselect':
        if '23 ·' in title:
            pieces.append(block('23а · Где находится вода',
                'В какой части гидросферы находится основная масса воды Земли?',
                '=В Мировом океане\n~В атмосфере\n~В реках'))
            pieces.append(block('23б · Пресные воды',
                'Где находится основная масса пресной воды Земли?',
                '=В ледниках и ледяных покровах\n~В реках\n~В атмосфере'))
        else:
            pieces.append(block('24а · Океаническая вода',
                'Какой является вода Мирового океана?',
                '=Солёной\n~Пресной\n~Дистиллированной'))
            pieces.append(block('24б · Речная вода',
                'Какой обычно является вода большинства рек?',
                '=Пресной\n~Солёной\n~Морской'))
    else: skipped.append(kind)
    n+=1
assert not skipped,skipped
gift=''.join(pieces)
out=Path('Moodle_Gidrosfera_GIFT_voprosy.txt')
out.write_text(gift,encoding='utf-8')
assert gift.count('::')%2==0
assert gift.count('::')//2==34, gift.count('::')//2
print('created',out,'original assessed questions',n,'GIFT questions',gift.count('::')//2,'size',out.stat().st_size)
