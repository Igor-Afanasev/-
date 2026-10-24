import xml.etree.ElementTree as ET
from pathlib import Path
from collections import Counter
from zipfile import ZipFile, ZIP_DEFLATED
original='Moodle_Gidrosfera_bank_27_questions.xml'
root=ET.parse(original).getroot()
basic={'multichoice','truefalse','shortanswer','numerical','matching','essay','description'}
extended={'cloze','ddwtos','gapselect'}
basic_root=ET.Element('quiz')
extra_root=ET.Element('quiz')
for q in root.findall('question'):
    kind=q.get('type')
    if kind=='category': continue
    if kind in basic:
        if kind=='matching':
            for sub in q.findall('subquestion'):
                txt=sub.find('text')
                if txt is not None and txt.find('text') is not None:
                    inner=txt.find('text')
                    txt.text=inner.text
                    txt.remove(inner)
                if txt is None:
                    raise ValueError('Missing matching prompt')
        if kind=='essay' and q.find('answer') is None:
            a=ET.SubElement(q,'answer',{'fraction':'0','format':'moodle_auto_format'})
            ET.SubElement(a,'text').text=''
        basic_root.append(q)
    elif kind in extended:
        extra_root.append(q)
for q in basic_root.findall('question'):
    assert q.find('name/text') is not None
    assert q.find('questiontext/text') is not None
    if q.get('type')=='matching':
        for sq in q.findall('subquestion'):
            assert sq.find('text') is not None and sq.find('text/text') is None
            assert sq.find('answer/text') is not None
    if q.get('type')=='multichoice':
        fractions=[float(a.get('fraction','0')) for a in q.findall('answer')]
        assert abs(sum(v for v in fractions if v>0)-100)<0.01
for name,node in [
 ('Moodle_Gidrosfera_ISPRAVLENO_21_vopros.xml',basic_root),
 ('Moodle_Gidrosfera_dopolnitelnye_6_voprosov.xml',extra_root)
]:
    ET.indent(node,space='  ')
    ET.ElementTree(node).write(name,encoding='utf-8',xml_declaration=True)
    again=ET.parse(name).getroot()
    print(name,len(again.findall('question')),Counter(x.get('type') for x in again.findall('question')))
with ZipFile('Moodle_Gidrosfera_ISPRAVLENO.zip','w',ZIP_DEFLATED) as z:
    z.write('Moodle_Gidrosfera_ISPRAVLENO_21_vopros.xml')
    z.write('Moodle_Gidrosfera_dopolnitelnye_6_voprosov.xml')
