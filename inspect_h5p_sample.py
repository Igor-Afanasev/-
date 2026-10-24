import requests,zipfile,io,json
url='https://h5p.org/sites/default/files/h5p/exports/612.h5p'
r=requests.get(url,timeout=45);print(r.status_code,len(r.content),r.headers.get('content-type'));r.raise_for_status()
z=zipfile.ZipFile(io.BytesIO(r.content))
j=json.loads(z.read('h5p.json'));c=json.loads(z.read('content/content.json'))
out={'manifest':j,'root_keys':list(c.keys()),'sample_slide':c.get('presentation',{}).get('slides',[{}])[0],'file_count':len(z.namelist()),'libs':z.namelist()[:30]}
with open('h5p_sample_structure.json','w') as f:json.dump(out,f,ensure_ascii=False,indent=2)
with open('h5p_example_612.h5p','wb') as f:f.write(r.content)
