import requests,zipfile,io,json
url='https://raw.githubusercontent.com/h5p/h5p-integration-test-suite/master/test-content/h5p-org-examples/course-presentation-08042017.h5p'
r=requests.get(url,timeout=60);print(r.status_code,len(r.content));r.raise_for_status()
z=zipfile.ZipFile(io.BytesIO(r.content));j=json.loads(z.read('h5p.json'));c=json.loads(z.read('content/content.json'))
p=c.get('presentation',{});slides=p.get('slides',[])
out={'manifest':j,'root_keys':list(c.keys()),'presentation_keys':list(p.keys()),'slide_count':len(slides),'slides':slides[:3],'file_count':len(z.namelist()),'libs':z.namelist()[:35]}
with open('h5p_sample_structure.json','w') as f:json.dump(out,f,ensure_ascii=False,indent=2)
