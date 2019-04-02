import zipfile
import json

archive = zipfile.ZipFile('easychair.docx', 'r')
data = archive.read('word/document.xml')


import xmltodict

data  = json.dumps(xmltodict.parse(data))
data = json.loads(data)
data = data['w:document']['w:body']['w:p']

content = list()
for index, paragraph in enumerate(data):
    para_content = list()
    para_index = index
    if paragraph.get("w:r"):
        for item in paragraph.get("w:r"):
            di = dict()
            if type(item) is dict:
                if  type(item.get('w:t')) is dict:
                    text = item.get('w:t').get('#text')
                else:
                    text = item.get('w:t')
                di['text'] = text
            para_content.append(di)
    if "a:blip" in str(paragraph):
        di = dict()
        embed = str(paragraph).split('@r:embed')[1].split('}')[0].split()[1].replace("u'", "").replace("'", "")
        di['embed'] = embed
        para_content.append(di)
    di = dict()
    di['para_content'] = para_content
    di['para_index'] = para_index
    content.append(di)
    # print(di)
print(content)
# print "creating file ..."
with open('document_content.json', 'w') as outfile:
    json.dump(content, outfile, indent=4)
# print "file created ..."
