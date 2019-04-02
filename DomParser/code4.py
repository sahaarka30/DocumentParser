import json
import xmltodict
import zipfile
import json

data = json.load(open('./demo.json'))

styles = set()
for paragraph in data:
    style = paragraph.get('style_name')
    styles.add(style)

archive = zipfile.ZipFile('easychair.docx', 'r')
data = archive.read('word/styles.xml')

data = json.dumps(xmltodict.parse(data))
data = json.loads(data)

data = data['w:styles']['w:style']

styles_lst = list()
for style in styles:
    style_id = style.replace(" ", "")
    link_id = ''
    for item in data:
        if style_id == item.get('@w:styleId'):
            if item.get('w:link'):
                link_id = item.get('w:link').get('@w:val')
                break

    di = dict()
    di['styles'] = dict()
    di['style_name'] = style
    print (style)
    for item in data:
        if link_id == item.get('@w:styleId'):
            font = item.get('w:rPr')
            if font.get('w:spacing'):
                di['styles']['spacing'] = font.get('w:spacing').get('@w:val')
            di['styles']['size'] = font.get('w:sz').get('@w:val')
            di['styles']['fonts'] = {
                'ascii': font.get('w:rFonts').get('@w:ascii'),
                'eastAsiaTheme': font.get('w:rFonts').get('@w:eastAsiaTheme'),
                'hAnsi': font.get('w:rFonts').get('@w:hAnsi'),
                'cstheme': font.get('w:rFonts').get('@w:cstheme'),
            }
            styles_lst.append(di)

            break

print(styles_lst)
# print"creating file ..."
with open('document_styles.json', 'w') as outfile:
    json.dump(styles_lst, outfile, indent=4)
# print "file created ..."
