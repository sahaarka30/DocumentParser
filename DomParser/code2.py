from docx import *
import officedissector
import json

from xmltodict import parse

doc = officedissector.doc.Document('easychair.docx')
mp = doc.main_part()


doc_data = doc.to_json()

# print "creating file ..."
with open('data.json', 'w') as outfile:
    json.dump(json.loads(doc_data), outfile, indent=4)
# print "file created ..."

parts = list()
for document in json.loads(doc_data)['document']:
    if "parts" in document:
        parts = document['parts']
di = dict()
for item in parts:
    if "image" in item['content-type']:        
        uri = item['uri']
        relationship = item['relationships_in'][0].split()[1].replace('[', '').replace(']', '')
        di[relationship] = uri

print ("relationships :", di)
for key, value in di.iteritems():
    print ("key :", key)
