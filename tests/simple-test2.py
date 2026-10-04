from docx_serializer import DocxBlockSerializer

serializer = DocxBlockSerializer()
text = serializer.serialize('new.docx')
print(text)