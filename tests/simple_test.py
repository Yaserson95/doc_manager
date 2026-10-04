from docx_parser import DocxBlockParser
# ----------------------------------------------------------------------
# Пример использования
# ----------------------------------------------------------------------
if __name__ == '__main__':
    # Пример входного текста
    sample = '''
    {block type:"h:1" format:"text-style(bold,underline);txt-color(#0000ff);padding(0pt,30pt,0pt,10pt)"}Синий полужирный подчеркнутый заголовок первого уровня и 16-размера{/block}
    {block}Это просто абзац который содержит текст с {text format:"txt-color(#ffff00)"}желтым выделением{/text}{/block}
    '''

    parser = DocxBlockParser()
    doc = parser.parse(sample)
    doc.save('output.docx')
    print('Документ сохранён как output.docx')