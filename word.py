import json
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


class BookmarkExtractor:
    def __init__(self, docx_file):
        self.docx_file = docx_file #инициализиция файла
        self.templates = []  #массив всех шаблонов
        self.current_template = [] #текущий шаблон
        self.current_block = {}  #массив закладок
        self.template_count = 0  #кол-во шаблонов
        self.page_break_count = 0  #счетчик для разрыва страниц
        self.new_page_flag = False  #флаг для новой страницы

    def extract_and_save_bookmarks(self, json_file):
        doc = Document(self.docx_file) #инициализируем документ

        #проход по всем элементам(включая таблицы и тд)
        for element in doc.element.body.iter():
            if element.tag == qn('w:bookmarkStart'): #тег закладок для поиска в xml
                bookmark_name = element.get(qn('w:name'))
                if bookmark_name:
                    text = self.get_text_from_bookmarks(doc, element)
                    self.current_block[bookmark_name] = text

            #проверка на новую страницу
            if element.tag == qn('w:br') and element.get(qn('w:type')) == 'page':
                self.page_break_count += 1
                self.new_page_flag = True

                #логика для 1 разрыва страницы
                if self.page_break_count == 1 and self.current_block:
                    self.current_template.append(self.current_block)
                    self.current_block = {}

                #логика для 2 разрывов страницы
                if self.page_break_count == 2:
                    if self.current_block:
                        self.current_template.append(self.current_block)
                    self.template_count += 1
                    self.templates.append({
                        f"block {self.template_count}": self.current_template,
                        "newpage": str(self.new_page_flag).lower()
                    })
                    self.current_template = []
                    self.current_block = {}
                    self.page_break_count = 0
                    self.new_page_flag = False

        #для незакрытых блоков
        if self.current_block:
            self.current_template.append(self.current_block)
        if self.current_template:
            self.template_count += 1
            self.templates.append({
                f"block {self.template_count}": self.current_template,
                "newpage": "false"
            })

        output_data = {
            "краткая информация": self.templates  # Список всех шаблонов с массивами
        }
        self.save_to_json(output_data, json_file)

    #получаем текст с закладок
    def get_text_from_bookmarks(self, doc, bookmark_start):
        bookmark_text = []
        inside_bookmark = False
        start_id = bookmark_start.get(qn('w:id'))

        for element in doc.element.body.iter():
            #начало закладки
            if element.tag == qn('w:bookmarkStart') and element.get(qn('w:id')) == start_id:
                inside_bookmark = True

            #собираем текст между началом и концом закладки
            if inside_bookmark and element.tag == qn('w:t'):
                bookmark_text.append(element.text)

            #конец закладки
            if element.tag == qn('w:bookmarkEnd') and element.get(qn('w:id')) == start_id:
                #брейкаем после конца закладки
                break

        return ' '.join(bookmark_text).strip()

    #сохраняем в json
    def save_to_json(self, data, json_file):
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=4)

#main функция для тестов
def main():
    docx_file = 'C:/Users/approxii/Desktop/word pravki/ext.docx' #изменить путь используя os
    json_file = 'C:/Users/approxii/Desktop/word pravki/bookmarks.json' #изменить путь используя os

    extractor = BookmarkExtractor(docx_file)
    extractor.extract_and_save_bookmarks(json_file)

main()