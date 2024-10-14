from io import BytesIO
from io import FileIO
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import os
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P

from core.services.base import BaseDocumentService


class WordService:
    def __init__(self):
        self.docx_file = None

    def load(self, file) -> None:
        self.docx_file = Document(file)


    def update(self, params: dict) -> None:
        if not self.docx_file:
            raise ValueError("Word файл не загружен.")

        #проход по всем элементам(включая таблицы и тд)
        for index, block in enumerate(params['blocks']):
            doc, temp_filename = self.copy_to_temp(index)
            for key, value in block.items():
                print(f"Обработка блока {index}, ключ: {key}, значение: {value}")
                bookmark_found = False
                for element in doc.element.body.iter():
                    if element.tag == qn('w:bookmarkStart'):  # тег закладок для поиска в списке xml
                        bookmark_name = element.get(qn('w:name'))
                        if bookmark_name == key:
                            print(f"Закладка найдена: {bookmark_name}, текст для замены: {value}")
                            self.replace_text(doc, element, value)
                            bookmark_found = True
                if not bookmark_found:
                    print(f"Закладки в документе не найдены")

            doc.save(temp_filename)
            self.add_temp_to_original(self.docx_file, temp_filename, params)

            #удаление временных файлов
            if os.path.exists(temp_filename):
                os.remove(temp_filename)

    def copy_to_temp(self, index):
        #функция копирования данных во временные файлы
        temp_filename = f'temp{index}.docx'
        original_doc = self.docx_file
        original_doc.save(temp_filename)
        return original_doc, temp_filename

    def add_temp_to_original(self, original_doc, temp_doc_path, params: dict):
        #функция комбинирования временного файла с результатом
        temp_doc = Document(temp_doc_path)

        elements_to_copy = list(temp_doc.element.body)
        paragraph_index = 0
        table_index = 0

        for element in elements_to_copy:
            if element.tag.endswith('p'):
                if paragraph_index < len(temp_doc.paragraphs):
                    paragraph = temp_doc.paragraphs[paragraph_index]
                    self.copy_paragraph(paragraph, original_doc)
                    paragraph_index += 1
            elif element.tag.endswith('tbl'):
                if table_index < len(temp_doc.tables):
                    table = temp_doc.tables[table_index]

                    new_table = original_doc.add_table(rows=len(table.rows), cols=len(table.columns))
                    new_table.style = table.style

                    for row_index, row in enumerate(table.rows):
                        for col_index, cell in enumerate(row.cells):
                            new_table.cell(row_index, col_index).text = cell.text

                            tcPr = new_table.cell(row_index, col_index)._tc.get_or_add_tcPr()
                            tcBorders = OxmlElement("w:tcBorders")
                            for border in ["top", "left", "bottom", "right"]:
                                element = OxmlElement(f"w:{border}")
                                element.set(qn("w:val"), "single")
                                element.set(qn("w:sz"), "4")
                                element.set(qn("w:space"), "0")
                                element.set(qn("w:color"), "auto")
                                tcBorders.append(element)
                            tcPr.append(tcBorders)

                    table_index += 1

        if params['newpage'] == 'true':
            original_doc.add_page_break()


    def replace_text(self, doc, bookmark_element, new_text):
        #функция замены текста закладки на новый текст из json
        for sibling in bookmark_element.itersiblings():
            if sibling.tag == qn('w:r'):
                for child in sibling.iter():
                    if child.tag == qn('w:t'):
                        print(f"Изменили: {child.text} на {new_text}")
                        child.text = new_text
                        return


    def copy_paragraph(self, paragraph, document):
        #ункция для копирования абзаца в новый документ
        new_paragraph = document.add_paragraph()
        new_paragraph.style = paragraph.style  #стиль абзаца
        for run in paragraph.runs:
            new_run = new_paragraph.add_run(run.text)  #текст
            new_run.bold = run.bold  #жирность
            new_run.italic = run.italic  #курсив
            new_run.font.size = run.font.size  #размер шрифта
            if run.font.color and run.font.color.rgb:
                new_run.font.color.rgb = run.font.color.rgb  #цвет шрифта(если есть)

    def save_to_bytes(self) -> BytesIO:
        if not self.docx_file:
            raise ValueError("Word файл не загружен.")
        output = BytesIO()
        self.docx_file.save(output)
        output.seek(0)
        return output

    def save_to_file(self, file_path: str) -> None:
        if self.docx_file:
            self.docx_file.save(file_path)
        else:
            raise ValueError("Word файл не загружен.")