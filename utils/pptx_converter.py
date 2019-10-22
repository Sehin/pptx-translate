from pptx import Presentation
from pptx.slide import Slide


class PPTXConverter:
    def __init__(self, path):
        """
        Путь к pptx файлу
        :param path:
        """
        self.path = path
        self.presentation = Presentation(path)

    def translate(self, dictionary):
        for slide in self.presentation.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:    # Если есть текст в форме
                    # Прохожусь по словарю, если есть совпадения в текст фрейме
                    for key in dictionary.keys():
                        if key in shape.text_frame.text:
                            text_frame = shape.text_frame
                            cur_text = text_frame.paragraphs[0].runs[0].text
                            if len(cur_text) < len(shape.text):
                                new_text = str(dictionary[key])
                            else:
                                new_text = cur_text.replace(str(key), str(dictionary[key]))
                            print(f"Change {text_frame.text} ||| {new_text}")
                            text_frame.paragraphs[0].runs[0].text = new_text

                            if len(shape.text_frame.paragraphs[0].runs) > 1:
                                is_first = True
                                for run in shape.text_frame.paragraphs[0].runs:
                                    if is_first:
                                        is_first = False
                                        continue
                                    run.text = ''
                            break




    def translate1(self, dictionary):
        """
        Перевести весь текст, по словарю
        """
        for slide in self.presentation.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    if shape.text.lower() in [key.lower() for key in dictionary.keys()]:
                        text = dictionary[shape.text]

                        text_frame = shape.text_frame
                        cur_text = text_frame.paragraphs[0].runs[0].text
                        new_text = cur_text.replace(str(cur_text), str(text))
                        text_frame.paragraphs[0].runs[0].text = new_text


                        if len(text_frame.paragraphs[0].runs) > 1:
                            is_first = True
                            for run in text_frame.paragraphs[0].runs:
                                if is_first:
                                    is_first = False
                                    continue
                                run.text = ''

                        # shape.text_frame.paragraphs[0].runs[0].text = text


    def save(self, path):
        self.presentation.save(path)

    def search_and_replace(self, search_str, repl_str, input, output):
        """"search and replace text in PowerPoint while preserving formatting"""
        # Useful Links ;)
        # https://stackoverflow.com/questions/37924808/python-pptx-power-point-find-and-replace-text-ctrl-h
        # https://stackoverflow.com/questions/45247042/how-to-keep-original-text-formatting-of-text-with-python-powerpoint
        prs = self.presentation
        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    if (shape.text.find(search_str)) != -1:
                        text_frame = shape.text_frame
                        cur_text = text_frame.paragraphs[0].runs[0].text
                        new_text = cur_text.replace(str(search_str), str(repl_str))
                        text_frame.paragraphs[0].runs[0].text = new_text
        # prs.save(output)