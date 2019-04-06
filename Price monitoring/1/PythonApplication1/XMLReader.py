from lxml import etree

class XMLError(BaseException):
    """description of exception class"""
    def __init__(self, message): self.message = message
    def __str__(self): return self.message

class ReadingXMLError(BaseException):
    """description of exception class"""

    def __init__(self, message): self.message = message

    def __str__(self): return self.message

class XMLReader:
    def __init__(self, path):
        tree = etree.parse(path)
        self.root = tree.getroot()

    def get_list_of_dicts(self):
        return self._get_list_output(self.root)

    def _get_children_list(self, parent):
        # В данном цикле мы перебираем элементы, которые вложены в родительский элемент.
        # Для включения дочерних элементов, текущего элемента цикла, используем рекурсию,
        # если дочерних элементов нет -- функция вернет пустой список.
        children = list()
        for element in parent:
            d = dict()
            d['tag'] = element.tag
            d['attrib'] = dict(element.attrib)
            d['text'] = element.text
            d['children'] = self._get_children_list(element)
            children.append(d)
        return children

    def _get_list_output(self, root):
        # Включаем в словарь с данными корневого элемента, данные всех дочерних.
        children = self._get_children_list(root)
        #list_dict = [{'tag': root.tag, 'attrib': dict(root.attrib), 'text': root.text, 'children': children}]
        return children