from zipfile import ZipFile
import xml.etree.ElementTree as ET


def strip_known_namespace(tag, ns):
    if '}' in tag:
        uri, local_tag = tag[1:].split('}', 1)
        if uri in ns.values():
            return local_tag
    return tag


def get_attr_value(xml_el, key, ns: dict):
    return xml_el.get('{' + ns['w'] + '}' + key)


class DocxNumbering(object):
    def __init__(self, root, ns):
        self.root = root
        self.ns = ns
        self.content = {}

    def _get_abstract_num_id(self, num_id):
        element = self.root.find('.//w:num[@w:numId="{}"]'.format(num_id), self.ns)
        abstract_num_el = element.find('.//w:abstractNumId', self.ns)
        return get_attr_value(abstract_num_el, 'val', self.ns)

    def _get_lvl_el(self, abstract_num_id, ilvl):
        element = self.root.find('.//w:abstractNum[@w:abstractNumId="{}"]'.format(abstract_num_id), self.ns)
        lvl = element.find('.//w:lvl[@w:ilvl="{}"]'.format(ilvl), self.ns)
        return lvl

    def get_list_format(self, num_id, ilvl) -> dict:
        abstract_num_id = self._get_abstract_num_id(num_id)
        lvl_el = self._get_lvl_el(abstract_num_id, ilvl)

        num_fmt_el = lvl_el.find('.//w:numFmt', self.ns)
        return {
            'style': get_attr_value(num_fmt_el, 'val', self.ns)
        }


class DocxDocument(object):
    def __init__(self, root, ns):
        self.root = root
        self.ns = ns

        self._process()

    def get_body(self):
        return self.root.find('w:body', self.ns)

    def _process(self):
        pass


class PyDocxConverter(object):
    def __init__(self, file_name, *args, **kwargs):
        self.docx_file_name = file_name

        docx_zip = ZipFile(file_name)
        doc_xml = docx_zip.read('word/document.xml').decode("utf-8")
        self.doc_root = ET.fromstring(doc_xml)
        self.ns = self._get_namespaces()

        self.docx_document = DocxDocument(self.doc_root, self.ns)
        self.docx_numbering = None

        docx_files = {}
        files = docx_zip.filelist
        for file in files:
            docx_files[file.filename] = file

        if 'word/numbering.xml' in docx_files:
            numbering_xml = docx_zip.read('word/numbering.xml').decode("utf-8")
            self.numbering_root = ET.fromstring(numbering_xml)
            print(numbering_xml)
            self.docx_numbering = DocxNumbering(self.numbering_root, self.ns)

        print(doc_xml)

    def _get_namespaces(self):
        namespaces = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        }
        return namespaces

    def convert_to_html(self, *args, **kwargs):
        handler = _HtmlConvertHandler(self.docx_document, self.docx_numbering, self.ns)
        return handler.exec(*args, **kwargs)


class _ConvertHandler(object):
    def __init__(self, docx_document, docx_numbering, ns):
        self.docx_document: DocxDocument = docx_document
        self.docx_numbering: DocxNumbering = docx_numbering
        self.ns = ns

    def exec(self, *args, **kwargs):
        raise NotImplementedError()


class _ParagraphState(object):
    def __init__(self, p_el):
        self.p_el = p_el
        self.style = None
        self.css_styles = None
        self.ilvl = None
        self.num_id = None

    def set_style(self, style):
        self.style = style

    def set_css_styles(self, css_styles):
        self.css_styles = css_styles

    def set_number_ptr(self, num_id, ilvl):
        self.num_id = num_id
        self.ilvl = ilvl

    def get_num_id(self):
        return self.num_id

    def get_ilvl(self):
        return self.ilvl


class _HtmlElement(object):
    def __init__(self, tag_name=None):
        self.parent = None
        self.tag_name = tag_name
        self.text = None
        self.children: list[_HtmlElement] = []
        self.styles: list[str] = []
        self.classes: list[str] = []

    def set_tag_name(self, tag_name):
        self.tag_name = tag_name

    def set_parent(self, parent):
        self.parent = parent

    def get_parent(self):
        return self.parent

    def add_child(self, html_el):
        html_el.set_parent(self)
        self.children.append(html_el)

    def get_children(self):
        return self.children

    def has_children(self):
        return len(self.children) > 0

    def is_empty(self):
        return len(self.children) == 0 and self.text is None

    def get_tag_name(self):
        return self.tag_name

    def render(self):
        output = []

        if self.tag_name == 'br':
            output.append('<br/>')
        else:
            output.append('<' + self.tag_name + '>')

            if self.text is not None:
                output.append(self.text)

            for child in self.children:
                output.append(child.render())

            output.append('</' + self.tag_name + '>')

        return "".join(output)


class _HtmlConvertHandler(_ConvertHandler):
    def __init__(self, docx_document, docx_numbering, ns):
        super().__init__(docx_document, docx_numbering, ns)
        self.xml_el_stack = []
        self.html_el_tree: _HtmlElement = _HtmlElement('div')
        self.recently_added_html_el = self.html_el_tree
        self.output = []

        self.p_state: _ParagraphState = None

        self.is_text_bold = False
        self.is_text_italic = False
        self.is_text_underline = False

    def _process_italic_text(self, xml_el, output):
        output.append('<i>')
        output.append(xml_el.text)
        output.append('</i>')
        self.is_text_italic = False

    def _process_bold_text(self, xml_el, output):
        output.append('<b>')
        output.append(xml_el.text)
        output.append('</b>')
        self.is_text_bold = False

    def _process_underline_text(self, xml_el, output):
        output.append('<u>')
        output.append(xml_el.text)
        output.append('</u>')
        self.is_text_underline = False

    def _process_text(self, xml_el, _html_el):
        _html_el.text = xml_el.text
        """
        if self.is_text_bold:
            output.append('<b>')

        if self.is_text_italic:
            output.append('<i>')

        if self.is_text_underline:
            output.append('<u>')

        output.append(xml_el.text)

        if self.is_text_underline:
            output.append('</u>')
            self.is_text_underline = False

        if self.is_text_italic:
            output.append('</i>')
            self.is_text_italic = False

        if self.is_text_bold:
            output.append('</b>')
            self.is_text_bold = False
        """

    def _process_style_bold(self):
        self.is_text_bold = True

    def _process_style_italic(self):
        self.is_text_italic = True

    def _process_style_underline(self):
        self.is_text_underline = True

    def _process_style_font_color(self, xml_el, styles):
        for attr in xml_el.attrib:
            attr_name = strip_known_namespace(attr, self.ns)
            if attr_name == 'val':
                styles.append('color: #' + xml_el.get(attr))

    def _process_run_props(self, xml_el, _html_el):

        styles = []

        for child in xml_el:
            tag_name = strip_known_namespace(child.tag, self.ns)
            if tag_name == 'color':
                self._process_style_font_color(child, styles)
            elif tag_name == 'b':
                self._process_style_bold()
            elif tag_name == 'i':
                self._process_style_italic()
            elif tag_name == 'u':
                self._process_style_underline()

        if len(styles) > 0:
            self.p_state.set_css_styles(styles)

    def _process_line_break(self, xml_el, _html_el):
        #output.append('<br/>')
        pass

    def _process_run(self, xml_el, _html_el):
        span_el = _HtmlElement('span')
        _html_el.add_child(span_el)
        self._process_tags(xml_el, span_el)

    def _process_paragraph_props(self, xml_el, _html_el):
        style_el = xml_el.find('.//w:pStyle', self.ns)
        if style_el is not None:
            self.p_state.set_style(get_attr_value(style_el, 'val', self.ns))

        num_pr_el = xml_el.find('.//w:numPr', self.ns)
        if num_pr_el is not None:
            num_id_el = num_pr_el.find('.//w:numId', self.ns)
            num_id = get_attr_value(num_id_el, 'val', self.ns)
            ilvl_el = num_pr_el.find('.//w:ilvl', self.ns)
            ilvl = get_attr_value(ilvl_el, 'val', self.ns)
            self.p_state.set_number_ptr(num_id, ilvl)

        self._process_tags(xml_el, _html_el)

    def _process_paragraph(self, xml_el, _html_el):
        self.p_state = _ParagraphState(xml_el)

        html_el = _HtmlElement()
        _html_el.add_child(html_el)

        props_el = xml_el.find('.//w:pPr', self.ns)
        if props_el is not None:
            self._process_paragraph_props(props_el, html_el)

        if self.p_state.style is None:
            #output.append('<p>')
            html_el.set_tag_name('p')

            #if self.p_state.css_styles is not None:
            #    last_html_tag = output[-1]
            #    output[-1] = last_html_tag[:-1]
            #    output.append(' style="')
            #    output.extend(";".join(self.p_state.css_styles) + ';')
            #    output.append('"')
            #    output.append('>')

        else:
            num_id = self.p_state.get_num_id()
            ilvl = self.p_state.get_ilvl()
            list_format = self.docx_numbering.get_list_format(num_id, ilvl)
            style = list_format.get('style')

            if style == 'bullet':
                #output.append('<ul><li>')
                html_el.set_tag_name('ul')

                list_item_html_el = _HtmlElement('li')
                html_el.add_child(list_item_html_el)
            else:
                #output.append('<ol><li>')
                html_el.set_tag_name('ol')

                list_item_html_el = _HtmlElement('li')
                html_el.add_child(list_item_html_el)

            # push the element to the list item
            html_el = list_item_html_el

        run_els = xml_el.findall('.//w:r', self.ns)
        for run_el in run_els:
            self._process_run(run_el, html_el)

        if self.p_state.style is None:
            # if empty paragraph
            #if output[-1] == '<p>':
            #    output.append('<br/>')

            if html_el.get_tag_name() == 'p' and html_el.is_empty():
                html_el.add_child(_HtmlElement('br'))

            #output.append('</p>')

        else:
            pass
            #if style == 'bullet':
            #    output.append('</li></ul>')
            #else:
            #    output.append('</li></ol>')

    def _process_tag(self, xml_el, _html_el):
        self.xml_el_stack.append(xml_el)
        tag_name = strip_known_namespace(xml_el.tag, self.ns)

        if tag_name == 'p':
            self._process_paragraph(xml_el, _html_el)

        elif tag_name == 'pPr':
            self._process_paragraph_props(xml_el, _html_el)

        elif tag_name == 'r':
            self._process_run(xml_el, _html_el)

        elif tag_name == 'rPr':
            self._process_run_props(xml_el, _html_el)

        elif tag_name == 't':
            self._process_text(xml_el, _html_el)

        elif tag_name == 'br':
            self._process_line_break(xml_el, _html_el)

        self.xml_el_stack.pop()

    def _process_tags(self, xml_el, _html_el):
        for child in xml_el:
            self._process_tag(child, _html_el)

    def add_html_el(self, html_el):
        self.html_el_tree.add_child(html_el)
        self.recently_added_html_el = html_el

    def render_html_el(self):
        html_el = self.html_el_tree
        self.output.append(html_el.render())

    def get_html_el(self):
        return self.html_el_tree

    def get_recently_added_html_el(self):
        return self.recently_added_html_el

    def exec(self, *args, **kwargs):
        body = self.docx_document.get_body()

        self.xml_el_stack.append(body)
        self._process_tags(body, self.html_el_tree)
        self.xml_el_stack.pop()

        return self.html_el_tree.render()


if __name__ == '__main__':
    converter = PyDocxConverter("C:/Users/paulr/Downloads/info page text (1).docx")
    output = converter.convert_to_html()
    print(output)

