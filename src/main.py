from zipfile import ZipFile
import xml.etree.ElementTree as ET


class PyDocxConverter(object):
    def __init__(self, file_name, *args, **kwargs):
        self.docx_file_name = file_name

        docx_zip = ZipFile(file_name)
        doc_xml = docx_zip.read('word/document.xml').decode("utf-8")
        self.root = ET.fromstring(doc_xml)
        self.ns = self._get_namespaces()
        self.body = self.root.find('w:body', self.ns)

        docx_files = {}
        files = docx_zip.filelist
        for file in files:
            docx_files[file.filename] = file

        if 'word/numbering.xml' in docx_files:
            numbering_xml = docx_zip.read('word/numbering.xml').decode("utf-8")
            print(numbering_xml)

        print(doc_xml)

    def _get_namespaces(self):
        namespaces = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        }
        return namespaces

    def convert_to_html(self, *args, **kwargs):
        handler = _HtmlConvertHandler(self.root, self.ns)
        return handler.exec(*args, **kwargs)


class _ConvertHandler(object):
    def __init__(self, doc_root, ns):
        self.ns = ns
        self.doc_root = doc_root

    def exec(self, *args, **kwargs):
        raise NotImplementedError()


class _HtmlConvertHandler(_ConvertHandler):
    def __init__(self, doc_root, ns):
        super().__init__(doc_root, ns)
        self.el_stack = []
        self.is_text_bold = False
        self.is_text_italic = False
        self.is_text_underline = False

    def _strip_known_namespace(self, tag):
        if '}' in tag:
            uri, local_tag = tag[1:].split('}', 1)
            if uri in self.ns.values():
                return local_tag
        return tag

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

    def _process_text(self, xml_el, output):
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

    def _process_style_bold(self):
        self.is_text_bold = True

    def _process_style_italic(self):
        self.is_text_italic = True

    def _process_style_underline(self):
        self.is_text_underline = True

    def _process_style_font_color(self, xml_el, styles):
        for attr in xml_el.attrib:
            attr_name = self._strip_known_namespace(attr)
            if attr_name == 'val':
                styles.append('color: #' + xml_el.get(attr))

    def _process_run_props(self, xml_el, output):

        styles = []

        for child in xml_el:
            tag_name = self._strip_known_namespace(child.tag)
            if tag_name == 'color':
                self._process_style_font_color(child, styles)
            elif tag_name == 'b':
                self._process_style_bold()
            elif tag_name == 'i':
                self._process_style_italic()
            elif tag_name == 'u':
                self._process_style_underline()

        if len(styles) > 0:
            last_html_tag = output[-1]
            output[-1] = last_html_tag[:-1]
            output.append(' style="')
            output.extend(";".join(styles) + ';')
            output.append('"')
            output.append('>')

    def _process_line_break(self, xml_el, output):
        output.append('<br/>')

    def _process_run(self, xml_el, output):
        output.append('<span>')
        self._process_tags(xml_el, output)
        output.append('</span>')

    def _process_paragraph_props(self, xml_el, output):
        self._process_tags(xml_el, output)

    def _process_paragraph(self, xml_el, output):
        output.append('<p>')
        self._process_tags(xml_el, output)

        if output[-1] == '<p>':
            output.append('<br/>')

        output.append('</p>')

    def _process_tag(self, xml_el, output):
        self.el_stack.append(xml_el)
        tag_name = self._strip_known_namespace(xml_el.tag)

        if tag_name == 'p':
            self._process_paragraph(xml_el, output)

        elif tag_name == 'pPr':
            self._process_paragraph_props(xml_el, output)

        elif tag_name == 'r':
            self._process_run(xml_el, output)

        elif tag_name == 'rPr':
            self._process_run_props(xml_el, output)

        elif tag_name == 't':
            self._process_text(xml_el, output)

        elif tag_name == 'br':
            self._process_line_break(xml_el, output)

        self.el_stack.pop()

    def _process_tags(self, xml_el, output):
        for child in xml_el:
            self._process_tag(child, output)

    def exec(self, *args, **kwargs):
        body = self.doc_root.find('w:body', self.ns)

        output = []
        self.el_stack.append(body)
        self._process_tags(body, output)
        self.el_stack.pop()

        return "".join(output)


if __name__ == '__main__':
    converter = PyDocxConverter("C:/Users/paulr/Downloads/info page text (1).docx")
    output = converter.convert_to_html()
    print(output)

