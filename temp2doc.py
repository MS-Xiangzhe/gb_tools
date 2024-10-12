from argparse import ArgumentParser
from dataclasses import dataclass
import re
from docx import Document
from copy import deepcopy


@dataclass
class DocumentPart2:
    title: str
    overview_en: str
    overview_zh: str


@dataclass
class DocumentPart:
    title: str
    part_list: list[DocumentPart2]


def txt_lines_to_part2(lines: list[str]) -> DocumentPart2:
    title = lines[0]
    # Find zh title under startwith 中文翻译 and underlines have 2
    for i in range(1, len(lines)):
        if lines[i].startswith("中文翻译"):
            if len(lines) - i >= 3:
                title = lines[i + 1]
            overview_zh = "\n".join(lines[i:])
            overview_en = "\n".join(lines[:i])
            break
    return DocumentPart2(title=title, overview_en=overview_en, overview_zh=overview_zh)


def txt_parse_document_part(txt_path: str) -> DocumentPart:
    with open(txt_path, "r", encoding="utf-8") as f:
        lines = [line.strip() for line in f.readlines() if line.strip()]
        part1_list = []

        tmp_list = []
        for i in range(len(lines)):
            line = lines[i]
            create_new_part = False
            create_new_part2 = False
            # Check line is start with number)
            # And get the number
            m = re.match(r"^\d+\)", line)
            if m:
                number = m.group(0)[:-1]
                if int(number) == 1:
                    create_new_part = True
                else:
                    create_new_part2 = True
                # Remove d) from the line
                line = line[len(number) + 1 :].strip()
            if i >= len(lines) - 1:
                create_new_part2 = True
            if create_new_part:
                part1_list.append(DocumentPart(tmp_list[-1], part_list=[]))
                tmp_list = []
            if create_new_part2:
                part1_list[-1].part_list.append(txt_lines_to_part2(tmp_list))
                tmp_list = []
            tmp_list.append(line)
        return part1_list


PARAGRAPH_XML_MAP = {
    "title1": None,
    "title2": None,
    "overview_title": None,
    "overview_en": None,
    "overview_zh": None,
    "overview_end1": None,
    "overview_end2": None,
}


def doc_init_xml_map(doc: Document):
    for p in doc.paragraphs:
        if p.text == "title1":
            PARAGRAPH_XML_MAP["title1"] = p
        elif p.text == "title2":
            PARAGRAPH_XML_MAP["title2"] = p
        elif p.text == "概述":
            PARAGRAPH_XML_MAP["overview_title"] = p
        elif p.text == "eng_body":
            PARAGRAPH_XML_MAP["overview_en"] = p
        elif p.text == "zh_body":
            PARAGRAPH_XML_MAP["overview_zh"] = p
        elif p.text == "测试场景":
            PARAGRAPH_XML_MAP["overview_end1"] = p
        elif p.text == "此新功能没有GB18030相关的测试场景。":
            PARAGRAPH_XML_MAP["overview_end2"] = p


def _clear_content(element):
    """Remove all child elements"""
    for e in element.xpath("./*"):
        element.remove(e)


def doc_add_paragraph(doc: Document, para_type: str, text: str | None = None):
    paragraph = doc.add_paragraph()
    _clear_content(paragraph._element)
    items = PARAGRAPH_XML_MAP[para_type]._element.xpath("./*")
    items = deepcopy(items)
    for item in items:
        paragraph._element.append(item)
    run = paragraph.runs[0]
    if text is not None:
        run.text = text


def main():
    parser = ArgumentParser(description="Process some files.")
    parser.add_argument("--txt", required=True, help="Path to the input file")
    parser.add_argument("--temp", required=True, help="Path to the template file")
    parser.add_argument("--doc", required=True, help="Path to the output file")
    args = parser.parse_args()
    txt_path = args.txt
    part_list = txt_parse_document_part(txt_path)
    doc_path = args.temp
    doc = Document(doc_path)
    doc_init_xml_map(doc)
    for part in part_list:
        doc_add_paragraph(doc, "title1", part.title)
        for part2 in part.part_list:
            doc_add_paragraph(doc, "title2", part2.title)
            doc_add_paragraph(doc, "overview_title")
            doc_add_paragraph(doc, "overview_en", part2.overview_en)
            doc_add_paragraph(doc, "overview_zh", part2.overview_zh)
            doc_add_paragraph(doc, "overview_end1")
            doc_add_paragraph(doc, "overview_end2")
    doc.save(args.doc)


if __name__ == "__main__":
    main()
