from docx.text.paragraph import Paragraph
from docx.text.run import Run
from docx.shared import Pt
from basic_checker import BasicChecker
from docx.oxml.ns import qn
from lxml.etree import Element

from utils import printing

import re
from copy import copy


def paragragph_contains_image(paragraph: Paragraph) -> bool:
    return bool(
        paragraph._p.xpath(
            "./w:r/w:drawing/*[self::wp:inline | self::wp:anchor]/a:graphic/a:graphicData/pic:pic"
        )
    )


def run_contains_image(run: Run) -> bool:
    return bool(
        run._r.xpath(
            "./w:drawing/*[self::wp:inline | self::wp:anchor]/a:graphic/a:graphicData/pic:pic"
        )
    )


class DocumentChecker1(BasicChecker):
    @staticmethod
    def score(paragraph, all_text: tuple[str], line_number: int) -> int:
        if paragragph_contains_image(paragraph):
            return 0
        score = 0
        if paragraph.text.strip().startswith("验证：") and (
            paragraph.paragraph_format.line_spacing != 1.5
            or paragraph.style.name != "Normal"
        ):
            score += 1
        if "List Paragraph" == paragraph.style.name:
            if paragraph.runs and paragraph.runs[0].bold:
                score += 1
                if paragraph.paragraph_format.line_spacing != 1.5:
                    score += 1
        return score

    def process(self, para, all_text: tuple[str], line_number: int) -> str | None:
        if self.score(para, all_text, line_number) > 0 and not self.perfect_match(
            para, all_text, line_number
        ):
            printing("Style(Normal):", para.style.name, file=self.logfile)
            printing(
                "Line spacing(1.5):",
                para.paragraph_format.line_spacing,
                file=self.logfile,
            )
            printing("Text:", para.text, file=self.logfile)
            answer = self.ask_for_process(
                "Fix line spacing to 1.5 and style to Normal? (Y/n)", file=self.logfile
            )
            if answer:
                para.paragraph_format.line_spacing = 1.5
                for run in para.runs:
                    run.font.name = "SimSun"
                    run.font.size = Pt(10)
                return para

    @staticmethod
    def guess(paragraph, all_text: tuple[str], line_number: int) -> str:
        # Only can fix can't guess
        pass

    @staticmethod
    def perfect_match(paragraph, all_text: tuple[str], line_number: int) -> bool:
        return (
            paragraph.style.name == "List Paragraph"
            and paragraph.runs[0].bold
            and paragraph.paragraph_format.line_spacing == 1.5
        ) or (
            paragraph.text.strip().startswith("验证：")
            and paragraph.style.name == "Normal"
            and paragraph.paragraph_format.line_spacing == 1.5
        )


class DocumentChecker2(BasicChecker):
    @staticmethod
    def score(paragraph, all_text: tuple[str], line_number: int) -> int:
        if paragragph_contains_image(paragraph):
            return 0
        score = 0
        if paragraph.text.strip().startswith("验证："):
            return score
        if "List Paragraph" == paragraph.style.name and not (
            paragraph.runs and paragraph.runs[0].bold
        ):
            score += 1
            if paragraph.paragraph_format.line_spacing != 1:
                score += 1
        if "Normal" == paragraph.style.name:
            score += 1
        return score

    def process(self, para, all_text: tuple[str], line_number: int) -> str | None:
        if self.score(para, all_text, line_number) > 0 and not self.perfect_match(
            para, all_text, line_number
        ):
            printing("Style:", para.style.name, file=self.logfile)
            printing("Text:", para.text, file=self.logfile)
            printing(
                "Line spacing:", para.paragraph_format.line_spacing, file=self.logfile
            )
            answer = self.ask_for_process(
                input_text="Fix line spacing to 1? (Y/n)", file=self.logfile
            )
            if answer:
                para.paragraph_format.line_spacing = 1
                return para

    @staticmethod
    def guess(paragraph, all_text: tuple[str], line_number: int) -> str:
        # Only can fix can't guess
        pass

    @staticmethod
    def perfect_match(paragraph, all_text: tuple[str], line_number: int) -> bool:
        return (
            (
                paragraph.style.name == "List Paragraph"
                and not (paragraph.runs and paragraph.runs[0].bold)
            )
            or (paragraph.style.name == "Normal")
            and paragraph.paragraph_format.line_spacing == 1
        )


class DocumentChecker3(BasicChecker):
    @staticmethod
    def score(paragraph, all_text: tuple[str], line_number: int) -> int:
        score = 0
        if paragragph_contains_image(paragraph):
            last_paragraph = all_text[line_number - 1]
            if last_paragraph.style.name == "List Paragraph":
                score += 1
                if paragraph.style.name != "List Paragraph":
                    score += 1
                    if not paragraph.paragraph_format.right_indent:
                        score += 1
                        if paragraph.paragraph_format.line_spacing != 1:
                            score += 1
        return score

    def process(self, para, all_text: tuple[str], line_number: int) -> str | None:
        if self.score(para, all_text, line_number) > 0 and not self.perfect_match(
            para, all_text, line_number
        ):
            printing("Style (List Paragraph):", para.style.name, file=self.logfile)
            printing(
                "Right indent (0 or None):",
                para.paragraph_format.right_indent,
                file=self.logfile,
            )
            printing(
                "Line spacing (1):",
                para.paragraph_format.line_spacing,
                file=self.logfile,
            )
            answer = self.ask_for_process("Fix image style? (Y/n)", file=self.logfile)
            if answer:
                para.style = "List Paragraph"
                para.paragraph_format.right_indent = None
                para.paragraph_format.line_spacing = 1
                return para

    @staticmethod
    def guess(paragraph, all_text: tuple[str], line_number: int) -> str:
        # Only can fix can't guess
        pass

    @staticmethod
    def perfect_match(paragraph, all_text: tuple[str], line_number: int) -> bool:
        return (
            paragraph.style.name == "List Paragraph"
            and not paragraph.paragraph_format.right_indent
            and paragraph.paragraph_format.line_spacing == 1
        )


class DocumentChecker4(BasicChecker):
    @staticmethod
    def score(paragraph, all_text: tuple[str], line_number: int) -> int:
        if paragragph_contains_image(paragraph):
            return 0
        if not paragraph.text.strip():
            return 1
        return 0

    def process(self, para, all_text: tuple[str], line_number: int) -> str | None:
        if self.score(para, all_text, line_number) > 0:
            if not para.text.strip():
                printing("Paragraph text is empty: ", para.text, file=self.logfile)
                answer = self.ask_for_process("Remove it? (Y/n)", file=self.logfile)
                if answer:
                    p = para._element
                    p.getparent().remove(p)
                    p._p = p._element = None
                    return para

    @staticmethod
    def guess(paragraph, all_text: tuple[str], line_number: int) -> str:
        # Only can fix can't guess
        pass

    @staticmethod
    def perfect_match(paragraph, all_text: tuple[str], line_number: int) -> bool:
        pass


class DocumentChecker5(BasicChecker):
    @staticmethod
    def score(paragraph, all_text: tuple[str], line_number: int) -> int:
        if paragragph_contains_image(paragraph):
            return 0
        try:
            next_line = all_text[line_number + 1]
        except IndexError:
            return 0
        score = 0
        if next_line.text.startswith("验证："):
            if paragraph.hyperlinks:
                last_hyper = paragraph.hyperlinks[-1]
                hyper_text = last_hyper.text
                para_text = paragraph.text.strip()
                if para_text.endswith(hyper_text):
                    score += 1
                    if last_hyper.address.endswith(hyper_text):
                        score += 1
        return score

    def process(self, para, all_text: tuple[str], line_number: int) -> str | None:
        if self.score(para, all_text, line_number) > 0:
            printing("Paragraph text is:", para.text, file=self.logfile)
            changed = False
            for hyper in para.hyperlinks:
                printing("Hyperlink:", hyper.text, hyper.address, file=self.logfile)
                answer = self.ask_for_process("Remove it? (Y/n)", file=self.logfile)
                if answer:
                    p = para._p
                    p.remove(hyper._hyperlink)
                    changed = True
            for run in reversed(list(para.runs)[1:]):
                if not run.text.strip():
                    para._p.remove(run._r)
            run = para.runs[-1]
            run.text = run.text.strip()
            if changed:
                return para

    @staticmethod
    def guess(paragraph, all_text: tuple[str], line_number: int) -> str:
        # Only can fix can't guess
        pass

    @staticmethod
    def perfect_match(paragraph, all_text: tuple[str], line_number: int) -> bool:
        pass


class DocumentChecker6(BasicChecker):
    @staticmethod
    def score(paragraph, all_text: tuple[str], line_number: int) -> int:
        if paragragph_contains_image(paragraph):
            return 0
        try:
            next_line = all_text[line_number + 1]
        except IndexError:
            return 0
        score = 0
        if next_line.text.startswith("验证："):
            score += 1
            if re.match("[a-zA-Z0-9\\s]*$", paragraph.text):
                score += 1
        return score

    def process(self, para, all_text: tuple[str], line_number: int) -> str | None:
        if self.score(para, all_text, line_number) > 0:
            printing("Paragraph text is:", para.text, file=self.logfile)
            changed = False
            for run in reversed(para.runs):
                if re.match("[a-zA-Z0-9\\s]*$", run.text):
                    printing("Run text is:", run.text, file=self.logfile)
                    if self.ask_for_process("Remove the end? (Y/n)", file=self.logfile):
                        # Remove trailing letters, numbers and whitespace from the run text
                        run.text = re.sub("[a-zA-Z0-9\\s]*$", "", run.text)
                        changed = True
                else:
                    break
            if changed:
                return para

    @staticmethod
    def guess(paragraph, all_text: tuple[str], line_number: int) -> str:
        # Only can fix can't guess
        pass

    @staticmethod
    def perfect_match(paragraph, all_text: tuple[str], line_number: int) -> bool:
        pass


class DocumentChecker7(BasicChecker):
    @staticmethod
    def score(paragraph, all_text: tuple[str], line_number: int) -> int:
        if (
            paragragph_contains_image(paragraph)
            and len(paragraph.runs) > 1
            and run_contains_image(paragraph.runs[-1])
        ):
            return 1
        return 0

    def process(self, para, all_text: tuple[str], line_number: int) -> str | None:
        if self.score(para, all_text, line_number) > 0:
            printing(
                "Paragraph contains image in the end, text is:",
                para.text,
                file=self.logfile,
            )
            if not self.ask_for_process(
                "Move the image to new paragraph? (Y/n)", file=self.logfile
            ):
                return
            try:
                printing(
                    "Next paragraph text is:",
                    all_text[line_number + 1].text,
                    file=self.logfile,
                )
                next_para = all_text[line_number + 1]
                new_para = next_para.insert_paragraph_before(style=para.style)
            except IndexError:
                printing("No next paragraph", file=self.logfile)
                new_para = para._parent.add_paragraph(style=para.style)
            run = para.runs[-1]
            new_para._p.append(run._r)
            runs = copy(para.runs)
            for run in runs:
                if not run.text.strip():
                    para._p.remove(run._r)
            return para

    @staticmethod
    def guess(paragraph, all_text: tuple[str], line_number: int) -> str:
        # Only can fix can't guess
        pass

    @staticmethod
    def perfect_match(paragraph, all_text: tuple[str], line_number: int) -> bool:
        pass


class DocumentChecker8(BasicChecker):
    @staticmethod
    def score(paragraph, all_text: tuple[str], line_number: int) -> int:
        score = 0
        paragraph_format = paragraph.paragraph_format
        if paragraph_format.space_after != Pt(0):
            score += 1
        if paragraph_format.space_before != Pt(0):
            score += 1
        return score

    def process(self, para, all_text: tuple[str], line_number: int) -> str | None:
        if self.score(para, all_text, line_number) > 0:
            para_format = para.paragraph_format
            printing(
                "Paragraph space before or after is not 0",
                "before:",
                para_format.space_before,
                "after:",
                para_format.space_after,
                file=self.logfile,
            )
            if not self.ask_for_process(
                "Remove space before and after? (Y/n)", file=self.logfile
            ):
                return
            para_format.space_before = Pt(0)
            para_format.space_after = Pt(0)
            return para

    @staticmethod
    def guess(paragraph, all_text: tuple[str], line_number: int) -> str:
        # Only can fix can't guess
        pass

    @staticmethod
    def perfect_match(paragraph, all_text: tuple[str], line_number: int) -> bool:
        pass


class DocumentChecker9(BasicChecker):
    @staticmethod
    def score(paragraph, all_text: tuple[str], line_number: int) -> int:
        if paragragph_contains_image(paragraph):
            return 0
        try:
            next_line = all_text[line_number + 1]
        except IndexError:
            return 0
        score = 0
        if next_line.text.startswith("验证："):
            if "test" in paragraph.text.lower():
                score += 1
        return score

    def process(self, para, all_text: tuple[str], line_number: int) -> str | None:
        if self.score(para, all_text, line_number) > 0:
            printing("Paragraph text is:", para.text, file=self.logfile)
            changed_text = re.split("test|Test|TEST", para.text)[0]
            changed_text = changed_text.strip()
            printing("Changed text is:", changed_text, file=self.logfile)
            answer = self.ask_for_process(
                "Remove 'test' from text? (Y/n)", file=self.logfile
            )
            if answer:
                for run in list(reversed(para.runs))[:-1]:
                    para._p.remove(run._r)
                run = para.runs[-1]
                run.text = changed_text
                return para

    @staticmethod
    def guess(paragraph, all_text: tuple[str], line_number: int) -> str:
        # Only can fix can't guess
        pass

    @staticmethod
    def perfect_match(paragraph, all_text: tuple[str], line_number: int) -> bool:
        pass


class DocumentChecker10(BasicChecker):
    @staticmethod
    def score(paragraph, all_text: tuple[str], line_number: int) -> int:
        if paragragph_contains_image(paragraph):
            if paragraph.paragraph_format.first_line_indent:
                return 1
        return 0

    def process(self, para, all_text: tuple[str], line_number: int) -> str | None:
        if self.score(para, all_text, line_number) > 0:
            printing("Contains image and first line indent is not 0", file=self.logfile)
            printing("First line indent:", para.paragraph_format.first_line_indent)
            printing("Line indent:", para.paragraph_format.left_indent)
            answer = self.ask_for_process(
                "Switch first line indent to line indent? (Y/n)", file=self.logfile
            )
            if answer:
                para.paragraph_format.left_indent = (
                    para.paragraph_format.first_line_indent
                )
                para.paragraph_format.first_line_indent = None
                return para

    @staticmethod
    def guess(paragraph, all_text: tuple[str], line_number: int) -> str:
        # Only can fix can't guess
        pass

    @staticmethod
    def perfect_match(paragraph, all_text: tuple[str], line_number: int) -> bool:
        pass


class DocumentChecker11(BasicChecker):
    def __init__(self) -> None:
        super().__init__()
        self.first_numId = None

    @staticmethod
    def _check_para_bold(para: Paragraph) -> bool:
        for run in para.runs:
            if not run.bold:
                return False
        return True

    @staticmethod
    def score(paragraph, all_text: tuple[str], line_number: int) -> int:
        if paragragph_contains_image(paragraph):
            return 0
        try:
            next_line = all_text[line_number + 1]
            next_2_line = all_text[line_number + 2]
        except IndexError:
            return 0
        if next_line.text.startswith("验证：") or (
            next_2_line.text.startswith("验证：")
            and DocumentChecker11._check_para_bold(paragraph)
        ):
            numPr_elements = paragraph._p.xpath(".//w:numPr")
            if numPr_elements:
                return 1
        return 0

    def __get_first_numId(self, all_para: list[Paragraph]):
        if self.first_numId:
            return self.first_numId
        for para in all_para:
            numPr_elements = para._p.xpath('.//w:numPr[w:ilvl/@w:val="0"]')
            if numPr_elements:
                numId_val = numPr_elements[0].xpath(".//w:numId/@w:val")
                if numId_val:
                    numId_val = numId_val[0]
                    self.first_numId = numId_val
                    return numId_val

    @staticmethod
    def __get_num_level(paragraph, all_text: tuple[str], line_number: int) -> int:
        try:
            next_line = all_text[line_number + 1]
            next_2_line = all_text[line_number + 2]
        except IndexError:
            return None
        if next_line.text.startswith("验证："):
            return 1
        if next_2_line.text.startswith("验证：") and DocumentChecker11._check_para_bold(
            paragraph
        ):
            return 0

    def process(self, para, all_text: tuple[str], line_number: int) -> str | None:
        if self.score(para, all_text, line_number) > 0:
            first_numId = self.__get_first_numId(all_text)
            printing("First numId:", first_numId, file=self.logfile)
            numPr_elements = para._p.xpath(".//w:numPr")
            ilvl = numPr_elements[0].xpath(".//w:ilvl")
            printing("ilvl:", ilvl, file=self.logfile)
            numId_elements = numPr_elements[0].xpath(".//w:numId")
            printing("numId:", numId_elements, file=self.logfile)
            if self.ask_for_process("Fix numPr? (Y/n)", file=self.logfile):
                if ilvl:
                    level = self.__get_num_level(para, all_text, line_number)
                    ilvl[0].set(qn("w:val"), str(level))

                if numId_elements:
                    numId_elements[0].set(qn("w:val"), first_numId)
                return para

    @staticmethod
    def guess(paragraph, all_text: tuple[str], line_number: int) -> str:
        # Only can fix can't guess
        pass

    @staticmethod
    def perfect_match(paragraph, all_text: tuple[str], line_number: int) -> bool:
        pass


class DocumentChecker12(BasicChecker):
    @staticmethod
    def score(paragraph, all_text: tuple[str], line_number: int) -> int:
        if paragragph_contains_image(paragraph):
            return 0
        try:
            next_line = all_text[line_number + 1]
            next_2_line = all_text[line_number + 2]
        except IndexError:
            return 0
        if next_line.text.startswith("验证："):
            return 1
        elif next_2_line.text.startswith(
            "验证："
        ) and DocumentChecker11._check_para_bold(paragraph):
            return 1
        return 0

    def process(self, para, all_text: tuple[str], line_number: int) -> str | None:
        if self.score(para, all_text, line_number) > 0:
            printing("Paragraph sytle is:", para.style.name, file=self.logfile)
            printing(
                "Paragraph text style (bold): ", para.runs[0].bold, file=self.logfile
            )
            printing("Paragraph text is:", para.text, file=self.logfile)
            re_text = r"^\d+(\.\d+|．\d+)*(\s*|　*)"
            text = re.sub(re_text, "", para.text).strip()
            printing("Text without number is:", text, file=self.logfile)
            if self.ask_for_process("Fix style? (Y/n)", file=self.logfile):
                para.style = "List Paragraph"
                para.paragraph_format.left_indent = Pt(21.6)
                ppr = para._p.xpath(".//w:pPr")[0]
                if not ppr.xpath(".//w:numPr"):
                    numPr_element = Element(qn("w:numPr"))
                    ilvl = Element(qn("w:ilvl"))
                    ilvl.set(qn("w:val"), "0")
                    numPr_element.append(ilvl)
                    numId = Element(qn("w:numId"))
                    numId.set(qn("w:val"), "0")
                    numPr_element.append(numId)
                    ppr.append(numPr_element)
                title_index = para.text.find(text)
                remove_runs = []
                remove_index = 0
                for run in para.runs:
                    run_length = len(run.text)
                    title_index -= run_length
                    if title_index >= 0:
                        remove_runs.append(run)
                    else:
                        remove_index = title_index + run_length
                        break

                # Remove the runs
                if not remove_runs:
                    printing("No runs to remove", file=self.logfile)
                for run in remove_runs:
                    printing("Remove run:", run.text, file=self.logfile)
                    para._p.remove(run._r)

                if remove_index:
                    printing("Remove index:", remove_index, file=self.logfile)
                    printing("Run text:", para.runs[0].text, file=self.logfile)
                    printing(
                        "Remaining text:",
                        para.runs[0].text[remove_index:],
                        file=self.logfile,
                    )
                    para.runs[0].text = para.runs[0].text[remove_index:]

                for run in para.runs:
                    run.bold = True
                return para

    @staticmethod
    def guess(paragraph, all_text: tuple[str], line_number: int) -> str:
        # Only can fix can't guess
        pass

    @staticmethod
    def perfect_match(paragraph, all_text: tuple[str], line_number: int) -> bool:
        pass
