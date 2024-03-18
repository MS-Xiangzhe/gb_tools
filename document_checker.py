from docx.text.paragraph import Paragraph
from basic_checker import BasicChecker

from utils import printing

import re


def paragragph_contains_image(paragraph: Paragraph) -> bool:
    return bool(
        paragraph._p.xpath(
            "./w:r/w:drawing/*[self::wp:inline | self::wp:anchor]/a:graphic/a:graphicData/pic:pic"
        )
    )


class DocumentChecker1(BasicChecker):
    @staticmethod
    def score(paragraph, all_text: tuple[str], line_number: int) -> int:
        if paragragph_contains_image(paragraph):
            return 0
        score = 0
        if (
            paragraph.text.strip().startswith("验证：")
            and paragraph.style.name == "Normal"
            and paragraph.paragraph_format.line_spacing != 1.5
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
            printing("Style:", para.style.name, file=self.logfile)
            printing("Text:", para.text, file=self.logfile)
            printing(
                "Line spacing:", para.paragraph_format.line_spacing, file=self.logfile
            )
            answer = self.ask_for_process(
                self, "Fix line spacing to 1.5? (Y/n)", file=self.logfile
            )
            if answer:
                para.paragraph_format.line_spacing = 1.5
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
            printing("Style:", para.style.name)
            printing("Text:", para.text)
            printing("Line spacing:", para.paragraph_format.line_spacing)
            answer = self.ask_for_process(
                self, "Fix line spacing to 1? (Y/n)", file=self.logfile
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
            printing("Style (List Paragraph):", para.style.name)
            printing("Right indent (0 or None):", para.paragraph_format.right_indent)
            printing("Line spacing (1):", para.paragraph_format.line_spacing)
            answer = self.ask_for_process(
                self, "Fix image style? (Y/n)", file=self.logfile
            )
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
            printing("Paragraph text is empty: ", para.text)
            answer = self.ask_for_process(self, "Remove it? (Y/n)", file=self.logfile)
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
            score += 1
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
            printing("Paragraph text is:", para.text)
            changed = False
            for hyper in para.hyperlinks:
                printing("Hyperlink:", hyper.text, hyper.address)
                answer = self.ask_for_process(
                    self, "Remove it? (Y/n)", file=self.logfile
                )
                if answer:
                    p = para._p
                    p.remove(hyper._hyperlink)
                    changed = True
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
            printing("Paragraph text is:", para.text)
            changed = False
            for run in reversed(para.runs):
                if re.match("[a-zA-Z0-9\\s]*$", run.text):
                    printing("Run text is:", run.text)
                    if self.ask_for_process(
                        self, "Remove the end? (Y/n)", file=self.logfile
                    ):
                        # Remove trailing letters, numbers and whitespace from the run text
                        run.text = re.sub('[a-zA-Z0-9\\s]*$', '', run.text)
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
