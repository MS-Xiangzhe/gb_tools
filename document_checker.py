from docx.text.paragraph import Paragraph
from basic_checker import BasicChecker

from utils import printing


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
            printing("Line spacing:", para.paragraph_format.line_spacing, file=self.logfile)
            answer = input("Fix line spacing to 1.5? (Y/n)", file=self.logfile)
            answer = answer.strip().lower()
            if answer == "y" or not answer:
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
            answer = input("Fix line spacing to 1? (Y/n)")
            answer = answer.strip().lower()
            if answer == "y" or not answer:
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
            answer = input("Fix image style? (Y/n)")
            answer = answer.strip().lower()
            if answer == "y" or not answer:
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
