import re
from basic_checker import BasicChecker

from utils import printing


class BasicTextChecker(BasicChecker):
    def process(self, para, all_text: tuple[str], line_number: int) -> str | None:
        text = para.text
        if self.score(text, all_text, line_number) == 0:
            return
        guess = self.guess(text, all_text, line_number)
        if guess == text:
            return
        printing("Need to fix: ", end="", file=self.logfile)
        if not self.perfect_match(text, all_text, line_number):
            printing("Not perfect match", file=self.logfile)
        else:
            printing("Perfect match but not in guess", file=self.logfile)
        printing("Text:", text, file=self.logfile)
        printing("Guess:", guess, file=self.logfile)
        text = input("Enter to continue")
        if text.strip():
            if text.strip().lower() == "y":
                return guess
            else:
                return text
        else:
            return

    @staticmethod
    def score(text: str, all_text: tuple[str], line_number: int) -> int:
        pass

    @staticmethod
    def guess(text: str, all_text: tuple[str], line_number: int) -> str:
        pass

    @staticmethod
    def perfect_match(text: str, all_text: tuple[str], line_number: int) -> bool:
        pass


class TextChecker1(BasicTextChecker):

    # Get previous line title sub key words
    @staticmethod
    def __step1_get_title_keywords(all_text, line_number):
        para = all_text[line_number - 1]
        title = ""
        for run in para.runs:
            w = run.text
            if not w.isdigit():
                title += w
        title_keywords = []
        for w in title.split():
            title_keywords += w.split("-")
        if not title_keywords:
            title_keywords = [""]
        return title_keywords

    @staticmethod
    def score(text: str, all_text: tuple[str], line_number: int) -> int:

        if text.find("GB") == -1 or text.find("字符") == -1:
            return 0
        score = 0
        if text.find("验证") == 0:
            score += 1
            match_group = re.search(r"在(.*)中", text)
            if match_group:
                score += 1
                match_txt = match_group.group(1)
                title_keywords = TextChecker1.__step1_get_title_keywords(
                    all_text, line_number
                )
                for w in title_keywords:
                    if match_txt.find(w) != -1:
                        score += 1

            if (
                text.find("正确") != -1
                or text.find("处理") != -1
                or text.find("显示") != -1
            ):
                score += 1
            if text.find("字符") != -1:
                score += 1
                if text.find("定义的") != -1:
                    score += 1
        return score

    def guess(text: str, all_text: tuple[str], line_number: int) -> str:
        title_keywords = TextChecker1.__step1_get_title_keywords(all_text, line_number)
        return f"验证：在“{title_keywords[-1]}”功能中能够正确处理显示GB18030-2022中定义的字符"

    def perfect_match(text: str, all_text: tuple[str], line_number: int) -> bool:
        return (
            re.findall(
                r"验证：在“.*”功能中能够正确处理显示GB18030-2022中定义的字符", text
            )
            != []
        )


class TextChecker2(BasicTextChecker):
    @staticmethod
    def score(text: str, all_text: tuple[str], line_number: int) -> int:
        if not text.strip().lower().endswith("Outlook"):
            return 0
        score = 0
        if text.find("有效") == -1 or text.find("账号") == -1:
            score += 1
        if text.find("启动") == 0:
            score += 1
            if text.lower().find("outlook") != -1:
                score += 1
        return score

    def guess(text: str, all_text: tuple[str], line_number: int) -> str:
        return "以有效测试账号启动Outlook"

    def perfect_match(text: str, all_text: tuple[str], line_number: int) -> bool:
        return text == "以有效测试账号启动Outlook"
