from utils import printing


class BasicChecker:
    logfile = None
    default_yes = False

    def process(self, para, all_text: tuple[str], line_number: int) -> str | None:
        pass

    @staticmethod
    def score(obj, all_text: tuple[str], line_number: int) -> int:
        pass

    @staticmethod
    def guess(obj, all_text: tuple[str], line_number: int) -> str:
        pass

    @staticmethod
    def perfect_match(obj, all_text: tuple[str], line_number: int) -> bool:
        pass

    def ask_for_process(self, input_text: str, file=None) -> bool:
        file = file or self.logfile
        printing(input_text, file=file, end="")
        if self.default_yes:
            printing("Y (default)", file=file)
            return True
        answer = input()
        if answer == "y" or not answer:
            printing("Y", file=file)
            return True
        return False

    def ask_for_answer(self, input_text: str, file=None) -> str:
        file = file or self.logfile
        printing(input_text, file=file, end="")
        answer = input()
        printing(answer, file=file)
        return answer
