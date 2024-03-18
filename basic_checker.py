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

    @staticmethod
    def ask_for_process(self, input_text: str, **kwargs) -> bool:
        if self.default_yes:
            printing("Y (default)", file=self.logfile)
            return True
        answer = input(input_text, **kwargs)
        if answer == "y" or not answer:
            printing("Y", file=self.logfile)
            return True
        return False
