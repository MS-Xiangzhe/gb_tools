class BasicChecker:
    logfile = None

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
