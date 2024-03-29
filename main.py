import argparse
from os.path import expanduser

from utils import printing

from docx import Document
from text_checker import TextChecker1, TextChecker2
from document_checker import (
    DocumentChecker1,
    DocumentChecker2,
    DocumentChecker3,
    DocumentChecker4,
    DocumentChecker5,
    DocumentChecker6,
    DocumentChecker7,
    DocumentChecker8,
    DocumentChecker9,
    DocumentChecker10,
)


def get_all_text(doc: Document) -> tuple[str]:
    all_text = []
    for para in doc.paragraphs:
        all_text.append(para)
    return all_text


TEXT_CHECKER_LIST = [TextChecker1(), TextChecker2()]
DOCUMENT_CHECKER_LIST = [
    DocumentChecker1(),
    DocumentChecker2(),
    DocumentChecker3(),
    DocumentChecker4(),
    DocumentChecker7(),
    DocumentChecker8(),
]

EXTRA_CHECKER_LIST = [
    DocumentChecker5(),
    DocumentChecker6(),
    DocumentChecker9(),
    DocumentChecker10(),
]


def main(
    path,
    output,
    logfile=None,
    extra=None,
    extra_only=False,
    default_yes=False,
    skip_change=False,
    ask_guess_replace=False,
):
    # init
    for checker in TEXT_CHECKER_LIST + DOCUMENT_CHECKER_LIST + EXTRA_CHECKER_LIST:
        checker.logfile = logfile
        checker.default_yes = default_yes

    for checker in TEXT_CHECKER_LIST:
        checker.skip_change = skip_change
        checker.ask_guess_is_right = ask_guess_replace

    # process
    doc = Document(path)
    all_text = get_all_text(doc)
    changed = False
    printing("Auto-fixing document", file=logfile)
    for line_number, paragraph in enumerate(all_text):
        printing(line_number, paragraph.text, file=logfile)
        paragraph = all_text[line_number]
        for i in range(len(EXTRA_CHECKER_LIST)):
            if extra and i in extra:
                checker = EXTRA_CHECKER_LIST[i]
                para = checker.process(paragraph, all_text, line_number)
                if para:
                    changed = True
        for checker in DOCUMENT_CHECKER_LIST:
            if extra_only:
                continue
            para = checker.process(paragraph, all_text, line_number)
            if para:
                changed = True
    if changed:
        answer = input(f"Save document? (To {output}) (Y/n)")
        answer = answer.strip().lower()
        if answer == "y" or not answer:
            doc.save(output)
    printing("--" * 10, file=logfile)
    printing("Manual fixing document", file=logfile)
    for line_number, paragraph in enumerate(all_text):
        printing(line_number, paragraph.text, file=logfile)
        paragraph = all_text[line_number]
        for checker in TEXT_CHECKER_LIST:
            if extra_only:
                continue
            txt = checker.process(paragraph, all_text, line_number)
            if txt:
                changed = True
    if changed:
        answer = input(f"Save document? (To {output}) (Y/n)")
        answer = answer.strip().lower()
        if answer == "y" or not answer:
            doc.save(output)
            # if txt:
            #    inline = paragraph.runs
            #    for i in range(len(inline)):
            #        printing(i, inline[i].text)
            #        if old_text in inline[i].text:
            #            printing(f"Old: {old_text}\nNew: {txt}")
            #            answer = input("Replace it? (Y/n)")
            #            answer = answer.strip().lower()
            #            if answer == "y" or not answer:
            #                text = inline[i].text.replace(old_text, txt)
            #                inline[i].text = text


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Process some files.")
    parser.add_argument("--path", required=True, help="Path to the input file")
    parser.add_argument(
        "--output", default="output.docx", help="Path to the output file"
    )
    parser.add_argument("--logfile", default="logfile.txt", help="Path to the log file")
    parser.add_argument(
        "--extra", nargs="+", type=int, default=None, help="Select extra checkers"
    )
    parser.add_argument(
        "--extra-only", action="store_true", help="Run only extra checkers"
    )
    parser.add_argument("-y", action="store_true", help="Default answer is yes")
    parser.add_argument(
        "--skip-change",
        action="store_true",
        help="Skip change, only print guess to log",
    )
    parser.add_argument(
        "--ask-guess-replace",
        action="store_true",
        help="If guess is not match, ask user to type the correct text (Or answer yes/no)",
    )

    args = parser.parse_args()
    path = expanduser(args.path)
    output = expanduser(args.output)
    logfile_path = expanduser(args.logfile)

    printing("Path:", path)
    printing("Output:", output)
    printing("Logfile:", logfile_path)
    printing("Extra checkers:", args.extra)
    printing("Extra only:", args.extra_only)

    with open(logfile_path, "w+", encoding="utf-8") as logfile:
        main(
            path,
            output,
            logfile,
            extra=[i - 1 for i in args.extra] if args.extra else None,
            extra_only=args.extra_only,
            default_yes=args.y,
            skip_change=args.skip_change,
            ask_guess_replace=args.ask_guess_replace,
        )

    if logfile_path:
        printing(f"Log file: {logfile_path}")
