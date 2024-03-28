from typing import Union


def printing(*txt, **kwargs):
    if "file" in kwargs:
        print(*txt, file=kwargs["file"])
        del kwargs["file"]
    print(*txt, **kwargs)


def check_is_bool(s: str) -> Union[bool, None]:
    if s.lower() in ["yes", "y"]:
        return True
    if s.lower() in ["no", "n"]:
        return False
    return None
