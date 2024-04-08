import re
from typing import Union
from lxml.etree import XPathEvalError
from docx.oxml.ns import nsmap


def printing(*txt, **kwargs):
    if "file" in kwargs and kwargs["file"] is not None:
        print(*txt, file=kwargs["file"])
        del kwargs["file"]
    print(*txt, **kwargs)


def check_is_bool(s: str) -> Union[bool, None]:
    if s.lower() in ["yes", "y"]:
        return True
    if s.lower() in ["no", "n"]:
        return False
    return None


def xpath_auto_ns(xml, path):
    try:
        return xml.xpath(path)
    except XPathEvalError:
        return xml.xpath(path, namespaces=nsmap)
