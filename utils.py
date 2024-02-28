def printing(*txt, **kwargs):
    if "file" in kwargs:
        print(*txt, file=kwargs["file"])
        del kwargs["file"]
    print(*txt, **kwargs)
