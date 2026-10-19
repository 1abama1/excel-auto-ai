if __name__ == "__main__":
    from multiprocessing import freeze_support
    freeze_support()
    from document_excel.main import main
    raise SystemExit(main())
