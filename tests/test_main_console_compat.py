import importlib
import sys
import types
import unittest
from pathlib import Path


class MainConsoleCompatibilityTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        root = str(Path(__file__).resolve().parents[1])
        if root not in sys.path:
            sys.path.insert(0, root)

        tkinter = types.ModuleType("tkinter")
        tkinter.W = "W"
        tkinter.NW = "NW"
        tkinter.LEFT = "LEFT"
        tkinter.RIGHT = "RIGHT"
        tkinter.NORMAL = "NORMAL"
        tkinter.DISABLED = "DISABLED"
        tkinter.END = "END"
        tkinter.X = "X"
        tkinter.BOTH = "BOTH"
        tkinter.WORD = "WORD"
        tkinter.Module = types.ModuleType
        tkinter_ttk = types.ModuleType("tkinter.ttk")
        tkinter_filedialog = types.ModuleType("tkinter.filedialog")
        tkinter_messagebox = types.ModuleType("tkinter.messagebox")
        tkinter_scrolledtext = types.ModuleType("tkinter.scrolledtext")
        tkinter.ttk = tkinter_ttk
        tkinter.filedialog = tkinter_filedialog
        tkinter.messagebox = tkinter_messagebox
        tkinter.scrolledtext = tkinter_scrolledtext
        for submodule in (tkinter, tkinter_ttk, tkinter_filedialog, tkinter_messagebox, tkinter_scrolledtext):
            sys.modules.setdefault(submodule.__name__, submodule)

        ddddocr = types.ModuleType("ddddocr")
        ddddocr.DdddOcr = type("DdddOcr", (), {})
        sys.modules.setdefault("ddddocr", ddddocr)

        filelock = types.ModuleType("filelock")
        filelock.FileLock = type("FileLock", (), {})
        sys.modules.setdefault("filelock", filelock)

        telebot = types.ModuleType("telebot")

        class TeleBot:
            def __init__(self, *args, **kwargs):
                pass

            def message_handler(self, *args, **kwargs):
                return lambda func: func

        telebot.TeleBot = TeleBot
        sys.modules.setdefault("telebot", telebot)

        playwright = types.ModuleType("playwright")
        playwright_sync = types.ModuleType("playwright.sync_api")
        playwright_sync.sync_playwright = lambda: None
        sys.modules.setdefault("playwright", playwright)
        sys.modules.setdefault("playwright.sync_api", playwright_sync)

        cls.module = importlib.import_module("main_console")

    def test_mrz_facade_is_imported_and_legacy_classes_remain(self):
        from mrz import MRZParser

        self.assertIs(self.module.MRZParser, MRZParser)
        for name in (
            "ExcelManager",
            "TelegramBot",
            "PDFMRZProcessor",
            "PDFTelegramBot",
            "GmailPINFetcher",
            "MDACApp",
            "process_registration",
        ):
            self.assertTrue(hasattr(self.module, name), name)

    def test_old_ocr_classes_are_removed_without_breaking_parser_name(self):
        self.assertFalse(hasattr(self.module, "_PaddleOCRAdapter"))
        self.assertFalse(hasattr(self.module, "PaddleMRZParser"))


if __name__ == "__main__":
    unittest.main()
