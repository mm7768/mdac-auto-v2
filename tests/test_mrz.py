import unittest
from datetime import date

import numpy as np

from mrz.checksum import calculate_check_digit, validate_td3_checksums
from mrz.corrector import generate_correction_candidates
from mrz.engine import MRZEngine
from mrz.legacy import MRZParser
from mrz.models import OCRLine
from mrz.ocr import PaddleOCRBackend, extract_ocr_lines
from mrz.parser import TD3Parser
from mrz.preprocess import generate_preprocessed_variants


LINE_1 = "P<UTOERIKSSON<<ANNA<MARIA<<<<<<<<<<<<<<<<<<<"
LINE_2 = "L898902C36UTO7408122F1204159ZE184226B<<<<<10"


class StaticBackend:
    def __init__(self, lines, score=1.0):
        self.lines = lines
        self.score = score

    def recognize(self, image, variant=""):
        return [OCRLine(text=text, score=self.score, variant=variant) for text in self.lines]


class MRZTests(unittest.TestCase):
    def test_known_td3_checksum_and_field_layout(self):
        report = validate_td3_checksums(LINE_2)
        self.assertTrue(report.all_valid)
        self.assertEqual(calculate_check_digit("L898902C3"), "6")
        result = TD3Parser(today=date(2026, 8, 15)).parse(LINE_1, LINE_2)
        self.assertIsNotNone(result)
        self.assertEqual(result.passport_number, "L898902C3")
        self.assertEqual(result.surname, "ERIKSSON")
        self.assertEqual(result.given_names, "ANNA MARIA")
        self.assertEqual(result.nationality, "UTO")
        self.assertEqual(result.date_of_birth, "1974-08-12")
        self.assertEqual(result.sex, "F")
        self.assertEqual(result.expiry_date, "2012-04-15")
        self.assertEqual(result.personal_number, "ZE184226B")
        self.assertTrue(result.checksum_valid)

    def test_invalid_checksum_is_not_accepted_by_parser(self):
        broken = LINE_2[:9] + ("7" if LINE_2[9] != "7" else "8") + LINE_2[10:]
        result = TD3Parser().parse(LINE_1, broken)
        self.assertIsNotNone(result)
        self.assertFalse(result.checksum_valid)
        self.assertIn("passport_number", result.rejection_reason)

    def test_common_ocr_error_is_repaired_by_checksum(self):
        broken = LINE_2.replace("0", "O", 1)
        candidates = generate_correction_candidates(LINE_1, broken, max_edits=2)
        self.assertTrue(candidates)
        self.assertTrue(any(candidate.line_2 == LINE_2 for candidate in candidates))

    def test_engine_accepts_only_high_confidence_valid_candidate(self):
        backend = StaticBackend([LINE_1, LINE_2], score=1.0)
        engine = MRZEngine(ocr_backend=backend, max_variants=2, min_confidence=90)
        image = np.zeros((240, 400, 3), dtype=np.uint8)
        result = engine.parse(image)
        self.assertIsNotNone(result)
        self.assertTrue(result.accepted)
        self.assertTrue(result.checksum_valid)
        self.assertGreaterEqual(result.confidence, 90)
        self.assertEqual(result.passport_number, "L898902C3")

    def test_engine_groups_two_mrz_lines_using_boxes(self):
        class BoxBackend:
            def recognize(self, image, variant=""):
                return [
                    OCRLine(LINE_1, score=1.0, box=((0, 0), (100, 0), (100, 20), (0, 20))),
                    OCRLine(LINE_2, score=1.0, box=((0, 30), (100, 30), (100, 50), (0, 50))),
                ]

        engine = MRZEngine(ocr_backend=BoxBackend(), max_variants=1, min_confidence=90)
        result = engine.parse(np.zeros((240, 400, 3), dtype=np.uint8))
        self.assertIsNotNone(result)
        self.assertTrue(result.accepted)
        self.assertEqual(result.passport_number, "L898902C3")

    def test_engine_repairs_common_error_and_accepts_result(self):
        broken = LINE_2.replace("0", "O", 1)
        backend = StaticBackend([LINE_1, broken], score=1.0)
        engine = MRZEngine(ocr_backend=backend, max_variants=1, min_confidence=90)
        image = np.zeros((240, 400, 3), dtype=np.uint8)
        result = engine.parse(image)
        self.assertIsNotNone(result)
        self.assertTrue(result.accepted)
        self.assertTrue(result.checksum_valid)
        self.assertEqual(result.passport_number, "L898902C3")
        self.assertTrue(result.corrections)

    def test_engine_rejects_low_ocr_confidence(self):
        backend = StaticBackend([LINE_1, LINE_2], score=0.5)
        engine = MRZEngine(ocr_backend=backend, max_variants=1, min_confidence=90)
        image = np.zeros((240, 400, 3), dtype=np.uint8)
        result = engine.parse(image)
        self.assertIsNotNone(result)
        self.assertFalse(result.accepted)
        self.assertIn("confidence_below_threshold", result.rejection_reason)

    def test_ocr_adapter_handles_numpy_v3_output(self):
        raw = {
            "rec_texts": np.array([LINE_1, LINE_2]),
            "rec_scores": np.array([0.98, 0.97]),
            "rec_boxes": np.array([[0, 0, 100, 20], [0, 30, 100, 50]]),
        }
        lines = extract_ocr_lines(raw, variant="test")
        self.assertEqual([line.text for line in lines], [LINE_1, LINE_2])
        self.assertAlmostEqual(lines[0].score, 0.98)
        self.assertEqual(lines[0].variant, "test")

    def test_ocr_adapter_handles_generator_output(self):
        raw = ({"rec_texts": [text], "rec_scores": [0.9]} for text in (LINE_1, LINE_2))
        lines = extract_ocr_lines(raw, variant="generator")
        self.assertEqual([line.text for line in lines], [LINE_1, LINE_2])

    def test_paddle_backend_calls_predict_and_extracts_lines(self):
        class FakeEngine:
            def predict(self, image):
                return {
                    "rec_texts": [LINE_1, LINE_2],
                    "rec_scores": [0.96, 0.95],
                    "rec_boxes": [[0, 0, 100, 20], [0, 30, 100, 50]],
                }

        backend = PaddleOCRBackend(engine=FakeEngine())
        lines = backend.recognize(np.zeros((80, 120, 3), dtype=np.uint8), variant="fake")
        self.assertEqual(len(lines), 2)
        self.assertEqual(lines[1].text, LINE_2)
        self.assertAlmostEqual(lines[1].score, 0.95)

    def test_preprocess_generates_bounded_variants(self):
        image = np.zeros((240, 400, 3), dtype=np.uint8)
        variants = generate_preprocessed_variants(image, max_variants=8)
        self.assertGreaterEqual(len(variants), 1)
        self.assertLessEqual(len(variants), 8)
        self.assertTrue(all(variant.image.size > 0 for variant in variants))
        self.assertTrue(any(variant.name.startswith("bottom35") for variant in variants))

    def test_legacy_facade_keeps_existing_fields(self):
        class FakeEngine:
            def parse(self, image):
                return type(
                    "R",
                    (),
                    {
                        "accepted": True,
                        "passport_number": "L898902C3",
                        "surname": "ERIKSSON",
                        "given_names": "ANNA MARIA",
                        "nationality": "UTO",
                        "date_of_birth": "1974-08-12",
                        "sex": "F",
                        "expiry_date": "2012-04-15",
                        "personal_number": "ZE184226B",
                        "mrz_line_1": LINE_1,
                        "mrz_line_2": LINE_2,
                        "confidence": 95.0,
                        "checksum_valid": True,
                    },
                )()

        success, result = MRZParser(engine=FakeEngine()).parse_image(b"ignored")
        self.assertTrue(success)
        self.assertEqual(result["passport"], "L898902C3")
        self.assertEqual(result["name"], "ERIKSSON ANNA MARIA")
        self.assertEqual(result["sex_text"], "女")
        self.assertEqual(result["passport_number"], "L898902C3")


if __name__ == "__main__":
    unittest.main()
