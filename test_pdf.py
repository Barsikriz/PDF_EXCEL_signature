import tempfile
import unittest
from pathlib import Path

import pymupdf

from converter import add_signature, signature_box


class PDFTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.source = self.root / "table.pdf"
        self.image = self.root / "signature.png"
        self.target = self.root / "signed.pdf"
        with pymupdf.open() as doc:
            page = doc.new_page(width=180, height=60)
            page.insert_text((10, 35), "TEST SIGNATURE", fontsize=16)
            page.get_pixmap(alpha=True).save(self.image)

    def tearDown(self):
        self.temp.cleanup()

    def make_pdf(self, count=1, collision=False):
        with pymupdf.open() as doc:
            for i in range(count):
                page = doc.new_page()
                page.insert_text((35, 50), f"Table page {i + 1}: ABC 123.45")
            if collision:
                box = signature_box(page, 45)
                page.insert_text((box.x0 + 2, box.y0 + 15), "FOOTER")
            doc.save(self.source)

    def test_preserves_text_and_signs_last_page_only(self):
        for count in (1, 3):
            self.make_pdf(count)
            self.assertEqual(add_signature(self.source, self.image, self.target), count)
            with pymupdf.open(self.source) as before, pymupdf.open(self.target) as after:
                self.assertEqual([p.get_text() for p in before], [p.get_text() for p in after])
                self.assertEqual([len(p.get_images()) for p in after], [0] * (count - 1) + [1])
                bounds = pymupdf.Rect(after[-1].get_image_info()[0]["bbox"])
                box = signature_box(after[-1], 45)
                # PDF coordinates are rounded to single-precision floats.
                tolerance = pymupdf.Rect(box.x0 - .01, box.y0 - .01,
                                         box.x1 + .01, box.y1 + .01)
                self.assertTrue(tolerance.contains(bounds))
            self.source.unlink()
            self.target.unlink()

    def test_refuses_to_cover_footer(self):
        self.make_pdf(collision=True)
        with self.assertRaisesRegex(ValueError, "Место для подписи занято"):
            add_signature(self.source, self.image, self.target)
        self.assertFalse(self.target.exists())

    def test_refuses_pdf_without_text(self):
        with pymupdf.open() as doc:
            doc.new_page()
            doc.save(self.source)
        with self.assertRaisesRegex(ValueError, "без извлекаемого текста"):
            add_signature(self.source, self.image, self.target)


if __name__ == "__main__":
    unittest.main()
