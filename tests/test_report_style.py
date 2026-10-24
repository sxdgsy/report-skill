import base64
import importlib.util
import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[1]


def load_module(name):
    spec = importlib.util.spec_from_file_location(name, ROOT / 'scripts' / f'{name}.py')
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


style = load_module('report_style')
auditor = load_module('audit_layout')


class ReportStyleTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / 'report.docx'

    def test_blank_metadata_has_real_underlines(self):
        doc = style.create_report()
        doc.save(self.path)
        self.assertEqual(auditor.audit(self.path)['issues'], [])
        for p in doc.paragraphs[1:3]:
            self.assertTrue(any(r.underline and r.text.isspace() for r in p.runs))
        self.assertAlmostEqual(doc.sections[0].page_width.cm, 21, places=2)

    def test_provided_values_are_underlined(self):
        doc = style.create_report(experiment='验证实验', name='测试者', student_id='SAMPLE-ID', class_name='示例班', seat='A')
        for value in ['验证实验', '测试者', 'SAMPLE-ID', '示例班']:
            runs = [r for p in doc.paragraphs for r in p.runs if value in r.text]
            self.assertEqual(len(runs), 1)
            self.assertTrue(runs[0].underline)
        doc.save(self.path)
        self.assertEqual(auditor.audit(self.path)['issues'], [])

    def test_missing_underlines_are_reported(self):
        doc = style.create_report(name='测试者')
        for p in doc.paragraphs:
            for r in p.runs:
                if '测试者' in r.text:
                    r.underline = False
        doc.save(self.path)
        self.assertTrue(any('姓名' in issue for issue in auditor.audit(self.path)['issues']))

    def test_native_image_and_wide_table_have_numbered_captions(self):
        image = Path(self.temp.name) / 'pixel.png'
        image.write_bytes(base64.b64decode('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII='))
        doc = style.create_report()
        style.add_figure(doc, image, '图 1  验证图片', width_cm=1)
        table = style.add_wide_table(doc, ['物理量', '测点 1', '测点 2'], [['U / V', '1.00', '2.00']], '表 1  验证数据')
        doc.save(self.path)
        report = auditor.audit(self.path)
        self.assertEqual(report['issues'], [])
        self.assertEqual((report['figures'], report['tables']), (1, 1))
        width = sum(c.width for c in table.columns)
        section = doc.sections[0]
        self.assertLessEqual(width, section.page_width - section.left_margin - section.right_margin + 1000)
        with ZipFile(self.path) as z:
            xml = z.read('word/document.xml').decode('utf-8')
        self.assertIn('w:tcBorders', xml)

    def test_wrong_table_number_is_reported(self):
        doc = style.create_report()
        style.add_wide_table(doc, ['量', '值'], [['U / V', '1.00']], '表 2  错误编号')
        doc.save(self.path)
        self.assertTrue(any('表 1' in issue for issue in auditor.audit(self.path)['issues']))

    def test_reference_template_has_no_old_images(self):
        with ZipFile(ROOT / 'assets' / 'reference.docx') as z:
            self.assertFalse(any(n.startswith('word/media/') for n in z.namelist()))


if __name__ == '__main__':
    unittest.main()
