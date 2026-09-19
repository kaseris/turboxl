"""Exercise the installed extension, never a module from the source tree."""
from pathlib import Path
import struct
import sys
import tempfile
import turboxl
from fixtures import workbook


def main():
    expected_bits = int(sys.argv[1]) if len(sys.argv) > 1 else 64
    assert struct.calcsize('P') * 8 == expected_bits
    assert Path(turboxl.__file__).resolve().parent != Path(__file__).resolve().parents[2]
    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / 'conversion.xlsx'
        workbook(path)
        filename = str(path)
        assert turboxl.read_sheet_to_csv(filename) == '"café, ""quoted""",42,2024-01-15,2024-01-15T13:45:30\n'
        assert turboxl.read_sheet_to_csv(filename, 1) == 'secret\n'
        assert turboxl.read_specific_sheet(filename, 'Hidden') == 'secret\n'
        expected_sparse = 'origin\n' + '\n' * 3 + ',,,gap\n' + '\n' * 44 + ',' * 25 + 'far\n'
        assert turboxl.read_specific_sheet(filename, 'Sparse') == expected_sparse
        sheets = turboxl.get_sheet_list(filename)
        assert [(s.name, s.visible) for s in sheets] == [('Data', True), ('Hidden', False), ('Sparse', True)]
        assert [s.name for s in turboxl.get_visible_sheets(filename)] == ['Data', 'Sparse']
        options = turboxl.CsvOptions()
        options.date_mode = turboxl.DateMode.RAW
        assert turboxl.read_sheet_to_csv(filename, 0, options) == '"café, ""quoted""",42,45306,45306.573264\n'
        for bad in [Path(tmp) / 'missing.xlsx', Path(tmp) / 'invalid.xlsx']:
            if bad.name == 'invalid.xlsx':
                bad.write_text('not a ZIP archive')
            try:
                turboxl.read_sheet_to_csv(str(bad))
            except RuntimeError:
                pass
            else:
                raise AssertionError(f'Accepted invalid input: {bad}')
    print(f'Installed-wheel conversion tests passed: {sys.version}, {turboxl.__file__}')


if __name__ == '__main__':
    main()
