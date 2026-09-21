"""Exercise the installed extension, never a module from the source tree."""
from pathlib import Path
import struct
import sys
import tempfile
import turboxl
import turboxl._turboxl as native
from fixtures import workbook


def main():
    expected_bits = int(sys.argv[1]) if len(sys.argv) > 1 else 64
    assert struct.calcsize('P') * 8 == expected_bits
    assert Path(turboxl.__file__).resolve().parent != Path(__file__).resolve().parents[2]
    assert Path(native.__file__).resolve().parent == Path(turboxl.__file__).resolve().parent
    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / 'conversion.xlsx'
        workbook(path)
        filename = str(path)
        assert turboxl.read_sheet_to_csv(filename) == '"café, ""quoted""",42,2024-01-15,2024-01-15T13:45:30\n'
        assert turboxl.read_sheet_to_csv(filename, 1) == 'secret\n'
        typed = turboxl._read_sheet_to_python(filename, 0)
        assert typed == [['café, "quoted"', 42.0, 45306.0, 45306.57326388889]]
        assert isinstance(typed[0][1], float)
        sparse_typed = turboxl._read_sheet_to_python(filename, 'Sparse')
        assert len(sparse_typed) == 50
        assert all(len(row) == 26 for row in sparse_typed)
        assert sparse_typed[4][3] == 'gap'
        assert sparse_typed[49][25] == 'far'
        assert turboxl._read_sheet_to_python(filename, 'Sparse', nrows=0) == []
        bounded_typed = turboxl._read_sheet_to_python(
            filename, 'Sparse', nrows=5
        )
        assert len(bounded_typed) == 5
        assert all(len(row) == 4 for row in bounded_typed)
        cropped_typed = turboxl._read_sheet_to_python(
            filename, 'Sparse', skip_empty_area=True
        )
        assert len(cropped_typed) == 50
        assert all(len(row) == 26 for row in cropped_typed)
        try:
            turboxl._read_sheet_to_python(filename, 'Sparse', max_cells=10)
        except RuntimeError as error:
            assert 'max_cells=10' in str(error)
        else:
            raise AssertionError('Accepted typed extraction above max_cells')
        for kwargs in ({'nrows': -1}, {'max_cells': 0}):
            try:
                turboxl._read_sheet_to_python(filename, 'Sparse', **kwargs)
            except ValueError:
                pass
            else:
                raise AssertionError(f'Accepted invalid typed options: {kwargs}')
        output = Path(tmp) / 'conversion.csv'
        output.write_text('old')
        turboxl.read_sheet_to_file(filename, output)
        assert output.read_bytes() == b'"caf\xc3\xa9, ""quoted""",42,2024-01-15,2024-01-15T13:45:30\n'
        assert turboxl.read_specific_sheet(filename, 'Hidden') == 'secret\n'
        expected_sparse = 'origin\n' + '\n' * 3 + ',,,gap\n' + '\n' * 44 + ',' * 25 + 'far\n'
        assert turboxl.read_specific_sheet(filename, 'Sparse') == expected_sparse
        sheets = turboxl.get_sheet_list(filename)
        assert [(s.name, s.visible) for s in sheets] == [
            ('Chart', True), ('Data', True), ('Hidden', False),
            ('VeryHidden', False), ('Dialog', True), ('Sparse', True),
        ]
        assert [s.kind for s in sheets] == [
            turboxl.SheetKind.CHARTSHEET,
            turboxl.SheetKind.WORKSHEET,
            turboxl.SheetKind.WORKSHEET,
            turboxl.SheetKind.WORKSHEET,
            turboxl.SheetKind.OTHER,
            turboxl.SheetKind.WORKSHEET,
        ]
        assert sheets[2].visibility == turboxl.SheetVisibility.HIDDEN
        assert sheets[3].visibility == turboxl.SheetVisibility.VERY_HIDDEN
        assert [s.name for s in turboxl.get_visible_sheets(filename)] == [
            'Chart', 'Data', 'Dialog', 'Sparse'
        ]
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
