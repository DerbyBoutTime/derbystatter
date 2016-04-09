import os

from derbystatter import statbook
from derbystatter import brecre
from derbystatter.statbook import Team

APP_ROOT = os.path.abspath(os.path.dirname(__file__))


def _testdata(path):
    return os.path.join(APP_ROOT, 'data', path)


def test_import():
    book = statbook.StatBook(_testdata('testbook.xlsx'))
    assert book.version is not None


def test_brecre(capsys):
    book = statbook.StatBook(_testdata('testbook.xlsx'))
    brecre.check_pt(book, verbose=False)
    out, _ = capsys.readouterr()
    assert 'Ohio' in out
    assert 'Detroit' in out
    assert 'Error' not in out
    brecre.check_sk(book, verbose=False)
    out, _ = capsys.readouterr()
    assert 'Error' in out

