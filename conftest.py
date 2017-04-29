import pytest
import os.path
from fixture.application import Application1
from model.group import Group
fixture = None

import clr
clr.AddReferenceByName('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
from Microsoft.Office.Interop import Excel

@pytest.fixture()
def app(request):
    global fixture
    if fixture is None:
        fixture = Application1()
    return fixture


@pytest.fixture(scope="session", autouse=True)
def stop(request):
    def fin():
        fixture.close_app()
    request.addfinalizer(fin)
    return fixture


def pytest_generate_tests(metafunc):
    for fixture in metafunc.fixturenames:
        if fixture.startswith("xlsx_"):
            testdata = load_from_xlsx(fixture[5:])
            metafunc.parametrize(fixture, testdata, ids=[str(x) for x in testdata])


def load_from_xlsx(f):
    file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data\%s.xlsx" % f)
    excel = Excel.ApplicationClass()
    workbook = excel.Workbooks.Open(r"%s" % file)
    sheet = workbook.Worksheets[1]
    testdata = []
    for i in range(2,4):
        text = sheet.Rows[i].Value2[0,0]
        testdata.append(Group(name = text))
    excel.Quit()
    return testdata