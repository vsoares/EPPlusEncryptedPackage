﻿using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using Rhino.Mocks;
using OfficeOpenXml.FormulaParsing;

namespace EPPlusTest.FormulaParsing.IntegrationTests.BuiltInFunctions
{
    [TestClass]
    public class RefAndLookupTests : FormulaParserTestBase
    {
        private ExcelDataProvider _excelDataProvider;
        const string WorksheetName = null;
        private ExcelPackage _package;
        private ExcelWorksheet _worksheet;

        [TestInitialize]
        public void Initialize()
        {
            _excelDataProvider = MockRepository.GenerateStub<ExcelDataProvider>();
            _excelDataProvider.Stub(x => x.GetDimensionEnd(Arg<string>.Is.Anything)).Return(new ExcelCellAddress(10, 1));
            _parser = new FormulaParser(_excelDataProvider);
            _package = new ExcelPackage();
            _worksheet = _package.Workbook.Worksheets.Add("Test");
        }

        [TestCleanup]
        public void Cleanup()
        {
            _package.Dispose();
        }

        [TestMethod]
        public void VLookupShouldReturnCorrespondingValue()
        {
            using(var pck = new ExcelPackage())
            {
                var ws = pck.Workbook.Worksheets.Add("test");
                var lookupAddress = "A1:B2";
                ws.Cells["A1"].Value = 1;
                ws.Cells["B1"].Value = 1;
                ws.Cells["A2"].Value = 2;
                ws.Cells["B2"].Value = 5;
                ws.Cells["A3"].Formula = "VLOOKUP(2, " + lookupAddress + ", 2)";
                ws.Calculate();
                var result = ws.Cells["A3"].Value;
                Assert.AreEqual(5, result);
            }
        }

        [TestMethod]
        public void VLookupShouldReturnClosestValueBelowIfLastArgIsTrue()
        {
            using (var pck = new ExcelPackage())
            {
                var ws = pck.Workbook.Worksheets.Add("test");
                var lookupAddress = "A1:B2";
                ws.Cells["A1"].Value = 3;
                ws.Cells["B1"].Value = 1;
                ws.Cells["A2"].Value = 5;
                ws.Cells["B2"].Value = 5;
                ws.Cells["A3"].Formula = "VLOOKUP(4, " + lookupAddress + ", 2, true)";
                ws.Calculate();
                var result = ws.Cells["A3"].Value;
                Assert.AreEqual(1, result);
            }
        }

        [TestMethod]
        public void HLookupShouldReturnCorrespondingValue()
        {
            var lookupAddress = "A1:B2";
            _worksheet.Cells["A1"].Value = 1;
            _worksheet.Cells["B1"].Value = 2;
            _worksheet.Cells["A2"].Value = 2;
            _worksheet.Cells["B2"].Value = 5;
            _worksheet.Cells["A3"].Formula = "HLOOKUP(2, " + lookupAddress + ", 2)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A3"].Value;
            Assert.AreEqual(5, result);
        }

        [TestMethod]
        public void HLookupShouldReturnClosestValueBelowIfLastArgIsTrue()
        {
            var lookupAddress = "A1:B2";
            _excelDataProvider.Stub(x => x.GetDimensionEnd(Arg<string>.Is.Anything)).Return(new ExcelCellAddress(5, 5));
            _excelDataProvider.Stub(x => x.GetCellValue(WorksheetName,1, 1)).Return(3);
            _excelDataProvider.Stub(x => x.GetCellValue(WorksheetName,1, 2)).Return(5);
            _excelDataProvider.Stub(x => x.GetCellValue(WorksheetName,2, 1)).Return(1);
            _excelDataProvider.Stub(x => x.GetCellValue(WorksheetName,2, 2)).Return(2);
            var result = _parser.Parse("HLOOKUP(4, " + lookupAddress + ", 2, true)");
            Assert.AreEqual(1, result);
        }

        [TestMethod]
        public void LookupShouldReturnMatchingValue()
        {
            var lookupAddress = "A1:B2";
            _excelDataProvider.Stub(x => x.GetCellValue(WorksheetName,1, 1)).Return(3);
            _excelDataProvider.Stub(x => x.GetCellValue(WorksheetName,1, 2)).Return(5);
            _excelDataProvider.Stub(x => x.GetCellValue(WorksheetName,2, 1)).Return(4);
            _excelDataProvider.Stub(x => x.GetCellValue(WorksheetName,2, 2)).Return(1);
            var result = _parser.Parse("LOOKUP(4, " + lookupAddress + ")");
            Assert.AreEqual(1, result);
        }

        [TestMethod]
        public void MatchShouldReturnIndexOfMatchingValue()
        {
            var lookupAddress = "A1:A2";
            _excelDataProvider.Stub(x => x.GetCellValue(WorksheetName,1, 1)).Return(3);
            _excelDataProvider.Stub(x => x.GetCellValue(WorksheetName,1, 2)).Return(5);
            var result = _parser.Parse("MATCH(3, " + lookupAddress + ")");
            Assert.AreEqual(1, result);
        }

        [TestMethod]
        public void RowShouldReturnRowNumber()
        {
            _excelDataProvider.Stub(x => x.GetRangeFormula("", 4, 1)).Return("Row()");
            var result = _parser.ParseAt("A4");
            Assert.AreEqual(4, result);
        }

        [TestMethod]
        public void RowSholdHandleReference()
        {
            using (var package = new ExcelPackage())
            {
                var s1 = package.Workbook.Worksheets.Add("test");
                s1.Cells["A1"].Formula = "ROW(A4)";
                s1.Calculate();
                Assert.AreEqual(4, s1.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void ColumnShouldReturnRowNumber()
        {
            //_excelDataProvider.Stub(x => x.GetRangeValues("B4")).Return(new List<ExcelCell> { new ExcelCell(null, "Column()", 0, 0) });
            _excelDataProvider.Stub(x => x.GetRangeFormula("", 4, 2)).Return("Column()");
            var result = _parser.ParseAt("B4");
            Assert.AreEqual(2, result);
        }

        [TestMethod]
        public void ColumnSholdHandleReference()
        {
            using (var package = new ExcelPackage())
            {
                var s1 = package.Workbook.Worksheets.Add("test");
                s1.Cells["A1"].Formula = "COLUMN(B4)";
                s1.Calculate();
                Assert.AreEqual(2, s1.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void RowsShouldReturnNbrOfRows()
        {
            _excelDataProvider.Stub(x => x.GetRangeFormula("", 4, 1)).Return("Rows(A5:B7)");
            var result = _parser.ParseAt("A4");
            Assert.AreEqual(3, result);
        }

        [TestMethod]
        public void ColumnsShouldReturnNbrOfCols()
        {
            _excelDataProvider.Stub(x => x.GetRangeFormula("", 4, 1)).Return("Columns(A5:B7)");
            var result = _parser.ParseAt("A4");
            Assert.AreEqual(2, result);
        }

        [TestMethod]
        public void ChooseShouldReturnCorrectResult()
        {
            var result = _parser.Parse("Choose(1, \"A\", \"B\")");
            Assert.AreEqual("A", result);
        }

        [TestMethod]
        public void AddressShouldReturnCorrectResult()
        {
            _excelDataProvider.Stub(x => x.ExcelMaxRows).Return(12345);
            var result = _parser.Parse("Address(1, 1)");
            Assert.AreEqual("$A$1", result);
        }

        [TestMethod]
        public void IndirectShouldReturnARange()
        {
            using (var package = new ExcelPackage(new MemoryStream()))
            {
                var s1 = package.Workbook.Worksheets.Add("Test");
                s1.Cells["A1:A2"].Value = 2;
                s1.Cells["A3"].Formula = "SUM(Indirect(\"A1:A2\"))";
                s1.Calculate();
                Assert.AreEqual(4d, s1.Cells["A3"].Value);

                s1.Cells["A4"].Formula = "SUM(Indirect(\"A1:A\" & \"2\"))";
                s1.Calculate();
                Assert.AreEqual(4d, s1.Cells["A4"].Value);
            }
        }

        [TestMethod]
        public void OffsetShouldReturnASingleValue()
        {
            using (var package = new ExcelPackage())
            {
                var s1 = package.Workbook.Worksheets.Add("Test");
                s1.Cells["B3"].Value = 1d;
                s1.Cells["A5"].Formula = "OFFSET(A1, 2, 1)";
                s1.Calculate();
                Assert.AreEqual(1d, s1.Cells["A5"].Value);
            }
        }

        [TestMethod]
        public void OffsetShouldReturnARange()
        {
            using (var package = new ExcelPackage())
            {
                var s1 = package.Workbook.Worksheets.Add("Test");
                s1.Cells["B1"].Value = 1d;
                s1.Cells["B2"].Value = 1d;
                s1.Cells["B3"].Value = 1d;
                s1.Cells["A5"].Formula = "SUM(OFFSET(A1:A3, 0, 1))";
                s1.Calculate();
                Assert.AreEqual(3d, s1.Cells["A5"].Value);
            }
        }

        [TestMethod]
        public void OffsetDirectReferenceToMultiRangeShouldSetValueError()
        {
            using (var package = new ExcelPackage())
            {
                var s1 = package.Workbook.Worksheets.Add("Test");
                s1.Cells["B1"].Value = 1d;
                s1.Cells["B2"].Value = 1d;
                s1.Cells["B3"].Value = 1d;
                s1.Cells["A5"].Formula = "OFFSET(A1:A3, 0, 1)";
                s1.Calculate();
                var result = s1.Cells["A5"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), result);
            }
        }

        [TestMethod]
        public void OffsetShouldReturnARangeAccordingToWidth()
        {
            using (var package = new ExcelPackage())
            {
                var s1 = package.Workbook.Worksheets.Add("Test");
                s1.Cells["B1"].Value = 1d;
                s1.Cells["B2"].Value = 1d;
                s1.Cells["B3"].Value = 1d;
                s1.Cells["A5"].Formula = "SUM(OFFSET(A1:A3, 0, 1, 2))";
                s1.Calculate();
                Assert.AreEqual(2d, s1.Cells["A5"].Value);
            }
        }

        [TestMethod]
        public void OffsetShouldReturnARangeAccordingToHeight()
        {
            using (var package = new ExcelPackage())
            {
                var s1 = package.Workbook.Worksheets.Add("Test");
                s1.Cells["B1"].Value = 1d;
                s1.Cells["B2"].Value = 1d;
                s1.Cells["B3"].Value = 1d;
                s1.Cells["C1"].Value = 2d;
                s1.Cells["C2"].Value = 2d;
                s1.Cells["C3"].Value = 2d;
                s1.Cells["A5"].Formula = "SUM(OFFSET(A1:A3, 0, 1, 2, 2))";
                s1.Calculate();
                Assert.AreEqual(6d, s1.Cells["A5"].Value);
            }
        }

        [TestMethod]
        public void OffsetShouldCoverMultipleColumns()
        {
            using (var package = new ExcelPackage())
            {
                var s1 = package.Workbook.Worksheets.Add("Test");
                s1.Cells["C1"].Value = 1d;
                s1.Cells["C2"].Value = 1d;
                s1.Cells["C3"].Value = 1d;
                s1.Cells["D1"].Value = 2d;
                s1.Cells["D2"].Value = 2d;
                s1.Cells["D3"].Value = 2d;
                s1.Cells["A5"].Formula = "SUM(OFFSET(A1:B3, 0, 2))";
                s1.Calculate();
                Assert.AreEqual(9d, s1.Cells["A5"].Value);
            }
        }

        [TestMethod, Ignore]
        public void VLookupShouldHandleNames()
        {
            using (var package = new ExcelPackage(new FileInfo(@"c:\temp\Book3.xlsx")))
            {
                var s1 = package.Workbook.Worksheets.First();
                var v = s1.Cells["X10"].Formula;
                //s1.Calculate();
                v = s1.Cells["X10"].Formula;
            }
        }
    }
}
