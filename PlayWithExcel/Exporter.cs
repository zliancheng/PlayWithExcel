using System;
using System.IO;
using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace PlayWithExcel
{
    public abstract class Exporter
    {
        public string FormatType { get; set; }
        public string FileName { get; set; }
        protected HSSFWorkbook WorkBook;
        public Exporter()
        {
            this.FormatType = "XLS";
        }

        protected virtual void InitializeExcel()
        {
            WorkBook = new HSSFWorkbook();
            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "XXX company";
            WorkBook.DocumentSummaryInformation = dsi;

            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Subject = "XXX Report";
            si.Author = "zliancheng";
            WorkBook.SummaryInformation = si;
        }
        protected virtual void WriteToExcelSpreadsheet() { }
        protected MemoryStream ProcessSpreadsheet()
        {
            this.InitializeExcel();
            this.WriteToExcelSpreadsheet();

            var memoryStream = new MemoryStream();
            WorkBook.Write(memoryStream);
            return memoryStream;
        }

        public void AddDropDownListToCell(ISheet sheet, ICell cell, string[] list)
        {
            CellRangeAddressList cellRange = new CellRangeAddressList(cell.RowIndex, cell.RowIndex, cell.ColumnIndex, cell.ColumnIndex);
            DVConstraint constraint = null;
            if (string.Join("", list).Length < 200)
            {
                constraint = DVConstraint.CreateExplicitListConstraint(list);
            }
            else
            {
                var workBook = sheet.Workbook;
                var hiddenSheet = workBook.GetSheet("hidden") ?? workBook.CreateSheet("hidden");
                workBook.SetSheetHidden(workBook.GetSheetIndex("hidden"), SheetState.Hidden);
                var rowsCount = hiddenSheet.PhysicalNumberOfRows;
                for (int i = 0; i < list.Length; i++)
                {
                    hiddenSheet.CreateRow(rowsCount + i).CreateCell(0).SetCellValue(list[i]);
                }
                var formula = string.Format("hidden!$A{0}:$A{1}", rowsCount + 1, rowsCount + list.Length);
                constraint = DVConstraint.CreateFormulaListConstraint(formula);
            }
            HSSFDataValidation validation = new HSSFDataValidation(cellRange, constraint);
            ((HSSFSheet)sheet).AddValidationData(validation);
        }


        public void Generate()
        {
            var bytes = ProcessSpreadsheet().ToArray();
            var filePath = string.Format("../{0}{1}.xls", this.FileName.Replace(".xls", ""), DateTime.Now.ToString("yyyyMMddHHmmss"));
            using (FileStream fsStream = new FileStream(filePath, FileMode.Create))
            {
                foreach (var b in bytes)
                {
                    fsStream.WriteByte(b);
                }
            }
        }
    }
}
