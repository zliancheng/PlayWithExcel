using System;
using System.Collections.Generic;
using System.Linq;
using NPOI.HPSF;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using PlayWithExcel.Model;

namespace PlayWithExcel
{
    public class SampleExporter : Exporter
    {
        private List<CalendarEvent> data;
        private string startDate;

        public SampleExporter(List<CalendarEvent> data, string startDate)
        {
            this.data = data;
            this.startDate = startDate;
            FileName = "Week Calendar";
        }
        protected override void InitializeExcel()
        {
            base.InitializeExcel();
            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Subject = "Sample Report";
            base.WorkBook.SummaryInformation = si;
        }

        protected override void WriteToExcelSpreadsheet()
        {
            var sheet = base.WorkBook.CreateSheet("Sample Form");
            if (this.data == null || data.FirstOrDefault() == null) return;

            //Create a subject area merged all content columns in the first row.
            var mainHeaderRange = new CellRangeAddress(0, 0, 0, 7);
            sheet.AddMergedRegion(mainHeaderRange);

            //Create the subject cell in the subject area and fill with subject value
            //Since the first content row merged all columns, there is only one big cell in the first row.
            string subject = GenerateSubject();
            sheet.CreateRow(0).CreateCell(0).SetCellValue(subject);

            //Create header row and header column.
            var headRow = sheet.CreateRow(1);
            int colIndex = 0;
            headRow.CreateCell(colIndex++).SetCellValue("Time");
            headRow.CreateCell(colIndex++).SetCellValue("Sunnday");
            headRow.CreateCell(colIndex++).SetCellValue("Monday");
            headRow.CreateCell(colIndex++).SetCellValue("Tuesday");
            headRow.CreateCell(colIndex++).SetCellValue("Wednesday");
            headRow.CreateCell(colIndex++).SetCellValue("Thursday");
            headRow.CreateCell(colIndex++).SetCellValue("Friday");
            headRow.CreateCell(colIndex).SetCellValue("Saturday");
            
            int rowIndex = 2;
            DateTime calendarStart = DateTime.Parse("8:00AM");
            for (DateTime time = calendarStart; time <= DateTime.Parse("5:00PM"); time = time.AddMinutes(30))
            {
                sheet.CreateRow(rowIndex++).CreateCell(0).SetCellValue(time.ToString("HH:mm"));
            }

            var cellStyle = WorkBook.CreateCellStyle();
            cellStyle.VerticalAlignment=VerticalAlignment.Center;
            cellStyle.FillForegroundColor = HSSFColor.Grey25Percent.Index;
            cellStyle.FillPattern=FillPattern.SolidForeground;
            cellStyle.WrapText = true;
            rowIndex = 2;
            colIndex = 1;
            foreach (var calendarEvent in data)
            {
                int x = Convert.ToInt32((calendarEvent.StartTime - calendarStart).TotalMinutes/30);
                int y = (int) calendarEvent.Weekday;
                if (calendarEvent.HalfHourCount > 1)
                {
                    var eventRange = new CellRangeAddress(rowIndex + x, rowIndex + x+calendarEvent.HalfHourCount-1, colIndex + y, colIndex + y);
                    sheet.AddMergedRegion(eventRange);
                    var cell = sheet.GetRow(rowIndex + x).CreateCell(colIndex + y);
                    if (cell != null)
                    {
                        cell.SetCellValue(calendarEvent.Memo);
                        cell.CellStyle = cellStyle;
                    }
                }
            }
        }

        private string GenerateSubject()
        {
            return @"Weekly Calendar" +
                   "/n Week tart Date:" + startDate;
        }
    }
}
