using System;

namespace PlayWithExcel
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var monday = GetNextMonday();
            var calendar = new WeekCalendar(monday);
            calendar.AddEvent(DateTime.Parse("08:30AM"), 2, "Morning meeting", DayOfWeek.Monday);
            calendar.AddEvent(DateTime.Parse("04:00PM"), 2, "Retrospective meeting", DayOfWeek.Monday);
            calendar.AddEvent(DateTime.Parse("10:00AM"), 3, "Billing meeting", DayOfWeek.Tuesday);
            calendar.AddEvent(DateTime.Parse("4:00PM"), 2, "Home Health Gold meeting", DayOfWeek.Thursday);
            calendar.AddEvent(DateTime.Parse("9:00AM"), 6, "Hackthong", DayOfWeek.Friday);

            var exporter = new SampleExporter(calendar.Events, calendar.StartDate.ToShortDateString());
            exporter.Generate();
        }

        public static DateTime GetNextMonday()
        {
            DateTime today = DateTime.Today;
            return today.AddDays(((int) DayOfWeek.Monday - (int) today.DayOfWeek + 7)%7);
        }
    }
}
