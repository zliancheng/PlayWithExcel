using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PlayWithExcel.Model
{
    public class CalendarEvent
    {
        public DateTime StartTime { get; set; }
        public int HalfHourCount { get; set; }
        public DayOfWeek Weekday { get; set; }
        public string Memo { get; set; }

        public CalendarEvent(DateTime startTime, int halfHourCount, DayOfWeek weekday, string memo)
        {
            this.StartTime = startTime;
            this.HalfHourCount = halfHourCount;
            this.Weekday = weekday;
            this.Memo = memo;
        }
        public DateTime EndTime
        {
            get { return StartTime.AddMinutes(30*HalfHourCount); }
        }
    }
}
