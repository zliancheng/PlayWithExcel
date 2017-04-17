using System;
using System.Collections.Generic;
using System.Linq;
using PlayWithExcel.Model;

namespace PlayWithExcel
{
    public class WeekCalendar : ICalendar
    {
        public DateTime StartDate { get; set; }
        public List<CalendarEvent> Events { get; set; }
        public WeekCalendar(DateTime StartDate)
        {
            this.StartDate = StartDate;
            this.Events=new List<CalendarEvent>();
        }


        public bool AddEvent(DateTime startTime, int thirtyMins, string description, DayOfWeek weekday)
        {
            if (Events.Any(e => e.Weekday == weekday && e.StartTime < startTime && startTime < e.EndTime)) return false;
            Events.Add(new CalendarEvent(startTime, thirtyMins, weekday, description ));
            return true;
        }

    }
}
