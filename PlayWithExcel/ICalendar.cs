using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PlayWithExcel.Model;

namespace PlayWithExcel
{
    interface ICalendar
    {
        bool AddEvent(DateTime startTime, int thirtyMins, string description, DayOfWeek weekday);
    }
}
