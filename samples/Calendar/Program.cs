using Leren.Lerengine;
using Leren.Providers.Reflection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace CalendarSample
{
    class Program
    {
        static void Main(string[] args)
        {
            var engine = new Engine();
            engine.Provider = new ReflectionProvider(GetModel());
            string path = Path.GetTempFileName() + ".xlsx";
            using (FileStream fs = new FileStream(path, FileMode.CreateNew))
            {
                engine.Go(getResourceData(@"CalendarSample.template.xlsx"), fs);
            }
            Process.Start(path);
        }

        static object GetModel()
        {
            DateTime dt = DateTime.Now;

            DateTime currentMonth = new DateTime(dt.Year, dt.Month, 1);
            DateTime currentMonthEnd = new DateTime(dt.Year, dt.Month, DateTime.DaysInMonth(dt.Year, dt.Month));

            var days = new List<string>();

            DateTime start = currentMonth;

            while (start.DayOfWeek != DayOfWeek.Sunday)
                start = start.AddDays(-1);

            DateTime day = start;
            while (day <= currentMonthEnd)
            {
                days.Add(day.Month == currentMonth.Month ? day.Day.ToString() : " ");
                day = day.AddDays(1);
            }

            var weeks = new List<List<string>>();

            while (days.Count > 0)
            {
                weeks.Add(days.Take(7).ToList());
                days.RemoveRange(0, Math.Min(7, days.Count));
            }

            return new {
                Title = dt.ToString("MMM"),
                DayNames = new string[] { "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat" },
                Weeks_ = new List<object> {
                    new { DaysOfWeek = days.Take(7).ToArray() }
                },
                Weeks = (from z in weeks select new { DaysOfWeek = z }).ToArray()
            };
        }

        static private Stream getResourceData(string resourceName)
        {
            var assembly = Assembly.Load("CalendarSample");
            var names = assembly.GetManifestResourceNames();

            return assembly.GetManifestResourceStream(resourceName);
        }
    }
}
