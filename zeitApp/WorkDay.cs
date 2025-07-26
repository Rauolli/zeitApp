namespace zeitApp
{
    public class WorkDay
    {
        public DateTime Date { get; set; }
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public TimeSpan TotalWorkTime { get; set; }
        public TimeSpan TotalWorkTimeFrom6 { get; set; }
        public TimeSpan BreakTime { get; set; }
        public TimeSpan NightWorkTime { get; private set; }

        public TimeSpan NightWorkTimeWithBreak => NightWorkTime - BreakTime;
        public TimeSpan WorkTimeWithBreak => TotalWorkTime - BreakTime;
        public TimeSpan WorkTimeFrom6WithBreak => TotalWorkTimeFrom6 - BreakTime;

        public  bool IsWorkingDay { get; set; }

        public WorkDay(DateTime date, DateTime start, DateTime end)
        {
            IsWorkingDay = true;
            Date = date;
            StartTime = start;
            EndTime = end;
            TotalWorkTime = EndTime - StartTime;
            BreakTime = CalculateBreakTime();
            NightWorkTime = CalculateNightWorkTime();
            TotalWorkTimeFrom6 = CalulateTotalWorkTimeFrom6();
        }

        public WorkDay(DateTime date)
        {
            IsWorkingDay = false;
            Date = date;
            //StartTime = new DateTime(date.Year, date.Month, date.Day, 0, 0, 0);
        }

        private TimeSpan CalulateTotalWorkTimeFrom6()
        {
            int minutesToSechs = 60 - StartTime.Minute;
            DateTime StartTimeFrom6 = StartTime.AddMinutes(minutesToSechs);
            if ((StartTime.Hour < 18|| StartTime.Hour == 17) && StartTime.Hour > 16)
            {
                TotalWorkTimeFrom6 = EndTime - StartTimeFrom6;
            }
            else
            {
                TotalWorkTimeFrom6 = TotalWorkTime;
            }
            return TotalWorkTimeFrom6;
        }

        public string GetTotalWorkTimeFormatted() => $"{TotalWorkTime.Hours}:{TotalWorkTime.Minutes}";
        public string GetWorkTimeWithBreakFormatted() => $"{WorkTimeWithBreak.Hours}:{WorkTimeWithBreak.Minutes}";
        public string GetTotalWorkTimeFrom6Formatted() => $"{TotalWorkTimeFrom6.Hours}:{TotalWorkTimeFrom6.Minutes}";
        public string GetWorkTimeFrom6WithBreakFormatted() => $"{WorkTimeFrom6WithBreak.Hours}:{WorkTimeFrom6WithBreak.Minutes}";
        public string GetBreakTimeFormatted() => $"{BreakTime.Hours}:{BreakTime.Minutes}";
        public string GetNightWorkTimeFormatted() => $"{NightWorkTime.Hours}:{NightWorkTime.Minutes}";
        public string GetNightWorkTimeWithBreakFormatted() => $"{NightWorkTimeWithBreak.Hours}:{NightWorkTimeWithBreak.Minutes}";

        private TimeSpan CalculateBreakTime()
        {
            // Berechnung der Pausenzeit, z.B. wenn die Arbeitszeit mehr als 6 Stunden beträgt
            TimeSpan greaterThanSixHours = new TimeSpan(7, 0, 0);
            TimeSpan greaterThanFiveHour = new TimeSpan(5, 0, 0);
            if (TotalWorkTime > greaterThanSixHours)
            {
                 BreakTime = TimeSpan.FromMinutes(45);
            }
            else if (TotalWorkTime > greaterThanFiveHour)
            {
                 BreakTime = TimeSpan.FromMinutes(15);
            }
            else
            {
                BreakTime = TimeSpan.Zero;
            }
            return BreakTime;
        }

        private TimeSpan CalculateNightWorkTime()
        {
            TimeSpan nightWork = TimeSpan.Zero;

            // Definiere den Nachtarbeitszeitraum: 22:00 bis 06:00 Uhr
            DateTime nightStart = StartTime.Date.AddHours(22);   // 22:00 Uhr des Tages
            DateTime nightEnd = StartTime.Date.AddDays(1).AddHours(6); // 06:00 Uhr des nächsten Tages

            // Überprüfe, ob der Arbeitsbeginn in den Nachtzeitraum fällt
            if (StartTime < nightEnd && EndTime > nightStart)
            {
                // Berechne den tatsächlichen Start der Nachtarbeitszeit
                DateTime actualStart = StartTime > nightStart ? StartTime : nightStart;

                // Berechne das tatsächliche Ende der Nachtarbeitszeit
                DateTime actualEnd = EndTime < nightEnd ? EndTime : nightEnd;

                // Berechne die gearbeitete Zeit im Nachtzeitraum
                nightWork = actualEnd - actualStart;
            }

            return nightWork;
        }
    }
}