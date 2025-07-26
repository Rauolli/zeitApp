using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace zeitApp
{
    public class WorkMonth
    {
        public List<WorkDay> WorkDays { get; set; } = new List<WorkDay>();

        public WorkMonth(){}

        public WorkMonth(List<WorkDay> workDays)
        {
            WorkDays = workDays;
        }

        public void AddWorkDays(List<WorkDay> workDays)
        {
            WorkDays = workDays;
        }

        public string CalculateTotalWorkTime()
        {
            TimeSpan days = WorkDays.Aggregate(TimeSpan.Zero, (sum, day) => sum + day.TotalWorkTime);
            return $"{(int)days.TotalHours}:{days.Minutes}";
        }
        public string CalculateTotalWorkTimeFrom6()
        {
            TimeSpan days =  WorkDays.Aggregate(TimeSpan.Zero, (sum, day) => sum + day.TotalWorkTimeFrom6);
            return $"{(int)days.TotalHours}:{days.Minutes}";
        }

        public string CalculateTotalBreakTime()
        {
            TimeSpan days = WorkDays.Aggregate(TimeSpan.Zero, (sum, day) => sum + day.BreakTime);
            return $"{(int)days.TotalHours}:{days.Minutes}";
        }

        public string CalculateTotalNightWorkTime()
        {
            TimeSpan days = WorkDays.Aggregate(TimeSpan.Zero, (sum, day) => sum + day.NightWorkTime);
            return $"{(int)days.TotalHours}:{days.Minutes}";
        }
        public string CalculateWorkTimeWithBreak()
        {
            TimeSpan days = WorkDays.Aggregate(TimeSpan.Zero, (sum, day) => sum + day.WorkTimeWithBreak);
            return $"{(int)days.TotalHours}:{days.Minutes}";
        }
        public string CalculateWorkTimeFrom6WithBreak()
        {
            TimeSpan days = WorkDays.Aggregate(TimeSpan.Zero, (sum, day) => sum + day.WorkTimeFrom6WithBreak);
            return $"{(int)days.TotalHours}:{days.Minutes}";
        }

        public string CalculateNightWorkTime()
        {
            TimeSpan days = WorkDays.Aggregate(TimeSpan.Zero, (sum, day) => sum + day.NightWorkTime);
            return $"{(int)days.TotalHours}:{days.Minutes}";
        }
        public string CalculateNightWorkTimeWithBreak()
        {
            TimeSpan days = WorkDays.Aggregate(TimeSpan.Zero, (sum, day) => sum + day.NightWorkTimeWithBreak);
            return $"{(int)days.TotalHours}:{days.Minutes}";
        }

    }
}
