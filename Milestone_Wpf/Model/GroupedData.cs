using System.Collections.Generic;

namespace Milestone_Wpf.Model
{
    public class GroupedData
    {
        public int Year { get; set; }
        public int Month { get; set; }
        public string EngineeringCenterCode { get; set; }
        public string GPID { get; set; }
        public List<MilestoneData> Data { get; set; }
    }
}
