using LinqToExcel.Attributes;

namespace Milestone_Wpf.Model
{
    public class RawData
    {
        [ExcelColumn("Year")]
        public int Year { get; set; }

        [ExcelColumn("Month Number")]
        public int Month { get; set; }

        [ExcelColumn("Engineering Center Code")]
        public string EngineeringCenterCode { get; set; }

        [ExcelColumn("GPID")]
        public string GPID { get; set; }

        [ExcelColumn("GPID Description")]
        public string GPIDDescription { get; set; }

        [ExcelColumn("RTO")]
        public double RTO { get; set; }

        [ExcelColumn("Workload mth GPID")]
        public double WorkloadmthGPID { get; set; }

        [ExcelColumn("Milestone")]
        public string Milestone { get; set; }

        [ExcelColumn("GPID SubGroup")]
        public string GPID_SubGroup { get; set; }
    }
}
