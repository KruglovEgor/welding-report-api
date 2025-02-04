using System.ComponentModel.DataAnnotations;

namespace welding_report.Models
{
    public class WeldingReportRequest
    {
        [Required]
        public string ReportNumber { get; set; }

        [Required]
        public List<WeldingJointRequest> Joints { get; set; } = new();

        public List<IFormFile> Photos { get; set; } = new();
    }

    public class WeldingJointRequest
    {
        [Required]
        public string EquipmentType { get; set; }

        [Required]
        public string PipelineNumber { get; set; }

        [Required]
        public string JointNumber { get; set; }

        [Required]
        [Range(1, 1000)]
        public double DiameterMm { get; set; }

        [Required]
        [Range(0.1, 100)]
        public double LengthMeters { get; set; }
    }
}
