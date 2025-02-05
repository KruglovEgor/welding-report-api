using System.ComponentModel.DataAnnotations;
using Microsoft.OpenApi.Any;
using Microsoft.OpenApi.Models;
using Swashbuckle.AspNetCore.Annotations;
using Swashbuckle.AspNetCore.SwaggerGen;

namespace welding_report.Models
{
    public class WeldingReportRequest
    {
        [Required]
        public string ReportNumber { get; set; }

        [Required]
        [SwaggerSchema(Description = "Массив стыков в формате JSON")]
        public List<WeldingJoint> Joints { get; set; } = new();

        public List<IFormFile> Photos { get; set; } = new();
    }

    public class WeldingJoint
    {
        [Required]
        public string EquipmentType { get; set; }

        [Required]
        public string PipelineNumber { get; set; }

        [Required]
        public string CompanyName { get; set; }

        [Required]
        public string JointNumber { get; set; }

        [Required]
        [Range(1, 1000)]
        public double DiameterMm { get; set; }

        [Range(0.1, 100)]
        public double LengthMeters { get; set; }
    }

    // Кастомный фильтр для обработки form-data
    public class FormDataOperationFilter : IOperationFilter
    {
        public void Apply(OpenApiOperation operation, OperationFilterContext context)
        {
            if (context.ApiDescription.HttpMethod == "POST" &&
                context.ApiDescription.RelativePath == "api/WeldingReport/generate")
            {
                operation.RequestBody = new OpenApiRequestBody
                {
                    Content =
                {
                    ["multipart/form-data"] = new OpenApiMediaType
                    {
                        Schema = new OpenApiSchema
                        {
                            Type = "object",
                            Properties =
                            {
                                ["ReportNumber"] = new OpenApiSchema { Type = "string" },
                                ["Joints"] = new OpenApiSchema
                                {
                                    Type = "string",
                                    Description = "JSON array of joints",
                                    Example = new OpenApiString("[" +
                                    "\n{\"EquipmentType\": \"Трубопровод\"," +
                                    "\n\"PipelineNumber\":\"TP-01\"," +
                                    "\n\"CompanyName\":\"ООО СваркаМонтаж\"," +
                                    "\n\"JointNumber\":\"123\"," +
                                    "\n\"DiameterMm\":150.5," +
                                    "\n\"LengthMeters\":2.75" +
                                    "\n}" +
                                    "\n]")
                                },
                                ["Photos"] = new OpenApiSchema
                                {
                                    Type = "array",
                                    Items = new OpenApiSchema { Type = "string", Format = "binary" }
                                }
                            }
                        }
                    }
                }
                };
            }
        }
    }
}
