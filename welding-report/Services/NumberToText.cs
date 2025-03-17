using Microsoft.Extensions.Primitives;
using System.Globalization;
using welding_report.Models;
using Humanizer;


namespace welding_report.Services
{
    public interface INumberToText
    {
        void FillCostText(RequestReportData data);
        
    }
    public class NumberToText : INumberToText
    {

        public void FillCostText(RequestReportData data)
        {
            if (int.TryParse(data.Cost, NumberStyles.Any, CultureInfo.InvariantCulture, out var parsedValue))
            {
                data.CostText = NumberToWordsRu(parsedValue);
            }
        }

        private static string NumberToWordsRu(int number)
        {
            return number.ToWords(new CultureInfo("ru"));
        }

    }
}
