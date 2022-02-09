using ClosedXML.Excel;
using BlazorExcelExport.Models;

namespace BlazorExcelExport.XLS;

public class WeatherForecastXLS
{
    public byte[] Edition(WeatherForecast[] data)
    {
        var wb = new XLWorkbook();
        wb.Properties.Author = "the Author";
        wb.Properties.Title = "the Title";
        wb.Properties.Subject = "the Subject";
        wb.Properties.Category = "the Category";
        wb.Properties.Keywords = "the Keywords";
        wb.Properties.Comments = "the Comments";
        wb.Properties.Status = "the Status";
        wb.Properties.LastModifiedBy = "the Last Modified By";
        wb.Properties.Company = "the Company";
        wb.Properties.Manager = "the Manager";

        var ws = wb.Worksheets.Add("Weather Forecast");

        ws.Cell(1, 1).Value = "Date";
        ws.Cell(1, 2).Value = "Temp. (C)";
        ws.Cell(1, 3).Value = "Temp. (F)";
        ws.Cell(1, 4).Value = "Summary";

        // Fill a cell with a date
        //var cellDateTime = ws.Cell(1, 2);
        //cellDateTime.Value = new DateTime(2010, 9, 2);
        //cellDateTime.Style.DateFormat.Format = "yyyy-MMM-dd";

        for (int row = 0; row < data.Length; row++)
        {
            // The apostrophe is to force ClosedXML to treat the date as a string
            ws.Cell(row + 1, 1).Value = "'" + data[row].Date.ToShortDateString();
            ws.Cell(row + 1, 2).Value = data[row].TemperatureC;
            ws.Cell(row + 1, 3).Value = data[row].TemperatureF;
            ws.Cell(row + 1, 4).Value = data[row].Summary;
        }

        MemoryStream XLSStream = new();
        wb.SaveAs(XLSStream);

        return XLSStream.ToArray();
    }
}
