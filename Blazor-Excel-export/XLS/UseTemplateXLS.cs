using BlazorExcelExport.Models;
using ClosedXML.Report;

namespace BlazorExcelExport.XLS;

public class UseTemplateXLS
{
    public string Edition(Stream streamTemplate, WeatherForecast[] data)
    {
        var template = new  XLTemplate(streamTemplate);

        template.AddVariable("WeatherForecasts", data);
        template.Generate();

        MemoryStream XLSStream = new();
        template.SaveAs(XLSStream);


        return Convert.ToBase64String(XLSStream.ToArray());
    }

}
