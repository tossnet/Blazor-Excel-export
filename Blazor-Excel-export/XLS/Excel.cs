using Microsoft.JSInterop;
using BlazorExcelExport.Models;

namespace BlazorExcelExport.XLS;

public class Excel
{
    public async Task GenerateWeatherForecastAsync(IJSRuntime js, WeatherForecast[] data, string filename = "export.xlsx")
    {
        var weatherForecast = new WeatherForecastXLS();
        string XLSStream = weatherForecast.Edition(data);

        await js.InvokeVoidAsync("jsSaveAsFile", filename, XLSStream);
    }

    public async Task TemplateWeatherForecastAsync(IJSRuntime js, Stream streamTemplate, WeatherForecast[] data, string filename = "export.xlsx")
    {
        var templateXLS = new UseTemplateXLS();
        string XLSStream = templateXLS.Edition(streamTemplate,data);

        await js.InvokeVoidAsync("jsSaveAsFile", filename, XLSStream);
    }

}
