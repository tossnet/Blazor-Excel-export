using Microsoft.JSInterop;
using BlazorExcelExport.Models;

namespace BlazorExcelExport.XLS;

public class Excel
{

    public async Task GenerateWeatherForecastAsync(IJSRuntime js, WeatherForecast[] data, string filename = "export.xlsx")
    {
        var weatherForecast = new WeatherForecastPDF();
        string XLSStream = weatherForecast.Edition(data);

        await js.InvokeVoidAsync("jsSaveAsFile", filename, XLSStream);
    }

}
