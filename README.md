# Blazor Excel export

DEMO : https://tossnet.github.io/Blazor-Excel-export/

Article on my blog : https://www.peug.net/en/blazor-create-or-export-your-data-to-excel/ 

![blazor excel export](https://user-images.githubusercontent.com/3845786/152016618-1aad643c-649a-41fb-afaa-8713023734df.png)



### Method 1 : Creation of a file in XLS format

For these two methods, I’m going to use a Nugets package named CloseXML : ttps://github.com/ClosedXML  which is licensed by MIT. So you can use it freely even if your application is commercial.

The principle is quite simple if you want to create a basic excel file: You create the excel file “XLWorkbook” in which you will add one (or more 🙂 ) tab “Worksheets” and you will move from cell to cell “Cell” to place your data and at the end you save your file by the function “SaveAs”. Simple, isn’t it?

```
var wb = new XLWorkbook();
wb.Properties.Author = "the Author";
wb.Properties.Title = "the Title";
wb.Properties.Subject = "the Subject";
  
var ws = wb.Worksheets.Add("Weather Forecast");
  
ws.Cell(1, 1).Value = "Temp. (C)";
ws.Cell(1, 2).Value = "Temp. (F)";
ws.Cell(1, 3).Value = "Summary";
  
for (int row = 0; row < data.Length; row++)
{
     ws.Cell(row + 1, 1).Value = data[row].TemperatureC;
     ws.Cell(row + 1, 2).Value = data[row].TemperatureF;
     ws.Cell(row + 1, 3).Value = data[row].Summary;
}
  
MemoryStream XLSStream = new();
wb.SaveAs(XLSStream);
```

Note that here I save my Excel spreadsheet in a Stream but we can directly write the path and the file name. Being on a web application, I will then use a javacript function to propose to my user to download his report.

### Method 2 – Use an existing excel file

For this method, we will use the Nugets CloseXML.Report package https://github.com/ClosedXML/ClosedXML.Report . You can easily find these packages in the Nuguets explorer of Visual Studio.

By using an Excel file as a Template, we can quickly make a nice presentation. It all depends on what you want. Honestly, if it’s just a data export for “non-geek” users, the first method will do the job well. Then yes, you could propose to export in raw formats like JSON, CSV etc… but depending on the audience using your application, it’s not nice.

On the code side, it’s much simpler:
```
var template = new  XLTemplate(streamTemplate);
  
template.AddVariable("WeatherForecasts", data);
template.Generate();
  
MemoryStream XLSStream = new();
```
emplate.SaveAs(XLSStream);

That’s it. We give XLTemplate the initial XLS file as a parameter or, as in this case, the Stream of my file. Then we pass it via “AddVariable” a List, Array IEnumerable… and indicate the name of the group of cells! This is where I had difficulties to master the beast. But once found, it’s like anything else; it’s easy 🙂

![name-manager-excel-1](https://user-images.githubusercontent.com/3845786/154007896-01e73e6b-1c47-4be0-acf8-b37a9e08809c.png)

And here you find the name I gave to my list: “WeatherForecasts“. Note that in the ” double braces you have item followed by the name of my properties. item is imposed by the library. There are 3 of them:

* item – element of the list.
* index – index of an item
* items – the whole set ex: TOTAL RECORDS : {{items.Count()}}

Naming a group is only useful for lists, otherwise it is enough to put a name like {{name}} in a cell and in your C# code to add :

```
template.AddVariable("name", "christophe");
```

There is quite a lot of documentation on these libraries: https://closedxml.github.io/ClosedXML.Report/docs/en/index


https://user-images.githubusercontent.com/3845786/152327358-ec77cd70-3a09-4e10-a66d-9886266119ce.mp4
