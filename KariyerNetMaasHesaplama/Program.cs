using HtmlAgilityPack;
using SpreadCheetah;
using SpreadCheetah.Styling;
using SpreadCheetah.Worksheets;
using System.Globalization;
using System.Web;

var outputFile = "Maaslar.xlsx";
if (File.Exists(outputFile))
    File.Delete(outputFile);

var document = new HtmlWeb().Load("https://www.kariyer.net/pozisyonlar");

using var stream = File.Create(outputFile);
using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
var worksheetOptions = new WorksheetOptions();
worksheetOptions.Column(1).Width = 50;
worksheetOptions.Column(2).Width = 18;
worksheetOptions.Column(3).Width = 18;
worksheetOptions.Column(4).Width = 18;
await spreadsheet.StartWorksheetAsync("Maaşlar", worksheetOptions);

var headerStyle = new Style();
headerStyle.Font.Bold = true;
var headerStyleId = spreadsheet.AddStyle(headerStyle);

var cells = new[]
{
    new Cell("POZİSYON", headerStyleId),
    new Cell("ORTALAMA MAAŞ", headerStyleId),
    new Cell("EN DÜŞÜK MAAŞ", headerStyleId),
    new Cell("EN YÜKSEK MAAŞ", headerStyleId)
};
await spreadsheet.AddRowAsync(cells);

var positionNodes = document.DocumentNode.SelectNodes("//div[@class='pgl-positions']//a");
for (int i = 0; i < positionNodes.Count; i++)
{
    var positionNode = positionNodes[i];
    document = new HtmlWeb().Load($"https://www.kariyer.net{positionNode.GetAttributeValue("href", null)}");

    try
    {
        var position = HttpUtility.HtmlDecode(document.DocumentNode.SelectSingleNode("//h1").InnerText);
        var averageSalary = document.DocumentNode.SelectSingleNode("//div[@class='pg-salary-box']/div[@class='pg-salary-box-value']")?.InnerText;

        if (averageSalary == null)
            continue;

        var minimumSalary = document.DocumentNode.SelectSingleNode("//div[@class='pg-salary-box double']/div[@class='left']/div[@class='pg-salary-box-value']").InnerText;
        var maximumSalary = document.DocumentNode.SelectSingleNode("//div[@class='pg-salary-box double']/div[@class='right']/div[@class='pg-salary-box-value']").InnerText;

        Console.WriteLine($"POZİSYON: {position}");
        Console.WriteLine($"ORTALAMA MAAŞ: {averageSalary}");
        Console.WriteLine($"EN DÜŞÜK MAAŞ: {minimumSalary}");
        Console.WriteLine($"EN YÜKSEK MAAŞ: {maximumSalary}");

        cells = new[]
        {
            new Cell(position),
            new Cell(Convert.ToDecimal(averageSalary.Replace(".", ""), CultureInfo.InvariantCulture)),
            new Cell(Convert.ToDecimal(minimumSalary.Replace(".", ""), CultureInfo.InvariantCulture)),
            new Cell(Convert.ToDecimal(maximumSalary.Replace(".", ""), CultureInfo.InvariantCulture))
        };
        await spreadsheet.AddRowAsync(cells);
    }
    catch (Exception e)
    {
        Console.WriteLine(e.Message);
        if (e.InnerException != null)
            Console.WriteLine(e.InnerException.Message);
    }
}
await spreadsheet.FinishAsync();
