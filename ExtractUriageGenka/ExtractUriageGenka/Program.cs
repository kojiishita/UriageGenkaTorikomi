// See https://aka.ms/new-console-template for more information
using ExtractUriageGenka;
using Microsoft.Extensions.Configuration;
using Microsoft.Office.Interop.Excel;
using PCM;


/// <summary>
/// シート名から数値型の月を得る
/// </summary>
int GetMonth(string sheetName)
{
    return sheetName switch
    {
        "７月" => 7,
        "８月" => 8,
        "９月" => 9,
        "１０月" => 10,
        "１１月" => 11,
        "１２月" => 12,
        "１月" => 1,
        "２月" => 2,
        "３月" => 3,
        "４月" => 4,
        "５月" => 5,
        "６月" => 6,
        "7月" => 7,
        "8月" => 8,
        "9月" => 9,
        "10月" => 10,
        "11月" => 11,
        "12月" => 12,
        "1月" => 1,
        "2月" => 2,
        "3月" => 3,
        "4月" => 4,
        "5月" => 5,
        "6月" => 6,
        _ => 0
    };
}


/// <summary>
/// Excelの出力範囲[A1形式]を取得する
/// </summary>
string GetRangeString(object[,] data, int startCol, int startRow)
{
    int row = data.GetLength(0);
    int col = Convert.ToInt32(data.Length / (double)row);
    return string.Format(
        "{0}{1}:{2}{3}",
        GetCellNameA1(startCol),
        startRow,
        GetCellNameA1(col + startCol - 1),
        (row + startRow - 1).ToString());
}

/// <summary>
/// 列番号からA1形式の列名を得る
/// </summary>
string GetCellNameA1(int c)
{
    var alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    var s = string.Empty;

    for (; c > 0; c = (c - 1) / 26)
    {
        int n = (c - 1) % 26;
        s = alpha.Substring(n, 1) + s;
    }
    return s;
}

/// <summary>
/// IEnumerable＜T＞を2次元配列[行・列]に変換する
/// </summary>
object[,] ToArray<T>(IEnumerable<T> enumerable)
{
    var props = typeof(T).GetProperties().ToList();
    object[,] data = new object[enumerable.Count(), props.Count()];
    int rowCount = 0;
    foreach (var obj in enumerable)
    {
        for (var colCount = 0; colCount <= props.Count - 1; colCount++)
        {
            var prop = props[colCount];
            data[rowCount, colCount] = prop.GetValue(obj);
        }

        rowCount++;
    }

    return data;
}

var configuration = new ConfigurationBuilder()
.SetBasePath(Directory.GetCurrentDirectory())
.AddJsonFile("appsettings.json", true, true)
.Build();

var targetFolder = configuration.GetSection("Folders")["TargetFolder"];
var outputFOlder = configuration.GetSection("Folders")["OutputFolder"];
var files = configuration.GetSection("Files").Get<List<string>>();

const int COL_BUNRUI = 1;
const int COL_SYUBETSU = 2;
const int COL_BUSHO = 3;
const int COL_KYAKUSAKIMEI = 4;
const int COL_KEIYAKU = 5;
const int COL_ANKENMEI = 6;
const int COL_JISSEKI = 10;
const int COL_JISSEKI_ZEIKOMI = 11;

var summaryList = new List<UriageGenkaSummary>();

foreach (var file in files)
{
    var shozoku = Path.GetFileNameWithoutExtension(file).Replace("売上原価表", string.Empty);
    var path = Path.Combine(targetFolder, file);
    if (File.Exists(path))
    {
        using var excelManager = new PcmExcelManager();
        var workBooks = excelManager.GetWorkbooks();
        var workBook = excelManager.OpenWorkbookReadOnly(workBooks, path);

        var workSheets = excelManager.GetSheets(workBook);
        foreach (Worksheet sheet in workSheets)
        {
            // シート名が月で終わるシート
            if (sheet.Name.EndsWith("月"))
            {
                // "社内費用"を含むセルを検索
                var endRange = excelManager.FindRange(sheet, "社内費用");
                if (endRange != null)
                {
                    // 最終行を取得
                    var endRow = endRange.Row;

                    // 年月
                    // 部署：システム一室、BC事業部他
                    // 客先名
                    // 契約
                    // 案件名
                    // 実績（税抜）
                    // 実績（税込）
                    var lastBunrui = string.Empty;
                    var lastShubetsu = string.Empty;
                    var lastBusho = string.Empty;

                    for (var row = 2; row < endRange.Row; row++)
                    {
                        var bunrui = excelManager.GetCell(sheet, row, COL_BUNRUI);
                        if (!string.IsNullOrEmpty(bunrui.Value))
                        {
                            lastBunrui = bunrui.Value;
                        }

                        var shubetsu = excelManager.GetCell(sheet, row, COL_SYUBETSU);
                        if (!string.IsNullOrEmpty(shubetsu.Value))
                        {
                            lastShubetsu = shubetsu.Value;
                        }

                        var busho = excelManager.GetCell(sheet, row, COL_BUSHO);
                        if (!string.IsNullOrEmpty(busho.Value))
                        {
                            lastBusho = busho.Value;
                        }

                        // 客先名が空以外＆売上から始まらない＆計で終わらない
                        var kyakusakimei = excelManager.GetCell(sheet, row, COL_KYAKUSAKIMEI);
                        if (!string.IsNullOrEmpty(kyakusakimei.Value)
                            && !((string)kyakusakimei.Value).StartsWith("売上")
                            && !((string)kyakusakimei.Value).EndsWith("計"))
                        {
                            var keiyaku = excelManager.GetCell(sheet, row, COL_KEIYAKU);
                            var ankenmei = excelManager.GetCell(sheet, row, COL_ANKENMEI);
                            var jisseki = excelManager.GetCell(sheet, row, COL_JISSEKI);
                            var jissekizei = excelManager.GetCell(sheet, row, COL_JISSEKI_ZEIKOMI);

                            summaryList.Add(new UriageGenkaSummary()
                            {
                                Year = 0,
                                Month = GetMonth(sheet.Name),
                                Bunrui = lastBunrui,
                                Busho = $"{shozoku}_{lastBusho}",
                                Kyakusakimei = kyakusakimei.Value,
                                Keiyaku = keiyaku.Value,
                                Ankenmei = ankenmei.Value,
                                Jisseki = (decimal?)jisseki.Value2 ?? 0m,
                                JissekZeikomi = (decimal?)jissekizei.Value2 ?? 0m,
                            });

                            Console.WriteLine($"{lastBunrui},{lastShubetsu},{lastBusho},{keiyaku.Value},{ankenmei.Value},{jisseki.Value2}");
                        }
                    }
                }
            }
        }

        workBook.Close();
    }
}

using (var excelManager = new PcmExcelManager())
{
    var workBooks = excelManager.GetWorkbooks();
    var workBook = workBooks.Add();

    var uriage = summaryList.Where(e => e.Bunrui == "売上");
    var shiire = summaryList.Where(e => e.Bunrui == "仕入");

    // 売上シート準備
    var uriageSheet = workBook.Sheets["Sheet1"];
    uriageSheet.Name = "売上";

    // 列ヘッダ設定
    uriageSheet.Cells[1, 1].Value = "年";
    uriageSheet.Cells[1, 2].Value = "月";
    uriageSheet.Cells[1, 3].Value = "分類";
    uriageSheet.Cells[1, 4].Value = "部署";
    uriageSheet.Cells[1, 5].Value = "客先名";
    uriageSheet.Cells[1, 6].Value = "契約";
    uriageSheet.Cells[1, 7].Value = "案件名";
    uriageSheet.Cells[1, 8].Value = "実績(税抜)";
    uriageSheet.Cells[1, 9].Value = "実績(税込)";

    // 2次元配列にして一括設定
    object[,] arrayValues = ToArray(uriage);
    excelManager.Write(uriageSheet, GetRangeString(arrayValues, 1, 2), arrayValues);

    // 仕入シート準備
    var totalCount = workBook.Sheets.Count;
    workBook.Worksheets.Add(Type.Missing, After: workBook.Worksheets[totalCount]);
    totalCount = workBook.Sheets.Count;
    var lastSheetName = workBook.Worksheets[totalCount].Name;
    workBook.Worksheets[lastSheetName].Name = "仕入";
    var shiireSheet = workBook.Sheets["仕入"];

    // 列ヘッダ設定
    shiireSheet.Cells[1, 1].Value = "年";
    shiireSheet.Cells[1, 2].Value = "月";
    shiireSheet.Cells[1, 3].Value = "分類";
    shiireSheet.Cells[1, 4].Value = "部署";
    shiireSheet.Cells[1, 5].Value = "客先名";
    shiireSheet.Cells[1, 6].Value = "契約";
    shiireSheet.Cells[1, 7].Value = "案件名";
    shiireSheet.Cells[1, 8].Value = "実績(税抜)";
    shiireSheet.Cells[1, 9].Value = "実績(税込)";

    // 2次元配列にして一括設定
    arrayValues = ToArray(shiire);
    excelManager.Write(shiireSheet, GetRangeString(arrayValues, 1, 2), arrayValues);

    // 保存
    workBook.SaveAs(Path.Combine(targetFolder, $"売上原価表集計_{DateTime.Now:yyyyMMddHHmmss}.xlsx"));
}