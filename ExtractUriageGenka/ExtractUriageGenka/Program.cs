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
    var props = 
        typeof(T).GetProperties().OrderBy(p => (Attribute.GetCustomAttribute(p, typeof(OrderAttribute)) as OrderAttribute).Value).ToList();
    //var props = typeof(T).GetProperties().ToList();
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

var columnHeaders = new List<ColumnHeader>() 
{ 
    new ColumnHeader() { Name = "年", Column = 1 },
    new ColumnHeader() { Name = "月", Column = 2 },
    new ColumnHeader() { Name = "分類", Column = 3 },
    new ColumnHeader() { Name = "部署", Column = 4 },
    new ColumnHeader() { Name = "客先名", Column = 5 },
    new ColumnHeader() { Name = "契約", Column = 6 },
    new ColumnHeader() { Name = "案件名", Column = 7 },
    new ColumnHeader() { Name = "実績(税抜)", Column = 8 },
    new ColumnHeader() { Name = "実績(税込)", Column = 9 },
};

var targetFolder = configuration.GetSection("Folders")["TargetFolder"];
var outputFOlder = configuration.GetSection("Folders")["OutputFolder"];
var files = configuration.GetSection("Files").Get<List<string>>();
var filters = configuration.GetSection("Filters").Get<List<ExtractCondition>>();

if (filters != null)
{
    foreach (var filter in filters)
    {
        var fieldName = filter.FieldName;
        if (!columnHeaders.Select(e => e.Name).Contains(fieldName))
        {
            throw new ArgumentException("$フィルターに設定されている[{fieldName}]は存在しない列名です");
        }

        var ope = filter.Operator;
        if (!new string[] { "=", "!=", "<", ">", "<=", "=>" }.Contains(ope))
        {
            throw new ArgumentException("$フィルターに設定されている比較演算子[{ope}]は無効です");
        }
    }
}

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

    // 売上側にフィルタをかける
    var exp = new List<Tuple<string, string, object, Type>>();
    exp.Add(new Tuple<string, string, object, Type>("Bunrui", "=", "売上", typeof(string)));
    if (filters != null)
    {
        foreach (var filter in filters)
        {
            (string field, string condition, object value, Type type) f = filter.FieldName switch
            {
                "年" => ("Year", filter.Operator, int.Parse(filter.Value), typeof(int)),
                "月" => ("Month", filter.Operator, int.Parse(filter.Value), typeof(int)),
                "分類" => ("Bunrui", filter.Operator, filter.Value, typeof(string)),
                "部署" => ("Busho", filter.Operator, filter.Value, typeof(string)),
                "客先名" => ("Kyakusakimei", filter.Operator, filter.Value, typeof(string)),
                "契約" => ("Keiyaku", filter.Operator, filter.Value, typeof(string)),
                "案件名" => ("Ankenmei", filter.Operator, filter.Value, typeof(string)),
                "実績(税抜)" => ("Jisseki", filter.Operator, decimal.Parse(filter.Value), typeof(decimal)),
                "実績(税込)" => ("JissekZeikomi", filter.Operator, decimal.Parse(filter.Value), typeof(decimal)),
                _ => (string.Empty, filter.Operator, filter.Value, typeof(string))
            };

            exp.Add(new Tuple<string, string, object, Type>(f.field, f.condition, f.value, f.type));
        }
    }
    var uriage = summaryList.DynamicWhere(exp).ToList();

    // 仕入側にフィルタをかける
    exp.Clear();
    exp.Add(new Tuple<string, string, object, Type>("Bunrui", "=", "仕入", typeof(string)));
    if (filters != null)
    {
        foreach (var filter in filters)
        {
            (string field, string condition, object value, Type type) f = filter.FieldName switch
            {
                "年" => ("Year", filter.Operator, int.Parse(filter.Value), typeof(int)),
                "月" => ("Month", filter.Operator, int.Parse(filter.Value), typeof(int)),
                "分類" => ("Bunrui", filter.Operator, filter.Value, typeof(string)),
                "部署" => ("Busho", filter.Operator, filter.Value, typeof(string)),
                "客先名" => ("Kyakusakimei", filter.Operator, filter.Value, typeof(string)),
                "契約" => ("Keiyaku", filter.Operator, filter.Value, typeof(string)),
                "案件名" => ("Ankenmei", filter.Operator, filter.Value, typeof(string)),
                "実績(税抜)" => ("Jisseki", filter.Operator, decimal.Parse(filter.Value), typeof(decimal)),
                "実績(税込)" => ("JissekZeikomi", filter.Operator, decimal.Parse(filter.Value), typeof(decimal)),
                _ => (string.Empty, filter.Operator, filter.Value, typeof(string))
            };

            exp.Add(new Tuple<string, string, object, Type>(f.field, f.condition, f.value, f.type));
        }
    }
    var shiire = summaryList.DynamicWhere(exp).ToList();

    // 売上シート準備
    var uriageSheet = workBook.Sheets["Sheet1"];
    uriageSheet.Name = "売上";

    // 列ヘッダ設定
    foreach (var ch in columnHeaders)
    {
        uriageSheet.Cells[1, ch.Column].Value = ch.Name;
    }

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
    foreach (var ch in columnHeaders)
    {
        shiireSheet.Cells[1, ch.Column].Value = ch.Name;
    }

    // 2次元配列にして一括設定
    arrayValues = ToArray(shiire);
    excelManager.Write(shiireSheet, GetRangeString(arrayValues, 1, 2), arrayValues);

    // 保存
    workBook.SaveAs(Path.Combine(targetFolder, $"売上原価表集計_{DateTime.Now:yyyyMMddHHmmss}.xlsx"));
}