
namespace PCM
{
    using System;
    using System.Diagnostics;
    using System.IO;
    using System.Linq;
    using System.Runtime.InteropServices;

    using Microsoft.Office.Interop.Excel;
    using Microsoft.VisualBasic;

    /// <summary>
    /// Excel <see cref="Application"/> 管理クラスです。
    /// </summary>
    public class PcmExcelManager : PcmAbstractComStackDisposableBase
    {
        /// <summary>
        /// Excel <see cref="Application"/> オブジェクトです。
        /// </summary>
        private Application excelApp;

        /// <summary>
        /// エクセルファイル名 (拡張子なし)です。
        /// </summary>
        private string excelFileNameWithoutExtension;

        /// <summary>
        /// 引数を指定せずに、 <see cref="ExcelManager"/> クラスの新しいインスタンスを初期化します。
        /// </summary>
        /// <remarks>
        /// 初期設定を行います。
        /// <para>
        /// [Excel の初期設定]
        /// <para>・<see cref="_Application.DisplayAlerts"/>：false</para>
        /// <para>・<see cref="_Application.ScreenUpdating"/>：false</para>
        /// <para>・<see cref="_Application.Visible"/>：false</para>
        /// </para>
        /// </remarks>
        public PcmExcelManager()
        {
            this.excelApp = new Application
            {
                DisplayAlerts = false,
                ScreenUpdating = false,
                Visible = false
            };
        }

        /// <summary>
        /// Excel <see cref="Application"/> オブジェクトです。
        /// </summary>
        public Application ExcelApplication
        {
            get { return this.excelApp; }
        }

        /// <summary>
        /// リソースを解放します。
        /// </summary>
        public override void Dispose()
        {
            this.QuitExcel();
            base.Dispose();
        }

        /// <summary>
        /// Excelを終了します。
        /// </summary>
        /// <remarks>
        /// <see cref="AbstractComStackDisposableBase.Release"/> を使用します。
        /// <see cref="excelApp"/> が非表示の場合、 <see cref="_Application.Quit"/> を実行します。
        /// <see cref="Marshal.ReleaseComObject(object)"/> を使用し <see cref="excelApp"/> を解放します。
        /// </remarks>
        public void QuitExcel()
        {
            this.Release();

            if (!this.excelApp.Visible)
            {
                this.excelApp.Quit();
            }

            Marshal.ReleaseComObject(this.excelApp);
            this.excelApp = null;
        }

        /// <summary>
        /// DisplayAlertsを[True]に設定します。
        /// </summary>
        public void DisplayAlertsOn()
        {
            this.excelApp.DisplayAlerts = true;
        }

        /// <summary>
        /// <see cref="Workbooks"/> を取得します。
        /// </summary>
        /// <returns><see cref="Workbooks"/> オブジェクト。</returns>
        public Workbooks GetWorkbooks()
        {
            Workbooks b = this.excelApp.Workbooks;
            this.Stack.Push(b);
            return b;
        }

        /// <summary>
        /// 指定された <see cref="Workbook"/> を開きます。
        /// </summary>
        /// <param name="books">対象の <see cref="Workbooks"/> 。</param>
        /// <param name="fileName">開くファイル名。</param>
        /// <returns>開いた <see cref="Workbook"/> 。</returns>
        public Workbook OpenWorkbook(Workbooks books, string fileName)
        {
            Workbook b;
            var fileExtension = Path.GetExtension(fileName);
            if (fileExtension == ".xlsx" || fileExtension == ".xls")
            {
                b = books.Add(fileName);
            }
            else
            {
                b = books.Open(fileName);
            }
            
            this.excelFileNameWithoutExtension = Path.GetFileNameWithoutExtension(fileName);

            this.Stack.Push(b);
            return b;
        }

        /// <summary>
        /// 指定された <see cref="Workbook"/> を読み取り専用で開きます。
        /// </summary>
        /// <param name="books">対象の <see cref="Workbooks"/> 。</param>
        /// <param name="fileName">開くファイル名。</param>
        /// <returns>開いた <see cref="Workbook"/> 。</returns>
        public Workbook OpenWorkbookReadOnly(Workbooks books, string fileName)
        {
            var fileExtension = Path.GetExtension(fileName);
            bool editable = false;
            if (fileExtension == ".xltx")
            {
                editable = true;
            }

            var b = books.Open(fileName, ReadOnly: true, Editable: editable);
            this.excelFileNameWithoutExtension = Path.GetFileNameWithoutExtension(fileName);

            this.Stack.Push(b);
            return b;
        }

        /// <summary>
        /// 指定された <see cref="Workbook"/> の <see cref="Sheets"/> オブジェクトを取得します。
        /// </summary>
        /// <param name="book"><see cref="Workbook"/> オブジェクト。</param>
        /// <returns><see cref="Sheets"/> オブジェクト。</returns>
        public Sheets GetSheets(Workbook book)
        {
            Sheets s = book.Worksheets;
            this.Stack.Push(s);
            return s;
        }

        /// <summary>
        /// 指定されたインデックスの <see cref="Worksheet"/> オブジェクトを取得します。
        /// </summary>
        /// <param name="sheets"><see cref="Sheets"/> オブジェクト。</param>
        /// <param name="index">インデックス。</param>
        /// <returns><see cref="Worksheet"/> オブジェクト。</returns>
        public Worksheet GetWorksheet(Sheets sheets, int index)
        {
            Worksheet s = sheets.Item[index] as Worksheet;
            this.Stack.Push(s);
            return s;
        }

        /// <summary>
        /// <see cref="ExcelApplication"/> のアクティブ <see cref="Window"/> を取得します。
        /// </summary>
        /// <returns><see cref="Window"/> オブジェクト。</returns>
        public Window GetActiveWindow()
        {
            Window w = this.excelApp.ActiveWindow;
            this.Stack.Push(w);
            return w;
        }

        /// <summary>
        /// <see cref="Worksheet"/> の <see cref="PageSetup"/> を取得します。
        /// </summary>
        /// <param name="worksheet"><see cref="Worksheet"/> オブジェクト。</param>
        /// <returns><see cref="PageSetup"/> オブジェクト。</returns>
        public PageSetup GetPageSetup(Worksheet worksheet)
        {
            PageSetup ps = worksheet.PageSetup;
            this.Stack.Push(ps);
            return ps;
        }

        /// <summary>
        /// <see cref="Worksheet"/> の <see cref="HPageBreaks"/> を取得します。
        /// </summary>
        /// <param name="worksheet"><see cref="Worksheet"/> オブジェクト。</param>
        /// <returns><see cref="HPageBreaks"/> オブジェクト。</returns>
        public HPageBreaks GetHPageBreaks(Worksheet worksheet)
        {
            HPageBreaks hpb = worksheet.HPageBreaks;
            this.Stack.Push(hpb);
            return hpb;
        }

        /// <summary>
        /// <see cref="Worksheet"/> の <see cref="Shapes"/> を取得します。
        /// </summary>
        /// <param name="worksheet"><see cref="Worksheet"/> オブジェクト。</param>
        /// <returns><see cref="Shapes"/> オブジェクト。</returns>
        public Shapes GetShapes(Worksheet worksheet)
        {
            Shapes s = worksheet.Shapes;
            this.Stack.Push(s);
            return s;
        }

        /// <summary>
        /// <see cref="Worksheet"/> 全セルの <see cref="Range"/> を取得します。
        /// </summary>
        /// <param name="worksheet"><see cref="Worksheet"/> オブジェクト。</param>
        /// <returns><see cref="Range"/> オブジェクト。</returns>
        public Range GetCells(Worksheet worksheet)
        {
            Range r = worksheet.Cells;
            this.Stack.Push(r);
            return r;
        }

        /// <summary>
        /// 指定された範囲を <see cref="Range"/> オブジェクトとして取得します。
        /// </summary>
        /// <param name="worksheet"><see cref="Worksheet"/> オブジェクト。</param>
        /// <param name="range">対象範囲。</param>
        /// <returns><see cref="Range"/> オブジェクト。</returns>
        public Range GetRange(Worksheet worksheet, string range)
        {
            Range r = worksheet.Range[range];
            this.Stack.Push(r);
            return r;
        }

        /// <summary>
        /// 指定された範囲を <see cref="Range"/> オブジェクトとして取得します。
        /// </summary>
        /// <param name="worksheet"><see cref="Worksheet"/> オブジェクト。</param>
        /// <param name="range1">範囲開始 <see cref="Range"/> 。</param>
        /// <param name="range2">範囲終了 <see cref="Range"/> 。</param>
        /// <returns><see cref="Range"/> オブジェクト。</returns>
        public Range GetRange(Worksheet worksheet, Range range1, Range range2)
        {
            Range r = worksheet.Range[range1, range2];
            this.Stack.Push(r);
            return r;
        }

        /// <summary>
        /// 指定された範囲を <see cref="Range"/> オブジェクトとして取得します。
        /// </summary>
        /// <param name="worksheet"><see cref="Worksheet"/> オブジェクト。</param>
        /// <param name="startRow">開始行</param>
        /// <param name="endRow">終了行</param>
        /// <returns><see cref="Range"/> オブジェクト。</returns>
        public Range GetRange(Worksheet worksheet, int startRow, int endRow)
        {
            string range = string.Format("{0}:{1}", startRow, endRow);
            Range r = worksheet.Range[range];
            this.Stack.Push(r);
            return r;
        }

        /// <summary>
        /// 指定された行一覧を <see cref="Range"/> オブジェクトとして取得します。
        /// </summary>
        /// <param name="worksheet"><see cref="Worksheet"/> オブジェクト。</param>
        /// <param name="range">対象範囲。</param>
        /// <returns><see cref="Range.Rows"/> が格納された <see cref="Range"/> オブジェクト。</returns>
        public Range GetRows(Worksheet worksheet, string range)
        {
            var rows = this.GetRange(worksheet, range).Rows;
            this.Stack.Push(rows);
            return rows;
        }

        /// <summary>
        /// 指定された行を <see cref="Range"/> オブジェクトとして取得します。
        /// </summary>
        /// <param name="worksheet"><see cref="Worksheet"/> オブジェクト。</param>
        /// <param name="rowIndex">行番号。</param>
        /// <returns><see cref="Range"/> オブジェクト。</returns>
        public Range GetRow(Worksheet worksheet, int rowIndex)
        {
            Range r = worksheet.Rows[rowIndex] as Range;
            this.Stack.Push(r);
            return r;
        }

        /// <summary>
        /// 指定された列を <see cref="Range"/> オブジェクトとして取得します。
        /// </summary>
        /// <param name="worksheet"><see cref="Worksheet"/> オブジェクト。</param>
        /// <param name="startCol">開始列</param>
        /// <param name="endCol">終了列</param>
        /// <returns><see cref="Range"/> オブジェクト。</returns>
        public Range GetColumns(Worksheet worksheet, int startCol, int endCol)
        {
            Range rs = worksheet.Columns[startCol] as Range;
            Range re = worksheet.Columns[endCol] as Range;
            this.Stack.Push(rs);
            this.Stack.Push(re);

            Range r = worksheet.Range[rs, re];
            this.Stack.Push(r);
            return r;
        }

        /// <summary>
        /// 指定されたセル を <see cref="Range"/> オブジェクトとして取得します。
        /// </summary>
        /// <param name="worksheet"><see cref="Worksheet"/> オブジェクト。</param>
        /// <param name="rowIndex">行インデックス。</param>
        /// <param name="colIndex">列インデックス。</param>
        /// <returns>該当セルを表す <see cref="Range"/> オブジェクト。</returns>
        public Range GetCell(Worksheet worksheet, int rowIndex, int colIndex)
        {
            Range r = worksheet.Cells;
            Range v = r[rowIndex, colIndex] as Range;
            this.Stack.Push(r);
            this.Stack.Push(v);
            return v;
        }

        /// <summary>
        /// 指定されたセルを <see cref="Range"/> オブジェクトとして取得します。
        /// </summary>
        /// <param name="worksheet"><see cref="Worksheet"/> オブジェクト。</param>
        /// <param name="coordinate">座標。</param>
        /// <returns>該当セルを表す <see cref="Range"/> オブジェクト。</returns>
        public Range GetCell(Worksheet worksheet, string coordinate)
        {
            Range v = worksheet.Range[coordinate];
            this.Stack.Push(v);
            return v;
        }

        /// <summary>
        /// 指定された <see cref="Worksheet"/> の中からキーフレーズを含む先頭の <see cref="Range"/> を取得します。
        /// </summary>
        /// <param name="worksheet">検索する <see cref="Worksheet"/> 。</param>
        /// <param name="keyPhrase">検索するキーフレーズ。</param>
        /// <returns>該当セルを表す <see cref="Range"/> オブジェクト。</returns>
        public Range FindRange(Worksheet worksheet, string keyPhrase)
        {
            var rs = worksheet.Cells;
            var r = rs.Find(keyPhrase, LookIn: XlFindLookIn.xlValues);
            this.Stack.Push(rs);
            if (r != null)
            {
                this.Stack.Push(r);
            }

            return r;
        }

        /// <summary>
        /// 指定された <see cref="Range"/> の中からキーフレーズを含む先頭の <see cref="Range"/> を取得します。
        /// </summary>
        /// <param name="range">検索範囲。</param>
        /// <param name="keyPhrase">検索するキーフレーズ。</param>
        /// <returns>該当セルを表す <see cref="Range"/> オブジェクト。</returns>
        public Range FindRange(Range range, string keyPhrase)
        {
            var r = range.Find(keyPhrase, LookIn: XlFindLookIn.xlValues, LookAt: XlLookAt.xlWhole);
            this.Stack.Push(r);
            return r;
        }

        /// <summary>
        /// 指定された <see cref="Range"/> の <see cref="Range.EntireColumn"/> を取得します。
        /// </summary>
        /// <param name="range"><see cref="Range"/> オブジェクト。</param>
        /// <returns><see cref="Range.EntireColumn"/> が設定された <see cref="Range"/> オブジェクト。</returns>
        public Range GetEntireColumn(Range range)
        {
            var c = range.EntireColumn;
            this.Stack.Push(c);
            return c;
        }

        /// <summary>
        /// 指定された <see cref="Range"/> の <see cref="Range.EntireRow"/> を取得します。
        /// </summary>
        /// <param name="range"><see cref="Range"/> オブジェクト。</param>
        /// <returns><see cref="Range.EntireColumn"/> が設定された <see cref="Range"/> オブジェクト。</returns>
        public Range GetEntireRow(Range range)
        {
            var r = range.EntireRow;
            this.Stack.Push(r);
            return r;
        }

        /// <summary>
        /// Excelを開きます。
        /// </summary>
        public void Show()
        {
            this.excelApp.ScreenUpdating = true;
            this.excelApp.Visible = true;
            Process excelProcess = Process.GetProcessesByName("Excel").Where(p => p.MainWindowTitle.Contains(this.excelFileNameWithoutExtension)).OrderByDescending(p => p.StartTime).First();
            Interaction.AppActivate(excelProcess.Id);
        }

        /// <summary>
        /// 指定された <see cref="Worksheet"/> をクリアします。
        /// </summary>
        /// <param name="worksheet"><see cref="Worksheet"/> オブジェクト。</param>
        public void Clear(Worksheet worksheet)
        {
            Range range = worksheet.Cells;
            this.Stack.Push(range);
            range.Clear();
        }

        /// <summary>
        /// 指定されたインデックスの <see cref="Worksheet"/> オブジェクトを取得します。
        /// </summary>
        /// <param name="sheets"><see cref="Sheets"/> オブジェクト。</param>
        /// <param name="sheetName">シート名。</param>
        /// <returns><see cref="Worksheet"/> オブジェクト。</returns>
        public Worksheet GetWorksheet(Sheets sheets, string sheetName)
        {
            Worksheet s = sheets[sheetName] as Worksheet;
            this.Stack.Push(s);
            return s;
        }

        /// <summary>
        /// 出力
        /// </summary>
        /// <param name="worksheet"><see cref="Worksheet"/> オブジェクト。</param>
        /// <param name="targetRange">出力範囲セル</param>
        /// <param name="values">出力値</param>
        public void Write(Worksheet worksheet, string targetRange, object[,] values)
        {
            Range range = this.GetRange(worksheet, targetRange);
            range.Value = values;
        }

        /// <summary>
        /// 出力
        /// </summary>
        /// <param name="worksheet"><see cref="Worksheet"/> オブジェクト。</param>
        /// <param name="targetRange">出力範囲セル</param>
        /// <param name="value">出力値</param>
        public void Write(Worksheet worksheet, string targetRange, object value)
        {
            Range range = this.GetRange(worksheet, targetRange);
            range.Value = value;
        }

        /// <summary>
        /// 行削除
        /// </summary>
        /// <param name="worksheet"><see cref="Worksheet"/> オブジェクト。</param>
        /// <param name="startRow">開始行</param>
        /// <param name="endRow">終了行</param>
        public void DeleteRows(Worksheet worksheet, int startRow, int endRow)
        {
            Range range = this.GetRange(worksheet, startRow, endRow);
            range.Delete();
        }

        /// <summary>
        /// 指定範囲のグラフオブジェクトを削除する
        /// </summary>
        /// <param name="worksheet"><see cref="Worksheet"/> オブジェクト。</param>
        /// <param name="startIndex">開始インデックス</param>
        /// <param name="endIndex">終了インデックス</param>
        public void DeleteChartObject(Worksheet worksheet, int startIndex, int endIndex)
        {
            foreach (ChartObject chtobj in worksheet.ChartObjects())
            {
                string chartNameKey = chtobj.Name.Substring(0, 2);
                int chartKey = Convert.ToInt32(chartNameKey);
                if (chartKey >= startIndex && chartKey <= endIndex)
                {
                    chtobj.Delete();
                }
            }
        }

        /// <summary>
        /// 列削除
        /// </summary>
        /// <param name="worksheet"><see cref="Worksheet"/> オブジェクト。</param>
        /// <param name="startCol">開始列</param>
        /// <param name="endCol">終了列</param>
        public void DeleteColumns(Worksheet worksheet, int startCol, int endCol)
        {
            Range range = this.GetColumns(worksheet, startCol, endCol);
            range.Delete(XlDeleteShiftDirection.xlShiftToLeft);
        }

        /// <summary>
        /// 指定列の削除
        /// </summary>
        /// <param name="worksheet"><see cref="Worksheet"/> オブジェクト。</param>
        /// <param name="targetRange">出力範囲セル[例："A1:B10"]</param>
        public void DeleteCells(Worksheet worksheet, string targetRange)
        {
            Range range = this.GetRange(worksheet, targetRange);
            range.Delete(XlDeleteShiftDirection.xlShiftToLeft);
        }

        /// <summary>
        /// 指定範囲のセルを削除
        /// </summary>
        /// <param name="worksheet"><see cref="Worksheet"/> オブジェクト。</param>
        /// <param name="startRow">開始行</param>
        /// <param name="startCol">開始列</param>
        /// <param name="endRow">終了行</param>
        /// <param name="endCol">終了列</param>
        public void DeleteCells(Worksheet worksheet, int startRow, int startCol, int endRow, int endCol)
        {
            Range startRrange = this.GetCell(worksheet, startRow, startCol);
            Range endRange = this.GetCell(worksheet, endRow, endCol);
            Range range = this.GetRange(worksheet, startRrange, endRange);

            range.Delete(XlDeleteShiftDirection.xlShiftToLeft);
        }

        /// <summary>
        /// 指定範囲のセルロックを設定する
        /// </summary>
        /// <param name="worksheet"><see cref="Worksheet"/> オブジェクト。</param>
        /// <param name="startRow">開始行</param>
        /// <param name="startCol">開始列</param>
        /// <param name="endRow">終了行</param>
        /// <param name="endCol">終了列</param>
        /// <param name="locked">true：ロックする、false：ロックしない</param>
        public void LockedCells(Worksheet worksheet, int startRow, int startCol, int endRow, int endCol, bool locked)
        {
            Range startRrange = this.GetCell(worksheet, startRow, startCol);
            Range endRange = this.GetCell(worksheet, endRow, endCol);
            Range range = this.GetRange(worksheet, startRrange, endRange);

            range.Locked = locked;
        }

        /// <summary>
        /// 列削除
        /// </summary>
        /// <param name="worksheet"><see cref="Worksheet"/> オブジェクト。</param>
        /// <param name="startCol">開始列</param>
        /// <param name="endCol">終了列</param>
        public void HiddenColumns(Worksheet worksheet, int startCol, int endCol)
        {
            Range range = this.GetColumns(worksheet, startCol, endCol);
            range.Hidden = true;
        }

        /// <summary>
        /// 数値をExcelのカラム文字へ変換します
        /// </summary>
        /// <param name="colIndex">列番号</param>
        /// <returns>Excelのカラム文字</returns>
        public string GetColumnName(int colIndex)
        {
            if (colIndex < 1)
            {
                return string.Empty;
            }

            return this.GetColumnName((colIndex - 1) / 26) + (char)('A' + ((colIndex - 1) % 26));
        }

        /// <summary>
        /// オートフィルタに対して抽出をかける
        /// </summary>
        /// <param name="worksheet"><see cref="Worksheet"/> オブジェクト。</param>
        /// <param name="targetCol">対象列</param>
        /// <param name="target">抽出文字列</param>
        public void AutoFilterColumns(Worksheet worksheet, int targetCol, string[] target)
        {
            Range range = this.GetRange(worksheet, 1, 1);
            range.AutoFilter(targetCol, target, XlAutoFilterOperator.xlFilterValues);
        }

        /// <summary>
        /// コメントを追加する
        /// </summary>
        /// <param name="range"><see cref="Range"/> オブジェクト。</param>
        /// <param name="comment">コメント</param>
        /// <param name="visible">コメントの表示／非表示[true：表示、false：非表示]</param>
        /// <param name="fontName">コメントに設定するフォント名</param>
        /// <param name="isAutoSize">true：AutoSizeを設定する、false：AutoSizeを設定しない</param>
        public void AddComment(Range range, string comment, bool visible, string fontName, bool isAutoSize)
        {
            // 対象セルにコメントが設定されている場合、対象セルのコメントをクリアする。
            if (range.Comment != null)
            {
                range.ClearComments();
            }

            range.AddComment(comment);
            range.Comment.Visible = visible;

            // フォント名が設定されている場合、コメントのフォントを設定する。
            if (string.IsNullOrWhiteSpace(fontName) == false)
            {
                int commentLength = 0;
                if (string.IsNullOrWhiteSpace(comment) == false)
                {
                    commentLength = comment.Length;
                }

                range.Comment.Shape.Placement = XlPlacement.xlMove;
                range.Comment.Shape.TextFrame.AutoSize = isAutoSize;
                range.Comment.Shape.TextFrame.Characters(0, commentLength).Font.Name = fontName;
            }
        }

        /// <summary>
        /// ハイパーリンクを設定する
        /// </summary>
        /// <param name="worksheet"><see cref="Worksheet"/> オブジェクト。</param>
        /// <param name="range"><see cref="Range"/> オブジェクト。</param>
        /// <param name="linkSheetName">リンク先シート名[例：sheet1!]</param>
        /// <param name="linkRange">リンク先セル[例：A1]</param>
        /// <param name="fontName">ハイパーリンクに設定するフォント名</param>
        public void AddHyperLink(
            Worksheet worksheet, 
            Range range, 
            string linkSheetName, 
            string linkRange,
            string fontName)
        {
            var hyperLinks = worksheet.Hyperlinks;
            this.Stack.Push(hyperLinks);
            hyperLinks.Add(range, string.Empty, linkSheetName + linkRange);

            if (string.IsNullOrWhiteSpace(fontName) == false)
            {
                range.Font.Name = fontName;
            }
        }

        /// <summary>
        /// 指定したシート名をアクティブにします。
        /// </summary>
        /// <param name="sheets"><see cref="Sheets"/> オブジェクト。</param>
        /// <param name="sheetName">シート名</param>
        public void SetActiveSheet(Sheets sheets, string sheetName)
        {
            var workSheetActive = this.GetWorksheet(sheets, sheetName);
            workSheetActive.Activate();
        }
    }
}
