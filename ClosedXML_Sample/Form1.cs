using ClosedXML.Attributes;
using ClosedXML.Excel;
using System.Collections;
using System.Data;

namespace ClosedXML_Sample
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e) { }

        /// <summary>
        /// クイックスタート
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            // 新規ワークブックを開く
            using var workbook = new XLWorkbook();

            // 新規ワークシートを追加
            var worksheet = workbook.AddWorksheet("Hello World");

            // A1セルにValue値を設定
            worksheet.Cell("A1").Value = "Hello World";

            // A2セルに関数を設定
            // MID(文字列, 開始位置, 文字数)
            worksheet.Cell("A2").FormulaA1 = "MID(A1, 7, 5)";

            // ワークブックを保存
            workbook.SaveAs("_HelloWorld.xlsx");
        }

        /// <summary>
        /// データ一括挿入 : DataTable
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            // データテーブルを用意
            var dataTable = new DataTable();
            dataTable.Columns.Add("Name", typeof(string));
            dataTable.Columns.Add("age", typeof(int));

            // データを追加
            dataTable.Rows.Add("明智", 40);
            dataTable.Rows.Add("石田", 23);
            dataTable.Rows.Add("宇喜多", 36);

            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();  // デフォルト

            // データを設定
            ws.Cell("B3").InsertData(dataTable);

            wb.SaveAs("_DataTable.xlsx");
        }

        /// <summary>
        /// データ一括挿入 : オブジェクトまたは構造体
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            var data = new[]
            {
                new User(1, "明智", 40),
                new User(2, "石田", 23),
                new User(3, "宇喜多", 36)
            };

            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("B3").InsertData(data);
            wb.SaveAs("_DataObjects.xlsx");
        }

        /// <summary>
        /// データ一括挿入 : IEnumerable
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            var data = new List<object[]>
            {
                new object[] { 1, "明智", 40 },
                new object[] { 2, "石田", 23 },
                new object[] { 3, "宇喜多", 36 },
            };

            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("B3").InsertData(data);
            wb.SaveAs("_DataIEnumerable.xlsx");
        }

        /// <summary>
        /// データ一括挿入 : Untyped Enumerable
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button5_Click(object sender, EventArgs e)
        {
            var data = new ArrayList
            {
                new object[] { 1, "明智", 40 },
                new User(2, "石田", 23),
                new object[] { 3, "宇喜多", 36 },
            };

            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("B3").InsertData(data);
            wb.SaveAs("_DataUntypedEnumerable.xlsx");
        }

        /// <summary>
        /// テーブル作成
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button6_Click(object sender, EventArgs e)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            // 全ての列幅を設定
            ws.ColumnWidth = 30;

            // 範囲内の左上のセルに、テーブルを作成
            ws.FirstCell().InsertTable(new[]
            {
                new User(1, "明智", 40),
                new User(2, "石田", 23),
                new User(3, "宇喜多", 36),
                new User(4, "遠藤", 58),
                new User(5, "織田", 28)
            }, "ユーザー名", true);

            // 範囲を指定し、空のテーブルを作成
            ws.Range("D2:D5").CreateTable();

            wb.SaveAs("_TableCreate.xlsx");
        }

        /// <summary>
        /// テーブルテーマ
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button7_Click(object sender, EventArgs e)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            // Value値は、SetValue()でも設定可能
            ws.Cell("A1").SetValue("ファースト");

            // Rangeで、整数の範囲を一括設定できる
            ws.Cell("A2").InsertData(Enumerable.Range(1, 5));

            ws.Cell("B1").SetValue("セカンド");
            ws.Cell("B2").InsertData(Enumerable.Range(1, 5));

            // テーブルの範囲を決めて、テーマを決める
            var table = ws.Range("A1:B6").CreateTable();
            table.Theme = XLTableTheme.TableStyleLight16;

            // テーブルをコピーして、テーマを変更する
            table = table.CopyTo(ws.Cell("D1")).CreateTable();
            table.Theme = XLTableTheme.TableStyleDark2;

            table = table.CopyTo(ws.Cell("G1")).CreateTable();
            table.Theme = XLTableTheme.TableStyleMedium15;

            wb.SaveAs("_TableTheme.xlsx");
        }

        /// <summary>
        /// テーブルスタイルオプション
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button8_Click(object sender, EventArgs e)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            ws.Cell("A1").SetValue("First");
            ws.Cell("A2").InsertData(Enumerable.Range(1, 5));
            ws.Cell("B1").SetValue("Second");
            ws.Cell("B2").InsertData(Enumerable.Range(1, 5));

            var table = ws.Range("A1:B6").CreateTable();
            table.CopyTo(ws.Cell("D1")).CreateTable();

            // テーブルスタイプオプション設定
            table
                .SetShowHeaderRow(false)
                .SetShowRowStripes(false)
                .SetShowColumnStripes(true)
                .SetShowAutoFilter(true)
                .SetShowTotalsRow(true);

            table.Field("First").TotalsRowFunction = XLTotalsRowFunction.Sum;
            table.Field("Second").TotalsRowFunction = XLTotalsRowFunction.Average;

            wb.SaveAs("_TableStyleOptions.xlsx");
        }

        /// <summary>
        /// フォント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button9_Click(object sender, EventArgs e)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            // フォント設定
            ws.Cell("A1").Style
                .Font.SetFontSize(20)
                .Font.SetFontName("Arial");

            // バックグラウンドカラー
            ws.Cell("A1").Style
                .Fill.SetBackgroundColor(XLColor.Red);

            // ボーダー
            ws.Cell("B2").Style
                .Border.SetTopBorder(XLBorderStyleValues.Medium)
                .Border.SetRightBorder(XLBorderStyleValues.Medium)
                .Border.SetBottomBorder(XLBorderStyleValues.Medium)
                .Border.SetLeftBorder(XLBorderStyleValues.Medium);
        }
    }

    public record User
        (
            [property: XLColumn(Ignore = true)] int ID,
            [property: XLColumn(Order = 1)] string Name,
            [property: XLColumn(Order = 2)] int age
        );
}