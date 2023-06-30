using System;
using System.Reflection;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel_Operation_with_C_
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;
            Excel.Range oRng;

            try
            {
                // Excelを起動し、Applicationオブジェクトを取得する
                oXL = new Excel.Application
                {
                    // 可視化
                    Visible = true
                };

                // 新しいワークブックを取得する
                oWB = oXL.Workbooks.Add();
                oSheet = (Excel._Worksheet)oWB.ActiveSheet;

                // テーブルヘッダを追加する（A1:D1）
                oSheet.Cells[1, 1] = "First Name";
                oSheet.Cells[1, 2] = "Last Name";
                oSheet.Cells[1, 3] = "Full Name";
                oSheet.Cells[1, 4] = "Salary";

                // セルの書式を変更する（A1:D1）
                oSheet.get_Range("A1", "D1").Font.Bold = true;
                oSheet.get_Range("A1", "D1").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                // 複数の値が設定された配列を作成
                string[,] saNames = new string[5, 2];

                saNames[0, 0] = "John";
                saNames[0, 1] = "Smith";
                saNames[1, 0] = "Tom";
                saNames[1, 1] = "Brown";
                saNames[2, 0] = "Sue";
                saNames[2, 1] = "Thomas";
                saNames[3, 0] = "Jane";
                saNames[3, 1] = "Jones";
                saNames[4, 0] = "Adam";
                saNames[4, 1] = "Johnson";

                // セルの値を配列で埋める（A2:B6）
                oSheet.get_Range("A2", "B6").Value2 = saNames;

                // 数式で埋める（C2:C6）
                oRng = oSheet.get_Range("C2", "C6");
                oRng.Formula = "=A2 & \" \" & B2";

                // 数式で埋め、フォーマットを適用する（D2:D6）
                oRng = oSheet.get_Range("D2", "D6");
                oRng.Formula = "=RAND()*100000";
                oRng.NumberFormat = "$0.00";

                // カラムの幅を調整する（A:D）
                oRng = oSheet.get_Range("A1", "D1");
                oRng.EntireColumn.AutoFit();

                // 四半期売上データのカラム数を可変にする
                DisplayQuarterlySales(oSheet);

                // エクセルを表示する
                // エクセルを操作できるようにする
                oXL.Visible = true;
                oXL.UserControl = true;
            }
            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, "Error");
            }
        }

        private void DisplayQuarterlySales(Excel._Worksheet oWS)
        {
            Excel._Workbook oWB;

            // グラフのデータ系列を表します
            Excel.Series oSeries;
            Excel.Range oResizeRange;
            Excel._Chart oChart;
            String sMsg;
            int iNumQtrs;

            // 何四半期分のデータを表示するかを決定する
            for (iNumQtrs = 4; iNumQtrs >= 2; iNumQtrs--)
            {
                sMsg = "Enter sales data for ";
                sMsg = String.Concat(sMsg, iNumQtrs);
                sMsg = String.Concat(sMsg, " quarter(s)?");

                DialogResult iRet = MessageBox.Show(sMsg, "Quarterly Sales?", MessageBoxButtons.YesNo);
                if (iRet == DialogResult.Yes)
                    break;
            }

            sMsg = "Displaying data for ";
            sMsg = String.Concat(sMsg, iNumQtrs);
            sMsg = String.Concat(sMsg, " quarter(s).");

            MessageBox.Show(sMsg, "Quarterly Sales");

            // E1セルから順に、選択した列数分のヘッダーを埋める
            oResizeRange = oWS.get_Range("E1", "E1").get_Resize(Missing.Value, iNumQtrs);
            oResizeRange.Formula = "=\"Q\" & COLUMN()-4 & CHAR(10) & \"Sales\"";

            // ヘッダーのプロパティを変更する
            oResizeRange.Orientation = 38;
            oResizeRange.WrapText = true;

            // ヘッダーの背景色を変更する
            oResizeRange.Interior.ColorIndex = 36;

            // 列を数式で埋め、数値の書式を適用する
            oResizeRange = oWS.get_Range("E2", "E6").get_Resize(Missing.Value, iNumQtrs);
            oResizeRange.Formula = "=RAND()*100";
            oResizeRange.NumberFormat = "$0.00";

            // 売上データおよび、ヘッダーにボーダーを適用する
            oResizeRange = oWS.get_Range("E1", "E6").get_Resize(Missing.Value, iNumQtrs);
            oResizeRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

            // 売上データのトータルを追加し、ボーダーを追加する
            oResizeRange = oWS.get_Range("E8", "E8").get_Resize(Missing.Value, iNumQtrs);
            oResizeRange.Formula = "=SUM(E2:E6)";
            oResizeRange.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle
            = Excel.XlLineStyle.xlDouble;
            oResizeRange.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight
            = Excel.XlBorderWeight.xlThick;

            // Parentプロパティ
            // 親階層（今回はWorkbook）を取得
            oWB = (Excel._Workbook)oWS.Parent;

            // Sheets.Add(Object, Object, Object, Object)メソッド
            // 新しいワークシート、グラフシート、またはマクロシートを作成する
            oChart = (Excel._Chart)oWB.Charts.Add();

            // ChartWizardを使用して、選択したデータから新しいチャートを作成する
            oResizeRange = oWS.get_Range("E2:E6").get_Resize(
            Missing.Value, iNumQtrs);
            oChart.ChartWizard(oResizeRange, Excel.XlChartType.xl3DColumn, Missing.Value, Excel.XlRowCol.xlColumns);
            oSeries = (Excel.Series)oChart.SeriesCollection(1);
            oSeries.XValues = oWS.get_Range("A2", "A6");
            for (int iRet = 1; iRet <= iNumQtrs; iRet++)
            {
                oSeries = (Excel.Series)oChart.SeriesCollection(iRet);
                String seriesName;
                seriesName = "=\"Q";
                seriesName = String.Concat(seriesName, iRet);
                seriesName = String.Concat(seriesName, "\"");
                oSeries.Name = seriesName;
            }

            oChart.Location(Excel.XlChartLocation.xlLocationAsObject, oWS.Name);

            // データを覆わないように、チャートを移動させる
            oResizeRange = (Excel.Range)oWS.Rows.get_Item(10);
            oWS.Shapes.Item("Chart 1").Top = (float)(double)oResizeRange.Top;
            oResizeRange = (Excel.Range)oWS.Columns.get_Item(2);
            oWS.Shapes.Item("Chart 1").Left = (float)(double)oResizeRange.Left;
        }
    }
}
