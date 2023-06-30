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
        /// �N�C�b�N�X�^�[�g
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            // �V�K���[�N�u�b�N���J��
            using var workbook = new XLWorkbook();

            // �V�K���[�N�V�[�g��ǉ�
            var worksheet = workbook.AddWorksheet("Hello World");

            // A1�Z����Value�l��ݒ�
            worksheet.Cell("A1").Value = "Hello World";

            // A2�Z���Ɋ֐���ݒ�
            // MID(������, �J�n�ʒu, ������)
            worksheet.Cell("A2").FormulaA1 = "MID(A1, 7, 5)";

            // ���[�N�u�b�N��ۑ�
            workbook.SaveAs("_HelloWorld.xlsx");
        }

        /// <summary>
        /// �f�[�^�ꊇ�}�� : DataTable
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            // �f�[�^�e�[�u����p��
            var dataTable = new DataTable();
            dataTable.Columns.Add("Name", typeof(string));
            dataTable.Columns.Add("age", typeof(int));

            // �f�[�^��ǉ�
            dataTable.Rows.Add("���q", 40);
            dataTable.Rows.Add("�Γc", 23);
            dataTable.Rows.Add("�F�쑽", 36);

            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();  // �f�t�H���g

            // �f�[�^��ݒ�
            ws.Cell("B3").InsertData(dataTable);

            wb.SaveAs("_DataTable.xlsx");
        }

        /// <summary>
        /// �f�[�^�ꊇ�}�� : �I�u�W�F�N�g�܂��͍\����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            var data = new[]
            {
                new User(1, "���q", 40),
                new User(2, "�Γc", 23),
                new User(3, "�F�쑽", 36)
            };

            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("B3").InsertData(data);
            wb.SaveAs("_DataObjects.xlsx");
        }

        /// <summary>
        /// �f�[�^�ꊇ�}�� : IEnumerable
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            var data = new List<object[]>
            {
                new object[] { 1, "���q", 40 },
                new object[] { 2, "�Γc", 23 },
                new object[] { 3, "�F�쑽", 36 },
            };

            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("B3").InsertData(data);
            wb.SaveAs("_DataIEnumerable.xlsx");
        }

        /// <summary>
        /// �f�[�^�ꊇ�}�� : Untyped Enumerable
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button5_Click(object sender, EventArgs e)
        {
            var data = new ArrayList
            {
                new object[] { 1, "���q", 40 },
                new User(2, "�Γc", 23),
                new object[] { 3, "�F�쑽", 36 },
            };

            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("B3").InsertData(data);
            wb.SaveAs("_DataUntypedEnumerable.xlsx");
        }

        /// <summary>
        /// �e�[�u���쐬
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button6_Click(object sender, EventArgs e)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            // �S�Ă̗񕝂�ݒ�
            ws.ColumnWidth = 30;

            // �͈͓��̍���̃Z���ɁA�e�[�u�����쐬
            ws.FirstCell().InsertTable(new[]
            {
                new User(1, "���q", 40),
                new User(2, "�Γc", 23),
                new User(3, "�F�쑽", 36),
                new User(4, "����", 58),
                new User(5, "�D�c", 28)
            }, "���[�U�[��", true);

            // �͈͂��w�肵�A��̃e�[�u�����쐬
            ws.Range("D2:D5").CreateTable();

            wb.SaveAs("_TableCreate.xlsx");
        }

        /// <summary>
        /// �e�[�u���e�[�}
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button7_Click(object sender, EventArgs e)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            // Value�l�́ASetValue()�ł��ݒ�\
            ws.Cell("A1").SetValue("�t�@�[�X�g");

            // Range�ŁA�����͈̔͂��ꊇ�ݒ�ł���
            ws.Cell("A2").InsertData(Enumerable.Range(1, 5));

            ws.Cell("B1").SetValue("�Z�J���h");
            ws.Cell("B2").InsertData(Enumerable.Range(1, 5));

            // �e�[�u���͈̔͂����߂āA�e�[�}�����߂�
            var table = ws.Range("A1:B6").CreateTable();
            table.Theme = XLTableTheme.TableStyleLight16;

            // �e�[�u�����R�s�[���āA�e�[�}��ύX����
            table = table.CopyTo(ws.Cell("D1")).CreateTable();
            table.Theme = XLTableTheme.TableStyleDark2;

            table = table.CopyTo(ws.Cell("G1")).CreateTable();
            table.Theme = XLTableTheme.TableStyleMedium15;

            wb.SaveAs("_TableTheme.xlsx");
        }

        /// <summary>
        /// �e�[�u���X�^�C���I�v�V����
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

            // �e�[�u���X�^�C�v�I�v�V�����ݒ�
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
        /// �t�H���g
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button9_Click(object sender, EventArgs e)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            // �t�H���g�ݒ�
            ws.Cell("A1").Style
                .Font.SetFontSize(20)
                .Font.SetFontName("Arial");

            // �o�b�N�O���E���h�J���[
            ws.Cell("A1").Style
                .Fill.SetBackgroundColor(XLColor.Red);

            // �{�[�_�[
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