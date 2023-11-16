using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace WinFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void CreateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ClearDataGridView();
        }

        private void OpenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new()
            {
                Title = "Выберите файл Excel",
                Filter = "Файлы Excel|*.xls;*.xlsx|Все файлы|*.*",
                CheckFileExists = true,
                CheckPathExists = true
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog1.FileName;
                LoadDataFromExcel(filePath);
            }
        }

        private void SaveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new()
            {
                Title = "Выберите файл Excel",
                Filter = "Файлы Excel|*.xls;*.xlsx|Все файлы|*.*",
                CheckFileExists = true,
                CheckPathExists = true
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog1.FileName;
                CopyDataGridViewToExcel(filePath);
            }
        }

        private void ClearDataGridView()
        {
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();
            dataGridView1.ColumnCount = 1;
        }

        private void LoadDataFromExcel(string filePath)
        {
            Excel.Application excelApp = new();
            Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1];
            try
            {
                int rows = worksheet.UsedRange.Rows.Count;
                int columns = worksheet.UsedRange.Columns.Count;
                dataGridView1.ColumnCount = columns;
                for (int i = 1; i <= columns; i++)
                {
                    dataGridView1.Columns[i - 1].HeaderText = $"Column {i}";
                }
                for (int i = 1; i <= rows; i++)
                {
                    dataGridView1.Rows.Add();
                    for (int j = 1; j <= columns; j++)
                    {
                        dataGridView1.Rows[i - 1].Cells[j - 1].Value = worksheet.Cells[i, j].Value;
                    }
                }
                workbook.Close(false);
                excelApp.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(excelApp);
            }
        }

        private void CopyDataGridViewToExcel(string filePath)
        {
            Excel.Application excelApp = new()
            {
                Visible = true
            };
            Excel.Workbook workbook = excelApp.Workbooks.Add();
            Excel.Worksheet worksheet = workbook.ActiveSheet;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    worksheet.Cells[i + 1, j + 1] = dataGridView1.Rows[i].Cells[j].Value;
                }
            }
            workbook.SaveAs(filePath);
            workbook.Close();
            excelApp.Quit();
            Marshal.ReleaseComObject(worksheet);
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(excelApp);
            MessageBox.Show("Данные успешно экспортированы в Excel.", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}