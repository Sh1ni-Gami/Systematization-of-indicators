using System;
using System.Collections.Generic;
using System.ComponentModel;
using SD = System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using MySql.Data.MySqlClient;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelDataReader;
using System.Data;
using Mysqlx.Expr;
using System.IO;
using Microsoft.Office.Interop.Excel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Programm
{

    public partial class Form2 : Form
    {


        public Form2()
        {
            InitializeComponent();
            StartPosition = FormStartPosition.CenterScreen;


        }

        private void Data()
        {


        }



        private void Form2_Load(object sender, EventArgs e)
        {
            Data();
            


        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }




        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            (dataGridView1.DataSource as SD.DataTable).DefaultView.RowFilter = $"[Субъект Российской Федерации] LIKE '%{textBox1.Text}%'";
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            (dataGridView2.DataSource as SD.DataTable).DefaultView.RowFilter = $"[Субъект Российской Федерации] LIKE '%{textBox2.Text}%'";
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            (dataGridView3.DataSource as SD.DataTable).DefaultView.RowFilter = $"[Субъект Российской Федерации] LIKE '%{textBox3.Text}%'";
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            (dataGridView4.DataSource as SD.DataTable).DefaultView.RowFilter = $"[Субъект Российской Федерации] LIKE '%{textBox4.Text}%'";
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            (dataGridView5.DataSource as SD.DataTable).DefaultView.RowFilter = $"[Субъект Российской Федерации] LIKE '%{textBox5.Text}%'";
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            (dataGridView6.DataSource as SD.DataTable).DefaultView.RowFilter = $"[Субъект Российской Федерации] LIKE '%{textBox6.Text}%'";
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            (dataGridView7.DataSource as SD.DataTable).DefaultView.RowFilter = $"[Субъект Российской Федерации] LIKE '%{textBox7.Text}%'";
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            (dataGridView8.DataSource as SD.DataTable).DefaultView.RowFilter = $"[Субъект Российской Федерации] LIKE '%{textBox8.Text}%'";
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            (dataGridView9.DataSource as SD.DataTable).DefaultView.RowFilter = $"[Субъект Российской Федерации] LIKE '%{textBox9.Text}%'";
        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {
            (dataGridView10.DataSource as SD.DataTable).DefaultView.RowFilter = $"[Субъект Российской Федерации] LIKE '%{textBox10.Text}%'";
        }

        private void textBox11_TextChanged(object sender, EventArgs e)
        {
            (dataGridView11.DataSource as SD.DataTable).DefaultView.RowFilter = $"[Субъект Российской Федерации] LIKE '%{textBox11.Text}%'";
        }

        private void button_loud1_Click(object sender, EventArgs e)
        {
            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Public\\Documents\\Файлы\\Кол-во ДТП.xlsx");
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1];

            int rowCount = worksheet.UsedRange.Rows.Count;
            int colCount = worksheet.UsedRange.Columns.Count;
            object[,] data = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[rowCount, colCount]].Value;

            SD.DataTable dt = new SD.DataTable();
            for (int i = 1; i <= colCount; i++)
            {
                dt.Columns.Add(new DataColumn(data[1, i].ToString()));
            }
            for (int row = 2; row <= rowCount; row++)
            {
                DataRow dr = dt.NewRow();
                for (int col = 1; col <= colCount; col++)
                {
                    dr[col - 1] = data[row, col];
                }
                dt.Rows.Add(dr);
            }

            dataGridView1.DataSource = dt;

            workbook.Close();
            excel.Quit();
        }

        private void button_save1_Click(object sender, EventArgs e)
        {
            // Путь к файлу Excel
            string filePath = "C:\\Users\\Public\\Documents\\Файлы\\Кол-во ДТП.xlsx";

            // Создание экземпляра приложения Excel
            Excel.Application excel = new Excel.Application();

            // Открытие файла Excel
            Excel.Workbook workbook = excel.Workbooks.Open(filePath);

            // Выбор листа для очистки данных
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[1];

            // Очистка данных на листе
            worksheet.UsedRange.ClearContents();

            // Сохранение изменений
            workbook.Save();

            // Определяем столбцы и строки для DataGridView
            int columnsCount = dataGridView1.Columns.Count;
            int rowsCount = dataGridView1.Rows.Count;

            for (int i = 0; i <= rowsCount - 2; i++)
            {
                for (int j = 0; j <= columnsCount - 1; j++)
                {
                    worksheet.Cells[1, j + 1] = dataGridView1.Columns[j].HeaderText.ToString();
                    worksheet.Cells[i + 2, j + 1] = dataGridView1[j, i].Value.ToString();
                }
            }

            workbook.Save();
            workbook.Close();
            // Закрытие приложения Excel
            excel.Quit(); 
            MessageBox.Show("Изменения сохранены");

        }

        private void button_export1_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Add();
            Excel.Worksheet excelWorksheet = excelWorkbook.ActiveSheet;

            // Заполнение заголовков столбцов
            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                excelWorksheet.Cells[1, i + 1] = dataGridView1.Columns[i].HeaderText;
            }

            // Заполнение ячеек таблицы данными
            for (int i = 0; i < dataGridView1.SelectedRows.Count; i++)
            {
                DataGridViewRow row = dataGridView1.SelectedRows[i];
                for (int j = 0; j < row.Cells.Count; j++)
                {
                    excelWorksheet.Cells[i + 2, j + 1] = row.Cells[j].Value.ToString();
                }
            }
            excelApp.Visible = true;
            
        }

        private void button_loud2_Click(object sender, EventArgs e)
        {
            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Public\\Documents\\Файлы\\Кол-во выездов ПСП.xlsx");
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1];

            int rowCount = worksheet.UsedRange.Rows.Count;
            int colCount = worksheet.UsedRange.Columns.Count;
            object[,] data = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[rowCount, colCount]].Value;

            SD.DataTable dt = new SD.DataTable();
            for (int i = 1; i <= colCount; i++)
            {
                dt.Columns.Add(new DataColumn(data[1, i].ToString()));
            }
            for (int row = 2; row <= rowCount; row++)
            {
                DataRow dr = dt.NewRow();
                for (int col = 1; col <= colCount; col++)
                {
                    dr[col - 1] = data[row, col];
                }
                dt.Rows.Add(dr);
            }

            dataGridView2.DataSource = dt;

            workbook.Close();
            excel.Quit();
        }

        private void button_save2_Click(object sender, EventArgs e)
        {
            // Путь к файлу Excel
            string filePath = "C:\\Users\\Public\\Documents\\Файлы\\Кол-во выездов ПСП.xlsx";

            // Создание экземпляра приложения Excel
            Excel.Application excel = new Excel.Application();

            // Открытие файла Excel
            Excel.Workbook workbook = excel.Workbooks.Open(filePath);

            // Выбор листа для очистки данных
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[1];

            // Очистка данных на листе
            worksheet.UsedRange.ClearContents();

            // Сохранение изменений
            workbook.Save();

            // Определяем столбцы и строки для DataGridView
            int columnsCount = dataGridView2.Columns.Count;
            int rowsCount = dataGridView2.Rows.Count;

            for (int i = 0; i <= rowsCount - 2; i++)
            {
                for (int j = 0; j <= columnsCount - 1; j++)
                {
                    worksheet.Cells[1, j + 1] = dataGridView2.Columns[j].HeaderText.ToString();
                    worksheet.Cells[i + 2, j + 1] = dataGridView2[j, i].Value.ToString();
                }
            }

            workbook.Save();
            workbook.Close();
            // Закрытие приложения Excel
            excel.Quit();
            MessageBox.Show("Изменения сохранены");
        }

        private void button_export2_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Add();
            Excel.Worksheet excelWorksheet = excelWorkbook.ActiveSheet;

            // Заполнение заголовков столбцов
            for (int i = 0; i < dataGridView2.Columns.Count; i++)
            {
                excelWorksheet.Cells[1, i + 1] = dataGridView2.Columns[i].HeaderText;
            }

            // Заполнение ячеек таблицы данными
            for (int i = 0; i < dataGridView2.SelectedRows.Count; i++)
            {
                DataGridViewRow row = dataGridView2.SelectedRows[i];
                for (int j = 0; j < row.Cells.Count; j++)
                {
                    excelWorksheet.Cells[i + 2, j + 1] = row.Cells[j].Value.ToString();
                }
            }
            excelApp.Visible = true;
        }

        private void button_loud3_Click(object sender, EventArgs e)
        {
            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Public\\Documents\\Файлы\\Кол-во граж кот оказана пом.xlsx");
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1];

            int rowCount = worksheet.UsedRange.Rows.Count;
            int colCount = worksheet.UsedRange.Columns.Count;
            object[,] data = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[rowCount, colCount]].Value;

            SD.DataTable dt = new SD.DataTable();
            for (int i = 1; i <= colCount; i++)
            {
                dt.Columns.Add(new DataColumn(data[1, i].ToString()));
            }
            for (int row = 2; row <= rowCount; row++)
            {
                DataRow dr = dt.NewRow();
                for (int col = 1; col <= colCount; col++)
                {
                    dr[col - 1] = data[row, col];
                }
                dt.Rows.Add(dr);
            }

            dataGridView3.DataSource = dt;

            workbook.Close();
            excel.Quit();
        }

        private void button_save3_Click(object sender, EventArgs e)
        {
            // Путь к файлу Excel
            string filePath = "C:\\Users\\Public\\Documents\\Файлы\\Кол-во граж., кот. оказана пом.xlsx";

            // Создание экземпляра приложения Excel
            Excel.Application excel = new Excel.Application();

            // Открытие файла Excel
            Excel.Workbook workbook = excel.Workbooks.Open(filePath);

            // Выбор листа для очистки данных
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[1];

            // Очистка данных на листе
            worksheet.UsedRange.ClearContents();

            // Сохранение изменений
            workbook.Save();

            // Определяем столбцы и строки для DataGridView
            int columnsCount = dataGridView3.Columns.Count;
            int rowsCount = dataGridView3.Rows.Count;

            for (int i = 0; i <= rowsCount - 2; i++)
            {
                for (int j = 0; j <= columnsCount - 1; j++)
                {
                    worksheet.Cells[1, j + 1] = dataGridView3.Columns[j].HeaderText.ToString();
                    worksheet.Cells[i + 2, j + 1] = dataGridView3[j, i].Value.ToString();
                }
            }

            workbook.Save();
            workbook.Close();
            // Закрытие приложения Excel
            excel.Quit();
            MessageBox.Show("Изменения сохранены");
        }

        private void button_export3_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Add();
            Excel.Worksheet excelWorksheet = excelWorkbook.ActiveSheet;

            // Заполнение заголовков столбцов
            for (int i = 0; i < dataGridView3.Columns.Count; i++)
            {
                excelWorksheet.Cells[1, i + 1] = dataGridView3.Columns[i].HeaderText;
            }

            // Заполнение ячеек таблицы данными
            for (int i = 0; i < dataGridView3.SelectedRows.Count; i++)
            {
                DataGridViewRow row = dataGridView3.SelectedRows[i];
                for (int j = 0; j < row.Cells.Count; j++)
                {
                    excelWorksheet.Cells[i + 2, j + 1] = row.Cells[j].Value.ToString();
                }
            }
            excelApp.Visible = true;
        }

        private void button_loud4_Click(object sender, EventArgs e)
        {
            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Public\\Documents\\Файлы\\Количество деблокированных.xlsx");
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1];

            int rowCount = worksheet.UsedRange.Rows.Count;
            int colCount = worksheet.UsedRange.Columns.Count;
            object[,] data = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[rowCount, colCount]].Value;

            SD.DataTable dt = new SD.DataTable();
            for (int i = 1; i <= colCount; i++)
            {
                dt.Columns.Add(new DataColumn(data[1, i].ToString()));
            }
            for (int row = 2; row <= rowCount; row++)
            {
                DataRow dr = dt.NewRow();
                for (int col = 1; col <= colCount; col++)
                {
                    dr[col - 1] = data[row, col];
                }
                dt.Rows.Add(dr);
            }

            dataGridView4.DataSource = dt;

            workbook.Close();
            excel.Quit();
        }

        private void button_save4_Click(object sender, EventArgs e)
        {
            // Путь к файлу Excel
            string filePath = "C:\\Users\\Public\\Documents\\Файлы\\Количество деблокированных.xlsx";

            // Создание экземпляра приложения Excel
            Excel.Application excel = new Excel.Application();

            // Открытие файла Excel
            Excel.Workbook workbook = excel.Workbooks.Open(filePath);

            // Выбор листа для очистки данных
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[1];

            // Очистка данных на листе
            worksheet.UsedRange.ClearContents();

            // Сохранение изменений
            workbook.Save();

            // Определяем столбцы и строки для DataGridView
            int columnsCount = dataGridView4.Columns.Count;
            int rowsCount = dataGridView4.Rows.Count;

            for (int i = 0; i <= rowsCount - 2; i++)
            {
                for (int j = 0; j <= columnsCount - 1; j++)
                {
                    worksheet.Cells[1, j + 1] = dataGridView4.Columns[j].HeaderText.ToString();
                    worksheet.Cells[i + 2, j + 1] = dataGridView4[j, i].Value.ToString();
                }
            }

            workbook.Save();
            workbook.Close();
            // Закрытие приложения Excel
            excel.Quit();
            MessageBox.Show("Изменения сохранены");
        }

        private void button_export4_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Add();
            Excel.Worksheet excelWorksheet = excelWorkbook.ActiveSheet;

            // Заполнение заголовков столбцов
            for (int i = 0; i < dataGridView4.Columns.Count; i++)
            {
                excelWorksheet.Cells[1, i + 1] = dataGridView4.Columns[i].HeaderText;
            }

            // Заполнение ячеек таблицы данными
            for (int i = 0; i < dataGridView4.SelectedRows.Count; i++)
            {
                DataGridViewRow row = dataGridView4.SelectedRows[i];
                for (int j = 0; j < row.Cells.Count; j++)
                {
                    excelWorksheet.Cells[i + 2, j + 1] = row.Cells[j].Value.ToString();
                }
            }
            excelApp.Visible = true;
        }

        private void button_loud5_Click(object sender, EventArgs e)
        {
            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Public\\Documents\\Файлы\\Среднее время прибытия.xlsx");
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1];

            int rowCount = worksheet.UsedRange.Rows.Count;
            int colCount = worksheet.UsedRange.Columns.Count;
            object[,] data = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[rowCount, colCount]].Value;

            SD.DataTable dt = new SD.DataTable();
            for (int i = 1; i <= colCount; i++)
            {
                dt.Columns.Add(new DataColumn(data[1, i].ToString()));
            }
            for (int row = 2; row <= rowCount; row++)
            {
                DataRow dr = dt.NewRow();
                for (int col = 1; col <= colCount; col++)
                {
                    dr[col - 1] = data[row, col];
                }
                dt.Rows.Add(dr);
            }

            dataGridView5.DataSource = dt;

            workbook.Close();
            excel.Quit();
        }

        private void button_save5_Click(object sender, EventArgs e)
        {
            // Путь к файлу Excel
            string filePath = "C:\\Users\\Public\\Documents\\Файлы\\Среднее время прибытия.xlsx";

            // Создание экземпляра приложения Excel
            Excel.Application excel = new Excel.Application();

            // Открытие файла Excel
            Excel.Workbook workbook = excel.Workbooks.Open(filePath);

            // Выбор листа для очистки данных
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[1];

            // Очистка данных на листе
            worksheet.UsedRange.ClearContents();

            // Сохранение изменений
            workbook.Save();

            // Определяем столбцы и строки для DataGridView
            int columnsCount = dataGridView5.Columns.Count;
            int rowsCount = dataGridView5.Rows.Count;

            for (int i = 0; i <= rowsCount - 2; i++)
            {
                for (int j = 0; j <= columnsCount - 1; j++)
                {
                    worksheet.Cells[1, j + 1] = dataGridView5.Columns[j].HeaderText.ToString();
                    worksheet.Cells[i + 2, j + 1] = dataGridView5[j, i].Value.ToString();
                }
            }

            workbook.Save();
            workbook.Close();
            // Закрытие приложения Excel
            excel.Quit();
            MessageBox.Show("Изменения сохранены");
        }

        private void button_export5_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Add();
            Excel.Worksheet excelWorksheet = excelWorkbook.ActiveSheet;

            // Заполнение заголовков столбцов
            for (int i = 0; i < dataGridView5.Columns.Count; i++)
            {
                excelWorksheet.Cells[1, i + 1] = dataGridView5.Columns[i].HeaderText;
            }

            // Заполнение ячеек таблицы данными
            for (int i = 0; i < dataGridView5.SelectedRows.Count; i++)
            {
                DataGridViewRow row = dataGridView5.SelectedRows[i];
                for (int j = 0; j < row.Cells.Count; j++)
                {
                    excelWorksheet.Cells[i + 2, j + 1] = row.Cells[j].Value.ToString();
                }
            }
            excelApp.Visible = true;
        }

        private void button_loud6_Click(object sender, EventArgs e)
        {
            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Public\\Documents\\Файлы\\Коэффициент реагирования.xlsx");
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1];

            int rowCount = worksheet.UsedRange.Rows.Count;
            int colCount = worksheet.UsedRange.Columns.Count;
            object[,] data = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[rowCount, colCount]].Value;

            SD.DataTable dt = new SD.DataTable();
            for (int i = 1; i <= colCount; i++)
            {
                dt.Columns.Add(new DataColumn(data[1, i].ToString()));
            }
            for (int row = 2; row <= rowCount; row++)
            {
                DataRow dr = dt.NewRow();
                for (int col = 1; col <= colCount; col++)
                {
                    dr[col - 1] = data[row, col];
                }
                dt.Rows.Add(dr);
            }

            dataGridView6.DataSource = dt;

            workbook.Close();
            excel.Quit();
        }

        private void button_save6_Click(object sender, EventArgs e)
        {
            // Путь к файлу Excel
            string filePath = "C:\\Users\\Public\\Documents\\Файлы\\Коэффициент реагирования.xlsx";

            // Создание экземпляра приложения Excel
            Excel.Application excel = new Excel.Application();

            // Открытие файла Excel
            Excel.Workbook workbook = excel.Workbooks.Open(filePath);

            // Выбор листа для очистки данных
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[1];

            // Очистка данных на листе
            worksheet.UsedRange.ClearContents();

            // Сохранение изменений
            workbook.Save();

            // Определяем столбцы и строки для DataGridView
            int columnsCount = dataGridView6.Columns.Count;
            int rowsCount = dataGridView6.Rows.Count;

            for (int i = 0; i <= rowsCount - 2; i++)
            {
                for (int j = 0; j <= columnsCount - 1; j++)
                {
                    worksheet.Cells[1, j + 1] = dataGridView6.Columns[j].HeaderText.ToString();
                    worksheet.Cells[i + 2, j + 1] = dataGridView6[j, i].Value.ToString();
                }
            }

            workbook.Save();
            workbook.Close();
            // Закрытие приложения Excel
            excel.Quit();
            MessageBox.Show("Изменения сохранены");
        }

        private void button_export6_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Add();
            Excel.Worksheet excelWorksheet = excelWorkbook.ActiveSheet;

            // Заполнение заголовков столбцов
            for (int i = 0; i < dataGridView6.Columns.Count; i++)
            {
                excelWorksheet.Cells[1, i + 1] = dataGridView6.Columns[i].HeaderText;
            }

            // Заполнение ячеек таблицы данными
            for (int i = 0; i < dataGridView6.SelectedRows.Count; i++)
            {
                DataGridViewRow row = dataGridView6.SelectedRows[i];
                for (int j = 0; j < row.Cells.Count; j++)
                {
                    excelWorksheet.Cells[i + 2, j + 1] = row.Cells[j].Value.ToString();
                }
            }
            excelApp.Visible = true;
        }

        private void button_loud7_Click(object sender, EventArgs e)
        {
            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Public\\Documents\\Файлы\\Работа на месте дтп.xlsx");
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1];

            int rowCount = worksheet.UsedRange.Rows.Count;
            int colCount = worksheet.UsedRange.Columns.Count;
            object[,] data = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[rowCount, colCount]].Value;

            SD.DataTable dt = new SD.DataTable();
            for (int i = 1; i <= colCount; i++)
            {
                dt.Columns.Add(new DataColumn(data[1, i].ToString()));
            }
            for (int row = 2; row <= rowCount; row++)
            {
                DataRow dr = dt.NewRow();
                for (int col = 1; col <= colCount; col++)
                {
                    dr[col - 1] = data[row, col];
                }
                dt.Rows.Add(dr);
            }

            dataGridView7.DataSource = dt;

            workbook.Close();
            excel.Quit();
        }

        private void button_save7_Click(object sender, EventArgs e)
        {
            // Путь к файлу Excel
            string filePath = "C:\\Users\\Public\\Documents\\Файлы\\Работа на месте дтп.xlsx";

            // Создание экземпляра приложения Excel
            Excel.Application excel = new Excel.Application();

            // Открытие файла Excel
            Excel.Workbook workbook = excel.Workbooks.Open(filePath);

            // Выбор листа для очистки данных
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[1];

            // Очистка данных на листе
            worksheet.UsedRange.ClearContents();

            // Сохранение изменений
            workbook.Save();

            // Определяем столбцы и строки для DataGridView
            int columnsCount = dataGridView7.Columns.Count;
            int rowsCount = dataGridView7.Rows.Count;

            for (int i = 0; i <= rowsCount - 2; i++)
            {
                for (int j = 0; j <= columnsCount - 1; j++)
                {
                    worksheet.Cells[1, j + 1] = dataGridView7.Columns[j].HeaderText.ToString();
                    worksheet.Cells[i + 2, j + 1] = dataGridView7[j, i].Value.ToString();
                }
            }

            workbook.Save();
            workbook.Close();
            // Закрытие приложения Excel
            excel.Quit();
            MessageBox.Show("Изменения сохранены");
        }

        private void button_export7_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Add();
            Excel.Worksheet excelWorksheet = excelWorkbook.ActiveSheet;

            // Заполнение заголовков столбцов
            for (int i = 0; i < dataGridView7.Columns.Count; i++)
            {
                excelWorksheet.Cells[1, i + 1] = dataGridView7.Columns[i].HeaderText;
            }

            // Заполнение ячеек таблицы данными
            for (int i = 0; i < dataGridView7.SelectedRows.Count; i++)
            {
                DataGridViewRow row = dataGridView7.SelectedRows[i];
                for (int j = 0; j < row.Cells.Count; j++)
                {
                    excelWorksheet.Cells[i + 2, j + 1] = row.Cells[j].Value.ToString();
                }
            }
            excelApp.Visible = true;
        }

        private void button_loud8_Click(object sender, EventArgs e)
        {
            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Public\\Documents\\Файлы\\Иная помощь.xlsx");
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1];

            int rowCount = worksheet.UsedRange.Rows.Count;
            int colCount = worksheet.UsedRange.Columns.Count;
            object[,] data = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[rowCount, colCount]].Value;

            SD.DataTable dt = new SD.DataTable();
            for (int i = 1; i <= colCount; i++)
            {
                dt.Columns.Add(new DataColumn(data[1, i].ToString()));
            }
            for (int row = 2; row <= rowCount; row++)
            {
                DataRow dr = dt.NewRow();
                for (int col = 1; col <= colCount; col++)
                {
                    dr[col - 1] = data[row, col];
                }
                dt.Rows.Add(dr);
            }

            dataGridView8.DataSource = dt;

            workbook.Close();
            excel.Quit();
        }

        private void button_save8_Click(object sender, EventArgs e)
        {
            // Путь к файлу Excel
            string filePath = "C:\\Users\\Public\\Documents\\Файлы\\Иная помощь.xlsx";

            // Создание экземпляра приложения Excel
            Excel.Application excel = new Excel.Application();

            // Открытие файла Excel
            Excel.Workbook workbook = excel.Workbooks.Open(filePath);

            // Выбор листа для очистки данных
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[1];

            // Очистка данных на листе
            worksheet.UsedRange.ClearContents();

            // Сохранение изменений
            workbook.Save();

            // Определяем столбцы и строки для DataGridView
            int columnsCount = dataGridView8.Columns.Count;
            int rowsCount = dataGridView8.Rows.Count;

            for (int i = 0; i <= rowsCount - 2; i++)
            {
                for (int j = 0; j <= columnsCount - 1; j++)
                {
                    worksheet.Cells[1, j + 1] = dataGridView8.Columns[j].HeaderText.ToString();
                    worksheet.Cells[i + 2, j + 1] = dataGridView8[j, i].Value.ToString();
                }
            }

            workbook.Save();
            workbook.Close();
            // Закрытие приложения Excel
            excel.Quit();
            MessageBox.Show("Изменения сохранены");
        }

        private void button_export8_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Add();
            Excel.Worksheet excelWorksheet = excelWorkbook.ActiveSheet;

            // Заполнение заголовков столбцов
            for (int i = 0; i < dataGridView8.Columns.Count; i++)
            {
                excelWorksheet.Cells[1, i + 1] = dataGridView8.Columns[i].HeaderText;
            }

            // Заполнение ячеек таблицы данными
            for (int i = 0; i < dataGridView8.SelectedRows.Count; i++)
            {
                DataGridViewRow row = dataGridView8.SelectedRows[i];
                for (int j = 0; j < row.Cells.Count; j++)
                {
                    excelWorksheet.Cells[i + 2, j + 1] = row.Cells[j].Value.ToString();
                }
            }
            excelApp.Visible = true;
        }

        private void button_loud9_Click(object sender, EventArgs e)
        {
            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Public\\Documents\\Файлы\\ДТП с пострадавшими.xlsx");
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1];

            int rowCount = worksheet.UsedRange.Rows.Count;
            int colCount = worksheet.UsedRange.Columns.Count;
            object[,] data = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[rowCount, colCount]].Value;

            SD.DataTable dt = new SD.DataTable();
            for (int i = 1; i <= colCount; i++)
            {
                dt.Columns.Add(new DataColumn(data[1, i].ToString()));
            }
            for (int row = 2; row <= rowCount; row++)
            {
                DataRow dr = dt.NewRow();
                for (int col = 1; col <= colCount; col++)
                {
                    dr[col - 1] = data[row, col];
                }
                dt.Rows.Add(dr);
            }

            dataGridView9.DataSource = dt;

            workbook.Close();
            excel.Quit();
        }

        private void button_save9_Click(object sender, EventArgs e)
        {
            // Путь к файлу Excel
            string filePath = "C:\\Users\\Public\\Documents\\Файлы\\ДТП с пострадавшими.xlsx";

            // Создание экземпляра приложения Excel
            Excel.Application excel = new Excel.Application();

            // Открытие файла Excel
            Excel.Workbook workbook = excel.Workbooks.Open(filePath);

            // Выбор листа для очистки данных
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[1];

            // Очистка данных на листе
            worksheet.UsedRange.ClearContents();

            // Сохранение изменений
            workbook.Save();

            // Определяем столбцы и строки для DataGridView
            int columnsCount = dataGridView9.Columns.Count;
            int rowsCount = dataGridView9.Rows.Count;

            for (int i = 0; i <= rowsCount - 2; i++)
            {
                for (int j = 0; j <= columnsCount - 1; j++)
                {
                    worksheet.Cells[1, j + 1] = dataGridView9.Columns[j].HeaderText.ToString();
                    worksheet.Cells[i + 2, j + 1] = dataGridView9[j, i].Value.ToString();
                }
            }

            workbook.Save();
            workbook.Close();
            // Закрытие приложения Excel
            excel.Quit();
            MessageBox.Show("Изменения сохранены");
        }

        private void button_export9_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Add();
            Excel.Worksheet excelWorksheet = excelWorkbook.ActiveSheet;

            // Заполнение заголовков столбцов
            for (int i = 0; i < dataGridView9.Columns.Count; i++)
            {
                excelWorksheet.Cells[1, i + 1] = dataGridView9.Columns[i].HeaderText;
            }

            // Заполнение ячеек таблицы данными
            for (int i = 0; i < dataGridView9.SelectedRows.Count; i++)
            {
                DataGridViewRow row = dataGridView9.SelectedRows[i];
                for (int j = 0; j < row.Cells.Count; j++)
                {
                    excelWorksheet.Cells[i + 2, j + 1] = row.Cells[j].Value.ToString();
                }
            }
            excelApp.Visible = true;
        }

        private void button_loud10_Click(object sender, EventArgs e)
        {
            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Public\\Documents\\Файлы\\ДТП без пострадавших.xlsx");
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1];

            int rowCount = worksheet.UsedRange.Rows.Count;
            int colCount = worksheet.UsedRange.Columns.Count;
            object[,] data = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[rowCount, colCount]].Value;

            SD.DataTable dt = new SD.DataTable();
            for (int i = 1; i <= colCount; i++)
            {
                dt.Columns.Add(new DataColumn(data[1, i].ToString()));
            }
            for (int row = 2; row <= rowCount; row++)
            {
                DataRow dr = dt.NewRow();
                for (int col = 1; col <= colCount; col++)
                {
                    dr[col - 1] = data[row, col];
                }
                dt.Rows.Add(dr);
            }

            dataGridView10.DataSource = dt;

            workbook.Close();
            excel.Quit();
        }

        private void button_save10_Click(object sender, EventArgs e)
        {
            // Путь к файлу Excel
            string filePath = "C:\\Users\\Public\\Documents\\Файлы\\ДТП без пострадавших.xlsx";

            // Создание экземпляра приложения Excel
            Excel.Application excel = new Excel.Application();

            // Открытие файла Excel
            Excel.Workbook workbook = excel.Workbooks.Open(filePath);

            // Выбор листа для очистки данных
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[1];

            // Очистка данных на листе
            worksheet.UsedRange.ClearContents();

            // Сохранение изменений
            workbook.Save();

            // Определяем столбцы и строки для DataGridView
            int columnsCount = dataGridView10.Columns.Count;
            int rowsCount = dataGridView10.Rows.Count;

            for (int i = 0; i <= rowsCount - 2; i++)
            {
                for (int j = 0; j <= columnsCount - 1; j++)
                {
                    worksheet.Cells[1, j + 1] = dataGridView10.Columns[j].HeaderText.ToString();
                    worksheet.Cells[i + 2, j + 1] = dataGridView10[j, i].Value.ToString();
                }
            }

            workbook.Save();
            workbook.Close();
            // Закрытие приложения Excel
            excel.Quit();
            MessageBox.Show("Изменения сохранены");
        }

        private void button_export10_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Add();
            Excel.Worksheet excelWorksheet = excelWorkbook.ActiveSheet;

            // Заполнение заголовков столбцов
            for (int i = 0; i < dataGridView10.Columns.Count; i++)
            {
                excelWorksheet.Cells[1, i + 1] = dataGridView10.Columns[i].HeaderText;
            }

            // Заполнение ячеек таблицы данными
            for (int i = 0; i < dataGridView10.SelectedRows.Count; i++)
            {
                DataGridViewRow row = dataGridView10.SelectedRows[i];
                for (int j = 0; j < row.Cells.Count; j++)
                {
                    excelWorksheet.Cells[i + 2, j + 1] = row.Cells[j].Value.ToString();
                }
            }
            excelApp.Visible = true;
        }

        private void button_loud11_Click(object sender, EventArgs e)
        {
            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Public\\Documents\\Файлы\\ДТП с участием пешеходов.xlsx");
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1];

            int rowCount = worksheet.UsedRange.Rows.Count;
            int colCount = worksheet.UsedRange.Columns.Count;
            object[,] data = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[rowCount, colCount]].Value;

            SD.DataTable dt = new SD.DataTable();
            for (int i = 1; i <= colCount; i++)
            {
                dt.Columns.Add(new DataColumn(data[1, i].ToString()));
            }
            for (int row = 2; row <= rowCount; row++)
            {
                DataRow dr = dt.NewRow();
                for (int col = 1; col <= colCount; col++)
                {
                    dr[col - 1] = data[row, col];
                }
                dt.Rows.Add(dr);
            }

            dataGridView11.DataSource = dt;

            workbook.Close();
            excel.Quit();
        }

        private void button_save11_Click(object sender, EventArgs e)
        {
            // Путь к файлу Excel
            string filePath = "C:\\Users\\Public\\Documents\\Файлы\\ДТП с участием пешеходов.xlsx";

            // Создание экземпляра приложения Excel
            Excel.Application excel = new Excel.Application();

            // Открытие файла Excel
            Excel.Workbook workbook = excel.Workbooks.Open(filePath);

            // Выбор листа для очистки данных
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[1];

            // Очистка данных на листе
            worksheet.UsedRange.ClearContents();

            // Сохранение изменений
            workbook.Save();

            // Определяем столбцы и строки для DataGridView
            int columnsCount = dataGridView11.Columns.Count;
            int rowsCount = dataGridView11.Rows.Count;

            for (int i = 0; i <= rowsCount - 2; i++)
            {
                for (int j = 0; j <= columnsCount - 1; j++)
                {
                    worksheet.Cells[1, j + 1] = dataGridView11.Columns[j].HeaderText.ToString();
                    worksheet.Cells[i + 2, j + 1] = dataGridView11[j, i].Value.ToString();
                }
            }

            workbook.Save();
            workbook.Close();
            // Закрытие приложения Excel
            excel.Quit();
            MessageBox.Show("Изменения сохранены");
        }

        private void button_export11_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Add();
            Excel.Worksheet excelWorksheet = excelWorkbook.ActiveSheet;

            // Заполнение заголовков столбцов
            for (int i = 0; i < dataGridView11.Columns.Count; i++)
            {
                excelWorksheet.Cells[1, i + 1] = dataGridView11.Columns[i].HeaderText;
            }

            // Заполнение ячеек таблицы данными
            for (int i = 0; i < dataGridView11.SelectedRows.Count; i++)
            {
                DataGridViewRow row = dataGridView11.SelectedRows[i];
                for (int j = 0; j < row.Cells.Count; j++)
                {
                    excelWorksheet.Cells[i + 2, j + 1] = row.Cells[j].Value.ToString();
                }
            }
            excelApp.Visible = true;
        }

      /*  private void button_update1_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\\Users\\Public\\Documents\\Файлы\\Коэффициент реагирования.xlsx");
            Excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1];

            SD.DataTable dt = new SD.DataTable();
            for (int i = 1; i <= xlWorksheet.UsedRange.Columns.Count; i++)
            {
                string columnName = xlWorksheet.Cells[1, i].Value.ToString();
                comboBox1.Items.Add(columnName);
                comboBox2.Items.Add(columnName);
                comboBox3.Items.Add(columnName);
            }
            xlWorkbook.Close(false);
            xlApp.Quit();
        } */

      /*  private void button_loud12_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(@"C:\\Users\\Public\\Documents\\Файлы\\Коэффициент реагирования.xlsx");
            Excel._Worksheet worksheet = workbook.Sheets[1];
            Excel.Range range = worksheet.UsedRange;

            // Очистка DataGridView перед загрузкой данных
            dataGridView12.Rows.Clear();
            dataGridView12.Columns.Clear();

            // Добавление выбранных столбцов в DataGridView
            DataGridViewTextBoxColumn column1 = new DataGridViewTextBoxColumn();
            column1.HeaderText = comboBox1.SelectedItem.ToString();
            dataGridView12.Columns.Add(column1);

            DataGridViewTextBoxColumn column2 = new DataGridViewTextBoxColumn();
            column2.HeaderText = comboBox2.SelectedItem.ToString();
            dataGridView12.Columns.Add(column2);

            DataGridViewTextBoxColumn column3 = new DataGridViewTextBoxColumn();
            column3.HeaderText = comboBox3.SelectedItem.ToString();
            dataGridView12.Columns.Add(column3);

            // Заполнение DataGridView данными из файла Excel
            for (int row = 2; row <= range.Rows.Count; row++)
            {
                dataGridView12.Rows.Add(
                    range.Cells[row, comboBox1.SelectedIndex + 1].Value.ToString(),
                    range.Cells[row, comboBox2.SelectedIndex + 1].Value.ToString(),
                    range.Cells[row, comboBox3.SelectedIndex + 1].Value.ToString());
            }

            // Закрытие файла Excel
            workbook.Close(false);
            excelApp.Quit();
        } */

        private void result1_Click(object sender, EventArgs e)
        {
            DataGridViewTextBoxColumn newColumn = new DataGridViewTextBoxColumn();
            newColumn.HeaderText = "Динамика";
            dataGridView12.Columns.Add(newColumn);

            for (int i = 0; i < dataGridView12.Rows.Count; i++)
            {
                double value1;
                double value2;

                if (double.TryParse(dataGridView12.Rows[i].Cells[1].Value.ToString(), out value1) &&
                    double.TryParse(dataGridView12.Rows[i].Cells[2].Value.ToString(), out value2))
                {
                    double result = value2 - value1;
                    dataGridView12.Rows[i].Cells[3].Value = result.ToString();
                }
                else
                {
                    // Обработка ошибок, если ввод был некорректным.
                    MessageBox.Show("Введите корректные значения в ячейки.");
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {

            // Получить выбранный файл Excel из ComboBox
            string selectedFile = Path.Combine("C:\\Users\\Public\\Documents\\Файлы", comboBox7.SelectedItem.ToString());

            // Создать объект Excel.Application
            Excel.Application excelApp = new Excel.Application();

            // Открыть выбранный файл Excel
            Excel.Workbook workbook = excelApp.Workbooks.Open(selectedFile);

            // Получить первый лист в файле Excel
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1];

            SD.DataTable dt = new SD.DataTable();
            for (int i = 1; i <= worksheet.UsedRange.Columns.Count; i++)
            {
                string columnName = worksheet.Cells[1, i].Value.ToString();
                comboBox4.Items.Add(columnName);
                comboBox5.Items.Add(columnName);
                comboBox6.Items.Add(columnName);
            }
            workbook.Close(false);
            excelApp.Quit();
        }

        private void button_loud13_Click(object sender, EventArgs e)
        {
            string selectedFile = Path.Combine("C:\\Users\\Public\\Documents\\Файлы", comboBox7.SelectedItem.ToString());

            // Создать объект Excel.Application
            Excel.Application excelApp = new Excel.Application();

            // Открыть выбранный файл Excel
            Excel.Workbook workbook = excelApp.Workbooks.Open(selectedFile);

            // Получить первый лист в файле Excel
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1];
            Excel.Range range = worksheet.UsedRange;

            // Очистка DataGridView перед загрузкой данных
            dataGridView12.Rows.Clear();
            dataGridView12.Columns.Clear();

            
            // Добавление выбранных столбцов в DataGridView
            DataGridViewTextBoxColumn column1 = new DataGridViewTextBoxColumn();
            column1.HeaderText = comboBox4.SelectedItem.ToString();
            dataGridView12.Columns.Add(column1);

            DataGridViewTextBoxColumn column2 = new DataGridViewTextBoxColumn();
            column2.HeaderText = comboBox5.SelectedItem.ToString();
            dataGridView12.Columns.Add(column2);

            DataGridViewTextBoxColumn column3 = new DataGridViewTextBoxColumn();
            column3.HeaderText = comboBox6.SelectedItem.ToString();
            dataGridView12.Columns.Add(column3);

            
                // Заполнение DataGridView данными из файла Excel
                for (int row = 2; row <= range.Rows.Count; row++)
                {
                    dataGridView12.Rows.Add(
                        range.Cells[row, comboBox4.SelectedIndex + 1].Value.ToString(),
                        range.Cells[row, comboBox5.SelectedIndex + 1].Value.ToString(),
                        range.Cells[row, comboBox6.SelectedIndex + 1].Value.ToString());
                }
            
            // Закрытие файла Excel
            workbook.Close(false);
            excelApp.Quit();

            DataGridViewTextBoxColumn newColumn = new DataGridViewTextBoxColumn();
            newColumn.HeaderText = "Динамика";
            dataGridView12.Columns.Add(newColumn);
            DataGridViewTextBoxColumn newColumn1 = new DataGridViewTextBoxColumn();
            newColumn1.HeaderText = "АППГ";
            dataGridView12.Columns.Add(newColumn1);
        }

       
        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Add();
            Excel.Worksheet excelWorksheet = excelWorkbook.ActiveSheet;

            // Заполнение заголовков столбцов
            for (int i = 0; i < dataGridView12.Columns.Count; i++)
            {
                excelWorksheet.Cells[1, i + 1] = dataGridView12.Columns[i].HeaderText;
            }

            // Заполнение ячеек таблицы данными
            for (int i = 0; i < dataGridView12.SelectedRows.Count; i++)
            {
                DataGridViewRow row = dataGridView12.SelectedRows[i];
                for (int j = 0; j < row.Cells.Count; j++)
                {
                    excelWorksheet.Cells[i + 2, j + 1] = row.Cells[j].Value.ToString();
                }
            }
            excelApp.Visible = true;
        }

        private void button_filter1_Click(object sender, EventArgs e)
        {
            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Public\\Documents\\Файлы\\Кол-во ДТП.xlsx");
            Excel.Worksheet worksheet;

            string selectedSheet = comboBoxfilter1.SelectedItem.ToString();
            worksheet = workbook.Sheets[selectedSheet];
            Excel.Range range = worksheet.UsedRange;

            // Загрузка данных из выбранного листа в DataTable
            SD.DataTable dt = new SD.DataTable();
            for (int i = 1; i <= range.Columns.Count; i++)
            {
                dt.Columns.Add(Convert.ToString(range.Cells[1, i].Value));
            }
            for (int row = 2; row <= range.Rows.Count; row++)
            {
                DataRow dr = dt.NewRow();
                for (int col = 1; col <= range.Columns.Count; col++)
                {
                    dr[col - 1] = range.Cells[row, col].Value;
                }
                dt.Rows.Add(dr);
            }

            // Отображение данных в DataGridView
            dataGridView1.DataSource = dt;
            workbook.Close();
            excel.Quit();
        }

        private void buttonClear1_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();
        }

        private void button_filter2_Click(object sender, EventArgs e)
        {
            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Public\\Documents\\Файлы\\Кол-во выездов ПСП.xlsx");
            Excel.Worksheet worksheet;

            string selectedSheet = comboBoxfilter2.SelectedItem.ToString();
            worksheet = workbook.Sheets[selectedSheet];
            Excel.Range range = worksheet.UsedRange;

            // Загрузка данных из выбранного листа в DataTable
            SD.DataTable dt = new SD.DataTable();
            for (int i = 1; i <= range.Columns.Count; i++)
            {
                dt.Columns.Add(Convert.ToString(range.Cells[1, i].Value));
            }
            for (int row = 2; row <= range.Rows.Count; row++)
            {
                DataRow dr = dt.NewRow();
                for (int col = 1; col <= range.Columns.Count; col++)
                {
                    dr[col - 1] = range.Cells[row, col].Value;
                }
                dt.Rows.Add(dr);
            }

            // Отображение данных в DataGridView
            dataGridView2.DataSource = dt;
            workbook.Close();
            excel.Quit();
        }

        private void buttonCleare2_Click(object sender, EventArgs e)
        {
            dataGridView2.DataSource = null;
            dataGridView2.Rows.Clear();
            dataGridView2.Columns.Clear();
        }

        private void button_filter3_Click(object sender, EventArgs e)
        {
            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Public\\Documents\\Файлы\\Кол-во граж кот оказана пом.xlsx");
            Excel.Worksheet worksheet;

            string selectedSheet = comboBoxfilter3.SelectedItem.ToString();
            worksheet = workbook.Sheets[selectedSheet];
            Excel.Range range = worksheet.UsedRange;

            // Загрузка данных из выбранного листа в DataTable
            SD.DataTable dt = new SD.DataTable();
            for (int i = 1; i <= range.Columns.Count; i++)
            {
                dt.Columns.Add(Convert.ToString(range.Cells[1, i].Value));
            }
            for (int row = 2; row <= range.Rows.Count; row++)
            {
                DataRow dr = dt.NewRow();
                for (int col = 1; col <= range.Columns.Count; col++)
                {
                    dr[col - 1] = range.Cells[row, col].Value;
                }
                dt.Rows.Add(dr);
            }

            // Отображение данных в DataGridView
            dataGridView3.DataSource = dt;
            workbook.Close();
            excel.Quit();
        }

        private void buttonClear3_Click(object sender, EventArgs e)
        {
            dataGridView3.DataSource = null;
            dataGridView3.Rows.Clear();
            dataGridView3.Columns.Clear();
        }

        private void button_filter4_Click(object sender, EventArgs e)
        {
            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Public\\Documents\\Файлы\\Количество деблокированных.xlsx");
            Excel.Worksheet worksheet;

            string selectedSheet = comboBoxfilter4.SelectedItem.ToString();
            worksheet = workbook.Sheets[selectedSheet];
            Excel.Range range = worksheet.UsedRange;

            // Загрузка данных из выбранного листа в DataTable
            SD.DataTable dt = new SD.DataTable();
            for (int i = 1; i <= range.Columns.Count; i++)
            {
                dt.Columns.Add(Convert.ToString(range.Cells[1, i].Value));
            }
            for (int row = 2; row <= range.Rows.Count; row++)
            {
                DataRow dr = dt.NewRow();
                for (int col = 1; col <= range.Columns.Count; col++)
                {
                    dr[col - 1] = range.Cells[row, col].Value;
                }
                dt.Rows.Add(dr);
            }

            // Отображение данных в DataGridView
            dataGridView4.DataSource = dt;
            workbook.Close();
            excel.Quit();
        }

        private void buttonClear4_Click(object sender, EventArgs e)
        {
            dataGridView4.DataSource = null;
            dataGridView4.Rows.Clear();
            dataGridView4.Columns.Clear();
        }

        private void button_filter5_Click(object sender, EventArgs e)
        {
            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Public\\Documents\\Файлы\\Среднее время прибытия.xlsx");
            Excel.Worksheet worksheet;

            string selectedSheet = comboBoxfilter5.SelectedItem.ToString();
            worksheet = workbook.Sheets[selectedSheet];
            Excel.Range range = worksheet.UsedRange;

            // Загрузка данных из выбранного листа в DataTable
            SD.DataTable dt = new SD.DataTable();
            for (int i = 1; i <= range.Columns.Count; i++)
            {
                dt.Columns.Add(Convert.ToString(range.Cells[1, i].Value));
            }
            for (int row = 2; row <= range.Rows.Count; row++)
            {
                DataRow dr = dt.NewRow();
                for (int col = 1; col <= range.Columns.Count; col++)
                {
                    dr[col - 1] = range.Cells[row, col].Value;
                }
                dt.Rows.Add(dr);
            }

            // Отображение данных в DataGridView
            dataGridView5.DataSource = dt;
            workbook.Close();
            excel.Quit();
        }

        private void buttonClear5_Click(object sender, EventArgs e)
        {
            dataGridView5.DataSource = null;
            dataGridView5.Rows.Clear();
            dataGridView5.Columns.Clear();
        }

        private void button_filter6_Click(object sender, EventArgs e)
        {
            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Public\\Documents\\Файлы\\Коэффициент реагирования.xlsx");
            Excel.Worksheet worksheet;

            string selectedSheet = comboBoxfilter6.SelectedItem.ToString();
            worksheet = workbook.Sheets[selectedSheet];
            Excel.Range range = worksheet.UsedRange;

            // Загрузка данных из выбранного листа в DataTable
            SD.DataTable dt = new SD.DataTable();
            for (int i = 1; i <= range.Columns.Count; i++)
            {
                dt.Columns.Add(Convert.ToString(range.Cells[1, i].Value));
            }
            for (int row = 2; row <= range.Rows.Count; row++)
            {
                DataRow dr = dt.NewRow();
                for (int col = 1; col <= range.Columns.Count; col++)
                {
                    dr[col - 1] = range.Cells[row, col].Value;
                }
                dt.Rows.Add(dr);
            }

            // Отображение данных в DataGridView
            dataGridView6.DataSource = dt;
            workbook.Close();
            excel.Quit();
        }

        private void buttonClear6_Click(object sender, EventArgs e)
        {
            dataGridView6.DataSource = null;
            dataGridView6.Rows.Clear();
            dataGridView6.Columns.Clear();
        }

        private void button_filter7_Click(object sender, EventArgs e)
        {
            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Public\\Documents\\Файлы\\Работа на месте дтп.xlsx");
            Excel.Worksheet worksheet;

            string selectedSheet = comboBoxfilter7.SelectedItem.ToString();
            worksheet = workbook.Sheets[selectedSheet];
            Excel.Range range = worksheet.UsedRange;

            // Загрузка данных из выбранного листа в DataTable
            SD.DataTable dt = new SD.DataTable();
            for (int i = 1; i <= range.Columns.Count; i++)
            {
                dt.Columns.Add(Convert.ToString(range.Cells[1, i].Value));
            }
            for (int row = 2; row <= range.Rows.Count; row++)
            {
                DataRow dr = dt.NewRow();
                for (int col = 1; col <= range.Columns.Count; col++)
                {
                    dr[col - 1] = range.Cells[row, col].Value;
                }
                dt.Rows.Add(dr);
            }

            // Отображение данных в DataGridView
            dataGridView7.DataSource = dt;
            workbook.Close();
            excel.Quit();
        }

        private void buttonClear7_Click(object sender, EventArgs e)
        {
            dataGridView7.DataSource = null;
            dataGridView7.Rows.Clear();
            dataGridView7.Columns.Clear();
        }

        private void button_filter8_Click(object sender, EventArgs e)
        {
            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Public\\Documents\\Файлы\\Иная помощь.xlsx");
            Excel.Worksheet worksheet;

            string selectedSheet = comboBoxfilter8.SelectedItem.ToString();
            worksheet = workbook.Sheets[selectedSheet];
            Excel.Range range = worksheet.UsedRange;

            // Загрузка данных из выбранного листа в DataTable
            SD.DataTable dt = new SD.DataTable();
            for (int i = 1; i <= range.Columns.Count; i++)
            {
                dt.Columns.Add(Convert.ToString(range.Cells[1, i].Value));
            }
            for (int row = 2; row <= range.Rows.Count; row++)
            {
                DataRow dr = dt.NewRow();
                for (int col = 1; col <= range.Columns.Count; col++)
                {
                    dr[col - 1] = range.Cells[row, col].Value;
                }
                dt.Rows.Add(dr);
            }

            // Отображение данных в DataGridView
            dataGridView8.DataSource = dt;
            workbook.Close();
            excel.Quit();
        }

        private void buttonClear8_Click(object sender, EventArgs e)
        {
            dataGridView8.DataSource = null;
            dataGridView8.Rows.Clear();
            dataGridView8.Columns.Clear();
        }

        private void button_filter9_Click(object sender, EventArgs e)
        {
            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Public\\Documents\\Файлы\\ДТП с пострадавшими.xlsx");
            Excel.Worksheet worksheet;

            string selectedSheet = comboBoxfilter9 .SelectedItem.ToString();
            worksheet = workbook.Sheets[selectedSheet];
            Excel.Range range = worksheet.UsedRange;

            // Загрузка данных из выбранного листа в DataTable
            SD.DataTable dt = new SD.DataTable();
            for (int i = 1; i <= range.Columns.Count; i++)
            {
                dt.Columns.Add(Convert.ToString(range.Cells[1, i].Value));
            }
            for (int row = 2; row <= range.Rows.Count; row++)
            {
                DataRow dr = dt.NewRow();
                for (int col = 1; col <= range.Columns.Count; col++)
                {
                    dr[col - 1] = range.Cells[row, col].Value;
                }
                dt.Rows.Add(dr);
            }

            // Отображение данных в DataGridView
            dataGridView9.DataSource = dt;
            workbook.Close();
            excel.Quit();
        }

        private void buttonClear9_Click(object sender, EventArgs e)
        {
            dataGridView9.DataSource = null;
            dataGridView9.Rows.Clear();
            dataGridView9.Columns.Clear();
        }

        private void button_filter10_Click(object sender, EventArgs e)
        {
            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Public\\Documents\\Файлы\\ДТП без пострадавших.xlsx");
            Excel.Worksheet worksheet;

            string selectedSheet = comboBoxfilter10.SelectedItem.ToString();
            worksheet = workbook.Sheets[selectedSheet];
            Excel.Range range = worksheet.UsedRange;

            // Загрузка данных из выбранного листа в DataTable
            SD.DataTable dt = new SD.DataTable();
            for (int i = 1; i <= range.Columns.Count; i++)
            {
                dt.Columns.Add(Convert.ToString(range.Cells[1, i].Value));
            }
            for (int row = 2; row <= range.Rows.Count; row++)
            {
                DataRow dr = dt.NewRow();
                for (int col = 1; col <= range.Columns.Count; col++)
                {
                    dr[col - 1] = range.Cells[row, col].Value;
                }
                dt.Rows.Add(dr);
            }

            // Отображение данных в DataGridView
            dataGridView10.DataSource = dt;
            workbook.Close();
            excel.Quit();
        }

        private void buttonClear10_Click(object sender, EventArgs e)
        {
            dataGridView10.DataSource = null;
            dataGridView10.Rows.Clear();
            dataGridView10.Columns.Clear();
        }

        private void button_filter11_Click(object sender, EventArgs e)
        {
            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Public\\Documents\\Файлы\\ДТП с участием пешеходов.xlsx");
            Excel.Worksheet worksheet;

            string selectedSheet = comboBoxfilter11.SelectedItem.ToString();
            worksheet = workbook.Sheets[selectedSheet];
            Excel.Range range = worksheet.UsedRange;

            // Загрузка данных из выбранного листа в DataTable
            SD.DataTable dt = new SD.DataTable();
            for (int i = 1; i <= range.Columns.Count; i++)
            {
                dt.Columns.Add(Convert.ToString(range.Cells[1, i].Value));
            }
            for (int row = 2; row <= range.Rows.Count; row++)
            {
                DataRow dr = dt.NewRow();
                for (int col = 1; col <= range.Columns.Count; col++)
                {
                    dr[col - 1] = range.Cells[row, col].Value;
                }
                dt.Rows.Add(dr);
            }

            // Отображение данных в DataGridView
            dataGridView11.DataSource = dt;
            workbook.Close();
            excel.Quit();
        }

        private void buttonClear11_Click(object sender, EventArgs e)
        {
            dataGridView11.DataSource = null;
            dataGridView11.Rows.Clear();
            dataGridView11.Columns.Clear();
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void result2_Click(object sender, EventArgs e)
        {

            /* DataGridViewTextBoxColumn newColumn = new DataGridViewTextBoxColumn();
             newColumn.HeaderText = "АППГ";
             dataGridView12.Columns.Add(newColumn);

             for (int i = 0; i < dataGridView12.Rows.Count; i++)
             {
                 double value1;
                 double value2;

                 if (double.TryParse(dataGridView12.Rows[i].Cells[1].Value.ToString(), out value1) &&
                     double.TryParse(dataGridView12.Rows[i].Cells[2].Value.ToString(), out value2))
                 {
                     double result = value2 - value1;
                     dataGridView12.Rows[i].Cells[3].Value = result.ToString();
                 }
                 else
                 {
                     // Обработка ошибок, если ввод был некорректным.
                     MessageBox.Show("Введите корректные значения в ячейки.");
                 }
             } */




            // Проверяем выбранный пункт
            /*  if (checkedListBox1.GetItemChecked(0)) // Первый пункт выбран
              {
                  DataGridViewTextBoxColumn newColumn = new DataGridViewTextBoxColumn();
                  newColumn.HeaderText = "Динамика";
                  dataGridView12.Columns.Add(newColumn);
                  // Вычитание столбцов
                  for (int i = 0; i < dataGridView12.Rows.Count; i++)
                  {
                      double value1;
                      double value2;

                      if (double.TryParse(dataGridView12.Rows[i].Cells[1].Value.ToString(), out value1) &&
                          double.TryParse(dataGridView12.Rows[i].Cells[2].Value.ToString(), out value2))
                      {
                          double result = value2 - value1;
                          dataGridView12.Rows[i].Cells[3].Value = result.ToString();
                      }
                      else
                      {
                          // Обработка ошибок, если ввод был некорректным.
                          MessageBox.Show("Введите корректные значения в ячейки.");
                      }
                  }
              }
              else if (checkedListBox1.GetItemChecked(1)) // Второй пункт выбран
              {
                  DataGridViewTextBoxColumn newColumn = new DataGridViewTextBoxColumn();
                  newColumn.HeaderText = "АППГ";
                  dataGridView12.Columns.Add(newColumn);
                  // Деление, умножение и вычитание столбцов
                  for (int i = 0; i < dataGridView12.Rows.Count; i++)
                  {
                      double value1;
                      double value2;

                      if (double.TryParse(dataGridView12.Rows[i].Cells[1].Value.ToString(), out value1) &&
                          double.TryParse(dataGridView12.Rows[i].Cells[2].Value.ToString(), out value2))
                      {
                          double result = (value2 / value1) * 100 - 100;
                          dataGridView12.Rows[i].Cells[4].Value = result.ToString();
                      }
                      else
                      {
                          // Обработка ошибок, если ввод был некорректным.
                          MessageBox.Show("Введите корректные значения в ячейки.");
                      }
                  }
              }  */

            bool subtractSelected = false;
            bool divideSelected = false;

            // Проверяем выбранные пункты
            foreach (var item in checkedListBox1.CheckedItems)
            {
                if (item.ToString() == "Динамика")
                {
                    subtractSelected = true;
                }
                else if (item.ToString() == "АППГ")
                {
                    divideSelected = true;
                }
            }
           
            // Выполняем соответствующие вычисления для выбранных пунктов
            for (int i = 0; i < dataGridView12.Rows.Count; i++)
            {
                double value1;
                double value2;

                if (double.TryParse(dataGridView12.Rows[i].Cells[1].Value.ToString(), out value1) &&
                    double.TryParse(dataGridView12.Rows[i].Cells[2].Value.ToString(), out value2))
                {
                    if (subtractSelected)
                    {
                        // Вычитание столбцов
                        double result = value2 - value1;
                        dataGridView12.Rows[i].Cells[3].Value = result.ToString();
                    }

                    if (divideSelected)
                    {
                        // Деление, умножение и вычитание столбцов
                        double result = (value2 / value1) * 100 - 100;
                        dataGridView12.Rows[i].Cells[4].Value = result.ToString("0.00");
                    }
                }
                else
                {
                    // Обработка ошибок, если ввод был некорректным.
                    MessageBox.Show("Введите корректные значения в ячейки.");
                }
            }

        }

        private void buttonClear12_Click(object sender, EventArgs e)
        {
            comboBox4.Items.Clear();
            comboBox5.Items.Clear();
            comboBox6.Items.Clear();
            dataGridView12.DataSource = null;
            dataGridView12.Rows.Clear();
            dataGridView12.Columns.Clear();
        }

       
    }
}
