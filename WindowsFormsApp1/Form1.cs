using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        IXLWorksheet ws1;
        IXLWorksheet ws2;
        IXLWorksheet ws3;
        IEnumerable<IXLRangeRow> rows1;
        IEnumerable<IXLRangeRow> rows2;
        IEnumerable<IXLRangeRow> rows3;
        XLWorkbook workbook;
        int rowN;
        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() != DialogResult.OK)
            {
                MessageBox.Show("Вы не выбрали файл для открытия", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            textBox1.Text += $"{openFileDialog1.FileName}\r\n";
            workbook = new XLWorkbook(openFileDialog1.FileName); //открыли файл
            ws1 = workbook.Worksheet(1);
            ws2 = workbook.Worksheet(2);
            ws3 = workbook.Worksheet(3); //получили 3 листа
            rows1 = ws1.RangeUsed().RowsUsed().Skip(1);
            rows2 = ws2.RangeUsed().RowsUsed().Skip(1);
            rows3 = ws3.RangeUsed().RowsUsed().Skip(1); //получили 3 коллецккии строк в этих листах
            foreach ( var row in rows1) 
            {
                if (!row.Cell(2).IsEmpty()) //чтобы не забирать пустые строки
                    comboBox1.Items.Add(row.Cell(2).GetValue<string>());   //Забиваем товар для 2 части задания
            }
            foreach (var row in rows2)
            {
                if (!row.Cell(2).IsEmpty())
                    comboBox2.Items.Add(row.Cell(2).GetValue<string>());  //забиваем фио для 3 части задания
            }
            comboBox1.Enabled = true;
            comboBox2.Enabled = true;
            comboBox3.Enabled = true;
            comboBox4.Enabled = true;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear(); //ощичаем таблицу
            string value = comboBox1.Text;
            string KodTovara;
            foreach (var row in rows1)
            {
                if (value == row.Cell(2).GetValue<string>()) //ищем совпадения в таблице товары
                {
                    KodTovara = row.Cell(1).GetValue<string>(); //получаем код товара
                    foreach (var row3 in rows3)
                    {
                        if (KodTovara == row3.Cell(2).GetValue<string>()) //ищем совпадения в таблицее заявки
                        {
                            string clientName = "";
                            foreach (var row2 in rows2)
                                if (row3.Cell(3).GetValue<string>() == row2.Cell(1).GetValue<string>()) //по коду клиента ищем имя клиента в таблице клиенты
                                    clientName = row2.Cell(2).GetValue<string>();
                            dataGridView1.Rows.Add(clientName, row3.Cell(5).GetValue<string>(), row.Cell(4).GetValue<string>(), row3.Cell(6).GetValue<string>()); //заполняем полученными данными таблицу
                        }
                    }
                }
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            string ContactFace = "";
            foreach (var row in rows2)
            {

                if (comboBox2.Text == row.Cell(2).GetValue<string>())
                {
                    ContactFace = row.Cell(4).GetValue<string>();
                    rowN = row.RowNumber();
                }
            }
            label2.Text = $"Контактное лицо: {ContactFace}";
            textBox3.Enabled = true;
            button2.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ws2.Row(rowN).Cell(4).Value = $"{textBox3.Text}";
            workbook.Save();
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox2.Text = "";
            List<string> result = new List<string>();
            for (int i = 2; i <= rows3.Count(); i++)
            {
                DateTime a = ws3.Row(i).Cell(6).Value;
                if (a.Year == Convert.ToInt32(comboBox4.Text))
                    result.Add(ws3.Row(i).Cell(3).GetString());
            }
            if (result.Count > 0)
            {
                string end = result.GroupBy(x => x).OrderByDescending(g => g.Count()).ThenBy(g => g.Key, StringComparer.Ordinal).First().Key;
                foreach (var s in rows2)
                    if (end == s.Cell(1).GetValue<string>())
                        textBox2.Text = $"{s.Cell(2).GetValue<string>()}";
            }
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox2.Text = "";
            List<string> result = new List<string>();
            for (int i = 2; i <= rows3.Count(); i++)
            {
                DateTime a = ws3.Row(i).Cell(6).Value;
                if (a.Month == Convert.ToInt32(comboBox3.Text))
                    result.Add(ws3.Row(i).Cell(3).GetString());
            }
            if (result.Count>0)
            {
                string end = result.GroupBy(x => x).OrderByDescending(g => g.Count()).ThenBy(g => g.Key, StringComparer.Ordinal).First().Key;
                foreach (var s in rows2)
                    if (end == s.Cell(1).GetValue<string>())
                        textBox2.Text = $"{s.Cell(2).GetValue<string>()}";
            }
        }
    }
}

