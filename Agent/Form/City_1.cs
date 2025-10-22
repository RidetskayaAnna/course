using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Agent.Form
{
    public partial class City : UserControl
    {
        datebase db = new datebase();
        public SqlConnection sqlConnection = null;
        public City()
        {
            InitializeComponent();
            sqlConnection = new SqlConnection(db.connection);
        }
        int id = 0;
        public void City_load()
        {
            sqlConnection.Open();
            DataSet dataSet = new DataSet();
            SqlDataAdapter command = new SqlDataAdapter($@"Select idcity, name as [Название города],indexcity as [Идекс],area as Регион from city", sqlConnection);
            command.Fill(dataSet);
            dataGridView1.DataSource = dataSet.Tables[0].DefaultView;
            sqlConnection.Close();
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.AllowUserToAddRows = false;
            sqlConnection.Close();
        }

        private void City_Load(object sender, EventArgs e)
        {
            City_load();
            panel2.Visible = false;
        }

        private void button4_Click(object sender, EventArgs e)
        {

            if (panel2.Visible == false)
            {
            dataGridView1.Enabled = true;
            panel2.Visible = true;
            textBox2.Text = "";
            comboBox1.SelectedIndex = -1;
            label6.Left = 71;
            label6.Text = "Добавить город";
            button11.Text = "Добавить";
            button11.Width = 155;
            button11.Left = 147;
            City_load();
            button11.Image = new Bitmap(@"D:\College\4kurs\Praktica4Kurs\proga\Agent\Agent\Resources\pngwing.com (31).png");
            }
            else
            {
                textBox3.Text = "";
                textBox2.Text = "";
                comboBox1.SelectedIndex = -1;
                panel2.Visible = false;
            }
              }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (panel2.Visible == false)
                {
                    textBox3.Text = "";
                    textBox2.Text = "";
                    comboBox1.SelectedIndex = -1;
                    if (id != 0)
                    {
                        dataGridView1.Enabled = false;
                        panel2.Visible = true;
                        label6.Left = 71;
                        label6.Text = "Редактировать город";
                        textBox3.Text= dataGridView1.CurrentRow.Cells[2].Value.ToString();
                        textBox2.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                        comboBox1.SelectedItem = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                        button11.Image = new Bitmap(@"D:\College\4kurs\Praktica4Kurs\proga\Agent\Agent\Resources\pngwing.com (29).png");
                        button11.Text = "Редактировать";
                        button11.Width = 207;
                        button11.Left = 120;
                    }
                    else
                    {
                        MessageBox.Show("Строка не выбрана!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        panel2.Visible = false;
                    }
                }
                else
                {
                    textBox2.Text = "";
                    comboBox1.SelectedIndex = -1;
                    textBox3.Text = "";
                    panel2.Visible = false;
                }
            }
            catch { }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView1.Enabled = true;
            textBox2.Text = "";
            textBox3.Text = "";
            comboBox1.SelectedIndex = -1;
            panel2.Visible = false;
            if (id != 0)
            {
                try
                {
                    if (MessageBox.Show($@"Вы уверены что хотите удалить город {dataGridView1.CurrentRow.Cells[1].Value.ToString()}?",
                        "Удаление", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {

                        sqlConnection.Open();
                        string query = $@"DELETE FROM [city] WHERE [idcity] = 
                             {dataGridView1.CurrentRow.Cells[0].Value.ToString()}";
                        SqlCommand command = new SqlCommand(query, sqlConnection);
                        command.ExecuteNonQuery();
                        sqlConnection.Close();
                        City_load();
                    }
                }
                catch { MessageBox.Show("Удаление невозможно, данные все еще нужны!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            }
            else { MessageBox.Show("Выберите строку для удаления!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }

        }

        private void button5_Click(object sender, EventArgs e)
        {
            dataGridView1.Enabled = true;
            try
            {
                Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
                Excel.Worksheet ExcelWorkSheet;
                ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
                ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
                ExcelApp.Columns.NumberFormat = "General";
                ExcelWorkSheet.StandardWidth = 30;
                ExcelWorkSheet.Columns.ColumnWidth = 20;
                ExcelApp.Rows[1].Columns[2] = "Город";
                ExcelApp.Rows[dataGridView1.RowCount + 3].Columns[2] = "Ридецкая Анна Михайловна";
                for (int i = 1; i < dataGridView1.Columns.Count; i++)
                {
                    ExcelApp.Cells[2, i] = dataGridView1.Columns[i].HeaderText;

                }
                for (int j = 1; j < dataGridView1.ColumnCount; j++)
                {
                    for (int i = 0; i < dataGridView1.RowCount; i++)
                    {
                        if (dataGridView1.Rows[i].Cells[j].Visible == true)
                        {
                                    ExcelApp.Cells[i + 3, j] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                        }
                    }

                }
                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    for (int j = 0; j < dataGridView1.ColumnCount; j++)
                    {
                        Microsoft.Office.Interop.Excel.Range range = ExcelWorkSheet.Range[$"A1:С{dataGridView1.RowCount + 3}"];
                        ExcelWorkSheet.Range[$"A1:С{dataGridView1.RowCount + 3}"].Cells.WrapText = true;
                        range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders.Weight = Excel.XlBorderWeight.xlMedium;
                    }
                }
                ExcelApp.Visible = true;
                ExcelApp.UserControl = true;
            }
            catch { }
        }
        int k = 0;
        int j = 0;
        private void button11_Click(object sender, EventArgs e)
        {
            //try
            //{
                k = 0;
                j = 0;
                if (label6.Text == "Добавить город")
                {
                    if (textBox2.Text != ""&& comboBox1.SelectedIndex != -1&& textBox3.Text != "")
                    {
                        for (int i = 0; i < dataGridView1.RowCount; i++)
                        {
                            if (textBox2.Text.ToLower() == dataGridView1[1, i].Value.ToString().ToLower())
                            {
                                k++;
                            }
                        }
                        if (k == 0)
                        {    
                                sqlConnection.Open();
                                SqlCommand command = new SqlCommand($@"INSERT INTO [city](name,indexcity,area) VALUES (@n,@in,@a);", sqlConnection);
                                command.Parameters.AddWithValue("@n", (textBox2.Text));
                            command.Parameters.AddWithValue("@in", (textBox3.Text));
                            command.Parameters.AddWithValue("@a", (comboBox1.Text));
                                command.ExecuteNonQuery();
                                sqlConnection.Close();
                                City_load();
                            textBox2.Text = "";
                            textBox3.Text = "";
                            comboBox1.SelectedIndex = -1;
                            panel2.Visible = false;
                            
                        }
                        else
                        {
                            MessageBox.Show("Такой город уже есть!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }

                    }
                    else
                    {
                        MessageBox.Show("Заполните поля!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    if (textBox2.Text != ""&& comboBox1.SelectedIndex != -1&& textBox3.Text!= "")
                    {
                        for (int i = 0; i < dataGridView1.RowCount; i++)
                        {
                            if (textBox2.Text.ToLower() == dataGridView1[1, i].Value.ToString().ToLower())
                            {
                                k++;
                                j = Convert.ToInt32(dataGridView1[0, i].Value.ToString());
                            }
                        }
                        if (k == 0 || j == id)
                        {
                            sqlConnection.Open();
                            SqlCommand command = new SqlCommand($@"UPDATE city SET name=@n,indexcity=@in,area=@a WHERE idcity=@id", sqlConnection);
                            command.Parameters.AddWithValue("@n", (textBox2.Text));
                            command.Parameters.AddWithValue("@in", (textBox3.Text));
                            command.Parameters.AddWithValue("@a", (comboBox1.Text));
                            command.Parameters.AddWithValue("@id", (id));
                            command.ExecuteNonQuery();
                            sqlConnection.Close();
                            City_load();
                            textBox3.Text = "";
                            textBox2.Text = "";
                            comboBox1.SelectedIndex = -1;
                            panel2.Visible = false;
                            dataGridView1.Enabled = true;
                        }
                        else
                        {
                            MessageBox.Show("Такой город уже есть!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Заполните поля!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            //}
            //catch { }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            id = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                dataGridView1.Rows[i].Selected = false;
                for (int j = 1; j < dataGridView1.ColumnCount; j++)
                {
                    if (dataGridView1.Rows[i].Cells[j].Value != null)
                    {
                        if (dataGridView1.Rows[i].Cells[j].Value.ToString().ToLower().Contains(textBox1.Text.ToLower()) && textBox1.Text != "")
                        {
                            dataGridView1.Rows[i].Selected = true;
                            dataGridView1.Rows[i].DefaultCellStyle.SelectionForeColor = Color.Black;
                            dataGridView1.Rows[i].DefaultCellStyle.SelectionBackColor = Color.FromArgb(212, 236, 252);
                            break;
                        }
                        else
                        {
                            dataGridView1.Rows[i].Selected = false;
                            dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.White;


                        }
                    }
                }
            }
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            e.Handled = !((c >= 'а' && c <= 'я') || (c >= 'А' && c <= 'Я') || (c == 'ё' || c == 8 || c == 32));
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.Length == 1)
                ((TextBox)sender).Text = ((TextBox)sender).Text.ToUpper();
            ((TextBox)sender).Select(((TextBox)sender).Text.Length, 0);
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && e.KeyChar != 8)
                e.Handled = true;
        }
    }
}
