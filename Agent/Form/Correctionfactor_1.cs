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
    public partial class Correctionfactor : UserControl
    {
        datebase db = new datebase();
        public SqlConnection sqlConnection = null;
        public Correctionfactor()
        {
            InitializeComponent();
            sqlConnection = new SqlConnection(db.connection);
        }
        int id = 0;int id2 = 0;
        public void Correctionfactor_load()
        {
            sqlConnection.Open();
            DataSet dataSet = new DataSet();
            SqlDataAdapter command = new SqlDataAdapter($@"Select  idcorrectionfactor,correctionfactor.name as [Название коэффициента],coefficient as [Корректировочный коэффициент],correctionfactor.note as Примечание,correctionfactor.idvida,vid.name as [Вид страхования] from correctionfactor inner join vid on vid.idvida=correctionfactor.idvida", sqlConnection);
            command.Fill(dataSet);
            dataGridView1.DataSource = dataSet.Tables[0].DefaultView;
            sqlConnection.Close();
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[4].Visible = false;
            dataGridView1.AllowUserToAddRows = false;
            sqlConnection.Close();
        }
        public void clear()
        {
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            comboBox1.SelectedIndex = -1;
        }
        public void vid_load()
        {
            sqlConnection.Open();
            DataSet dataSet = new DataSet();
            SqlDataAdapter command = new SqlDataAdapter($@"Select idvida,vid.name as [Вид страхования] from vid", sqlConnection);
            command.Fill(dataSet);
            dataGridView2.DataSource = dataSet.Tables[0].DefaultView;
            sqlConnection.Close();
            dataGridView2.Columns[0].Visible = false;
            dataGridView2.AllowUserToAddRows = false;
            sqlConnection.Close();
        }
        private void Correctionfactor_Load(object sender, EventArgs e)
        {
            dataGridView2.RowsDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView1.RowsDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            clear();
            panel2.Visible = false;
            Correctionfactor_load();
            vid_load();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (panel2.Visible == false)
            {
                clear();
                panel2.Visible = true;
                label6.Text = "Добавить корректировочный коэффициент";
                button11.Text = "Добавить";
                button11.Width = 174;
                button11.Left = 137;
                Correctionfactor_load();
                dataGridView1.Enabled = true;
                button11.Image = new Bitmap(@"D:\College\4kurs\Praktica4Kurs\proga\Agent\Agent\Resources\pngwing.com (31).png");
            }
            else
            {
                clear();
                panel2.Visible = false;
            }
            }
        int znak = 0;
        int countt = 0;
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (panel2.Visible == false)
                {
                    clear();
                    if (id != 0)
                    {
                        dataGridView1.Enabled = false;
                        panel2.Visible = true;
                        label6.Text = "Редактировать корректировочный коэффициент";
                        textBox2.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                        textBox3.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                        char[] o = { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' };
                        znak = dataGridView1.CurrentRow.Cells[3].Value.ToString().LastIndexOfAny(o);
                        countt = dataGridView1.CurrentRow.Cells[3].Value.ToString().IndexOfAny(o);
                       
                        if (countt == 2)
                        {    textBox4.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString().Substring(countt, znak - countt + 1);
                            comboBox1.SelectedItem = dataGridView1.CurrentRow.Cells[3].Value.ToString().Substring(0, 2);
                        }
                        else if (countt == 1)
                        {
                            textBox4.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString().Substring(countt, znak - countt + 1);
                            comboBox1.SelectedItem = dataGridView1.CurrentRow.Cells[3].Value.ToString().Substring(0, 1);
                        }
                        else
                        {
                            textBox4.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                        }

                        for (int i = 0; i < dataGridView2.RowCount; i++)
                        {
                            dataGridView2.Rows[i].Selected = false;
                        }
                            for (int i = 0; i < dataGridView2.RowCount; i++)
                        {
                            
                            if (dataGridView1.CurrentRow.Cells[4].Value.ToString() == dataGridView2[0, i].Value.ToString())
                            {
                                dataGridView2.Rows[i].Selected = true;
                                dataGridView2.Rows[i].DefaultCellStyle.SelectionForeColor = Color.Black;
                                dataGridView2.Rows[i].DefaultCellStyle.SelectionBackColor = Color.FromArgb(212, 236, 252);
                                break;
                            }
                        }

                        button11.Image = new Bitmap(@"D:\College\4kurs\Praktica4Kurs\proga\Agent\Agent\Resources\pngwing.com (29).png");
                        button11.Text = "Редактировать";
                        button11.Width = 207;
                        button11.Left = 121;
                    }
                    else
                    {
                        MessageBox.Show("Строка не выбрана!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        panel2.Visible = false;
                    }
                }
                else
                {
                    clear();
                    panel2.Visible = false;
                    dataGridView1.Enabled = true;
                }
            }
            catch { }
        }

        private void button2_Click(object sender, EventArgs e)
        {
             clear();
            dataGridView1.Enabled = true;
            panel2.Visible = false;
            if (id != 0)
            {
                try
                {
                    if (MessageBox.Show($@"Вы уверены что хотите удалить корректировочный коэффициент {dataGridView1.CurrentRow.Cells[1].Value.ToString()}?", "Удаление", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {

                        sqlConnection.Open();
                        string query = $@"DELETE FROM [correctionfactor] WHERE [idcorrectionfactor] = 
                             {dataGridView1.CurrentRow.Cells[0].Value.ToString()}";
                        SqlCommand command = new SqlCommand(query, sqlConnection);
                        command.ExecuteNonQuery();
                        sqlConnection.Close();
                        Correctionfactor_load();
                    }
                }
                catch { MessageBox.Show("Удаление невозможно, данные все еще нужны!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            }
            else { MessageBox.Show("Выберите строку для удаления!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }

        }
        int visible = 0;
        private void button5_Click(object sender, EventArgs e)
        {
            dataGridView1.Enabled = true;
            clear();
            panel2.Visible = false;
            visible = 0;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (dataGridView1.Rows[i].Visible == true)
                {
                    visible++;
                }
            }
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Excel.Worksheet ExcelWorkSheet;
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
            ExcelApp.Columns.NumberFormat = "General";
            ExcelWorkSheet.StandardWidth = 30;
            ExcelWorkSheet.Columns.ColumnWidth = 20;
            ExcelApp.Rows[1].Columns[2] = "Корректировоные коэффициенты";
            ExcelApp.Rows[visible + 3].Columns[2] = "Ридецкая Анна Михайловна";
            for (int i = 1; i < dataGridView1.Columns.Count; i++)
            {
                ExcelApp.Cells[2, i] = dataGridView1.Columns[i].HeaderText;

            }
            for (int j = 1; j < dataGridView1.ColumnCount; j++)
            {
                for (int i = 0; i < visible; i++)
                {
                    if (dataGridView1.Rows[i].Cells[j].Visible == true)
                    {
                        ExcelApp.Cells[i + 3, j] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                    }

                }

            }
            for (int i = 0; i < visible; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    Microsoft.Office.Interop.Excel.Range range = ExcelWorkSheet.Range[$"A1:C{visible + 3}"];
                    ExcelWorkSheet.Range[$"A1:C{visible + 3}"].Cells.WrapText = true;
                    range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders.Weight = Excel.XlBorderWeight.xlMedium;
                }
            }
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
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
        int k = 0;
        int j = 0;
        Decimal cof = 0;
        private void button11_Click(object sender, EventArgs e)
        {
            try
            {

                if (label6.Text == "Добавить корректировочный коэффициент")
                {
                    if (textBox2.Text != "" && textBox3.Text != "" && textBox4.Text!="")
                    {

                        k = 0;
                        for (int i = 0; i < dataGridView1.RowCount; i++)
                        {
                            if (textBox2.Text + Convert.ToString(cof) + textBox4.Text == dataGridView1[1, i].Value.ToString() + dataGridView1[2, i].Value.ToString() + dataGridView1[3, i].Value.ToString())
                            {
                                k++;
                            }
                        }
                        if (k == 0)
                        {
                           
                            sqlConnection.Open();
                            SqlCommand command = new SqlCommand($@"INSERT INTO [correctionfactor](name,coefficient,note,idvida) VALUES (@n,@c,@o,@i);", sqlConnection);
                            command.Parameters.AddWithValue("@n", (textBox2.Text));
                            command.Parameters.AddWithValue("@i", (id2));
                            command.Parameters.AddWithValue("@c", Convert.ToDecimal(textBox3.Text));
                            command.Parameters.AddWithValue("@o", (comboBox1.Text+textBox4.Text));
                            command.ExecuteNonQuery();
                            sqlConnection.Close();
                            Correctionfactor_load();
                            clear();
                            panel2.Visible = false;
                            dataGridView1.Enabled = true;
                            id2 = 0;
                        }
                        else
                        {
                            MessageBox.Show("Такой корректировочный коэффициент уже есть!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Заполните поля!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    if (textBox2.Text != "" && textBox3.Text != "" && textBox4.Text != "")
                    {

                        k = 0;
                        for (int i = 0; i < dataGridView1.RowCount; i++)
                        {
                            if (textBox2.Text + Convert.ToString(cof) + textBox4.Text == dataGridView1[1, i].Value.ToString() + dataGridView1[2, i].Value.ToString() + dataGridView1[3, i].Value.ToString())
                            {
                                k++;
                                j = Convert.ToInt32(dataGridView1[0, i].Value.ToString());
                            }
                        }
                        if (k == 0 || j == id)
                        {
                            
                            sqlConnection.Open();
                            SqlCommand command = new SqlCommand($@"UPDATE correctionfactor SET name=@n,coefficient=@c ," +
                            $"note=@o, idvida=@i WHERE idcorrectionfactor=@id", sqlConnection);
                            command.Parameters.AddWithValue("@n", (textBox2.Text));
                            command.Parameters.AddWithValue("@c", Convert.ToDecimal(textBox3.Text));
                            command.Parameters.AddWithValue("@o", (comboBox1.Text+textBox4.Text));
                            command.Parameters.AddWithValue("@i", (id2));
                            command.Parameters.AddWithValue("@id", (id));
                            command.ExecuteNonQuery();
                            sqlConnection.Close();
                            dataGridView1.Enabled = true;
                            clear();
                            Correctionfactor_load();
                            panel2.Visible = false;
                            id2 = 0;id = 0;
                        }
                        else
                        {
                            MessageBox.Show("Такой корректировочный коэффициент уже есть!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Заполните поля!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
        }
            catch { }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            id = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
            id2 = Convert.ToInt32(dataGridView1.CurrentRow.Cells[4].Value.ToString());
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            try
            {
                cof = Convert.ToDecimal(textBox3.Text) + Convert.ToDecimal(0.01) - +Convert.ToDecimal(0.01);
            }
            catch { }
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && e.KeyChar != 8 && e.KeyChar != '.' && e.KeyChar != ',')
                e.Handled = true;
            if (e.KeyChar == '.')
                e.KeyChar = ',';
            if (e.KeyChar == ',')
            {
                if (((sender as TextBox).Text.IndexOf(',') != -1) || (sender as TextBox).Text.Length == 0)
                {
                    e.Handled = true;
                    return;
                }
            }
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            e.Handled = !((c >= 'а' && c <= 'я') || (c >= 'А' && c <= 'Я') || (c == 'ё' || c == 8 || c == 32 || c == ',' || c == '.' || c == '0' || c == '1' || c == '2' || c == '3' || c == '4' || c == '5' || c == '6' || c == '7' || c == '8' || c == '9'));

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.Length == 1)
                ((TextBox)sender).Text = ((TextBox)sender).Text.ToUpper();
            ((TextBox)sender).Select(((TextBox)sender).Text.Length, 0);
        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (comboBox1.Text!="")
            {
                char c = e.KeyChar;
                e.Handled = !((c == 8 || c == ',' || c == '.' || c == '0' || c == '1' || c == '2' || c == '3' || c == '4' || c == '5' || c == '6' || c == '7' || c == '8' || c == '9'));
            }
            else
            {
                char c = e.KeyChar;
                e.Handled = !((c >= 'а' && c <= 'я') || (c >= 'А' && c <= 'Я') || (c == 'ё' || c == 8 || c == 32 || c == ',' || c == '.' || c == '0' || c == '1' || c == '2' || c == '3' || c == '4' || c == '5' || c == '6' || c == '7' || c == '8' || c == '9'));
            }
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            id2 = Convert.ToInt32(dataGridView2.CurrentRow.Cells[0].Value.ToString());
        }
    }
}
