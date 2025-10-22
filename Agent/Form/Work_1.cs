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
    public partial class Work : UserControl
    {
        datebase db = new datebase();
        public SqlConnection sqlConnection = null;
        public Work()
        {
            InitializeComponent();
            sqlConnection = new SqlConnection(db.connection);
        }
        int id = 0;
        int id2 = 0;
        public void Work_load()
        {
            sqlConnection.Open();
            DataSet dataSet = new DataSet();
            SqlDataAdapter command = new SqlDataAdapter($@"Select idwork, name as [Место работы] from work", sqlConnection);
            command.Fill(dataSet);
            dataGridView1.DataSource = dataSet.Tables[0].DefaultView;
            sqlConnection.Close();
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.AllowUserToAddRows = false;
            sqlConnection.Close();
        }
        public void clear()
        {
            comboBox2.SelectedIndex = -1;
            textBox3.Text = "";
            textBox4.Text = "";
            comboBox1.Text = "";
            comboBox1.SelectedIndex = -1;
        }
        public void comboBoxvid()
        {
            sqlConnection.Close();
            sqlConnection.Open();
            string query = "select idwork,name from work";
            SqlDataAdapter sqlDbDataAdapter = new SqlDataAdapter(query, sqlConnection);
            DataSet dataSet = new DataSet();
            sqlDbDataAdapter.Fill(dataSet);
            comboBox2.DataSource = dataSet.Tables[0];
            comboBox2.DisplayMember = "name";
            comboBox2.ValueMember = "idwork";
            comboBox2.SelectedIndex = -1;
            sqlConnection.Close();

        }
        public void comboBoxposition()
        {
            sqlConnection.Close();
            sqlConnection.Open();
            string query = "Select Distinct name from position";
            SqlDataAdapter sqlDbDataAdapter = new SqlDataAdapter(query, sqlConnection);
            DataSet dataSet = new DataSet();
            sqlDbDataAdapter.Fill(dataSet);
            comboBox1.DataSource = dataSet.Tables[0];
            comboBox1.DisplayMember = "name";
            comboBox1.ValueMember = "name";
            comboBox1.SelectedIndex = -1;
            sqlConnection.Close();

        }
        private void Objectinsurance_Load(object sender, EventArgs e)
        {
            dataGridView1.RowsDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView2.RowsDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            Work_load();
            comboBoxvid();
            comboBoxposition();
            clear();
            panel2.Visible = false;
            panel3.Visible = false;
            panel8.Visible = false;
      
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (panel3.Visible == false&&panel2.Visible==true)
            {
                dataGridView1.Enabled = true;
                dataGridView2.Enabled = true;
                panel2.Visible = false;
                panel3.Visible = false;
                panel8.Visible = false;
            }
            else if (panel2.Visible==true)
            {
                //clear();
                dataGridView1.Enabled = true;
                dataGridView2.Enabled = true;
                panel2.Visible = false;
                panel3.Visible = false;
                panel8.Visible = false;
            }
            else if (panel8.Visible == true)
            {
                //clear();
                dataGridView1.Enabled = true; 
                dataGridView2.Enabled=true;
                panel2.Visible = false;
                panel3.Visible = false;
                panel8.Visible = false;
            }
            else
            {
                panel3.Visible = true;
                button3.Image = new Bitmap(@"D:\College\4kurs\Praktica4Kurs\proga\Agent\Agent\Resources\pngwing.com (31).png");
                button3.Width = 313;
                button3.Left = 102;
                button6.Image = new Bitmap(@"D:\College\4kurs\Praktica4Kurs\proga\Agent\Agent\Resources\pngwing.com (31).png");
                button6.Width = 313;
                button6.Left = 102;
                button3.Text = "Добавить работу";
                button6.Text = "Добавить должность";
                panel3.BringToFront();
                dataGridView1.Enabled = false;
                dataGridView2.Enabled = false;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (panel3.Visible == false && panel2.Visible == true)
            {
                dataGridView1.Enabled = true;
                dataGridView2.Enabled = true;
                panel2.Visible = false;
                panel3.Visible = false;
                panel8.Visible = false;
            }
            else if (panel2.Visible == true)
            {
                //clear();
                dataGridView1.Enabled = true;
                dataGridView2.Enabled = true;
                panel2.Visible = false;
                panel3.Visible = false;
                panel8.Visible = false;
            }
            else if (panel8.Visible == true)
            {
                //clear();
                dataGridView1.Enabled = true;
                dataGridView2.Enabled = true;
                panel2.Visible = false;
                panel3.Visible = false;
                panel8.Visible = false;
            }
            else
            {
                panel3.Visible = true;
                button3.Image = new Bitmap(@"D:\College\4kurs\Praktica4Kurs\proga\Agent\Agent\Resources\pngwing.com (29).png");
                button3.Width = 377;
                button3.Left = 70;
                button6.Image = new Bitmap(@"D:\College\4kurs\Praktica4Kurs\proga\Agent\Agent\Resources\pngwing.com (29).png");
                button6.Width = 377;
                button6.Left = 70;
                button3.Text = "Редактировать работу";
                button6.Text = "Редактировать должность";
                panel3.BringToFront();
                dataGridView1.Enabled = false;
                dataGridView2.Enabled = false;
            }
          
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (panel3.Visible == false && panel2.Visible == true)
            {
                dataGridView1.Enabled = true;
                dataGridView2.Enabled = true;
                panel2.Visible = false;
                panel3.Visible = false;
                panel8.Visible = false;
            }
            else if (panel2.Visible == true)
            {
                //clear();
                dataGridView1.Enabled = true;
                dataGridView2.Enabled = true;
                panel2.Visible = false;
                panel3.Visible = false;
                panel8.Visible = false;
            }
            else if (panel8.Visible == true)
            {
                //clear();
                dataGridView1.Enabled = true;
                dataGridView2.Enabled = true;
                panel2.Visible = false;
                panel3.Visible = false;
                panel8.Visible = false;
            }
            else
            {
                panel3.Visible = true;
                button3.Image = new Bitmap(@"D:\Diplom\proga\Agent\Agent\Resources\pngwing.com (30).png");
                button3.Width = 298;
                button3.Left = 109;
                button6.Image = new Bitmap(@"D:\Diplom\proga\Agent\Agent\Resources\pngwing.com (30).png");
                button6.Width = 298;
                button6.Left = 109;
                button3.Text = "Удалить работу";
                button6.Text = "Удалить должность";
                panel3.BringToFront();
                dataGridView1.Enabled = false;
                dataGridView2.Enabled = false;
            }
            
        }
        int visible = 0;
        private void button5_Click(object sender, EventArgs e)
        {
            if (panel3.Visible == false && panel2.Visible == true)
            {
                dataGridView1.Enabled = true;
                dataGridView2.Enabled = true;
                panel2.Visible = false;
                panel3.Visible = false;
                panel8.Visible = false;
            }
            else if (panel2.Visible == true)
            {
                //clear();
                dataGridView1.Enabled = true;
                dataGridView2.Enabled = true;
                panel2.Visible = false;
                panel3.Visible = false;
                panel8.Visible = false;
            }
            else if (panel8.Visible == true)
            {
                //clear();
                dataGridView1.Enabled = true;
                dataGridView2.Enabled = true;
                panel2.Visible = false;
                panel3.Visible = false;
                panel8.Visible = false;
            }
            else
            {
                panel3.Visible = true;
                button3.Image = new Bitmap(@"D:\Diplom\proga\Agent\Agent\Resources\pngwing.com (33).png");
                button3.Width = 217;
                button3.Left = 150;
                button6.Image = new Bitmap(@"D:\Diplom\proga\Agent\Agent\Resources\pngwing.com (33).png");
                button6.Width = 396;
                button6.Left = 60;
                button3.Text = "Вывод работ";
                button6.Text = "Вывод работ и должностей";
                panel3.BringToFront();
                dataGridView1.Enabled = false;
                dataGridView2.Enabled = false;
            }
           
        }
        int k = 0;
        int j = 0;
        
        private void button11_Click(object sender, EventArgs e)
        {
            try
            {
                if (label6.Text == "Добавить место работы")
                {
                    if (textBox3.Text != "")
                    {
                        k = 0;
                        for (int i = 0; i < dataGridView1.RowCount; i++)
                        {
                            if (textBox3.Text == dataGridView1[1, i].Value.ToString())
                            {
                                k++;
                            }
                        }
                        if (k == 0)
                        {
                            sqlConnection.Open();
                            SqlCommand command = new SqlCommand($@"INSERT INTO [work](name) VALUES (@n);", sqlConnection);
                            command.Parameters.AddWithValue("@n", (textBox3.Text));
                            command.ExecuteNonQuery();
                            sqlConnection.Close();
                            Work_load();
                            clear();
                            panel2.Visible = false;
                        }
                        else
                        {
                            MessageBox.Show("Такое место работы уже есть!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Заполните поле!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    if (textBox3.Text != "")
                    {

                        k = 0;
                        for (int i = 0; i < dataGridView1.RowCount; i++)
                        {
                            if (textBox3.Text == dataGridView1[1, i].Value.ToString())
                            {
                                k++;
                                j = Convert.ToInt32(dataGridView1[0, i].Value.ToString());
                            }
                        }
                        if (k == 0 || j == id)
                        {
                            sqlConnection.Open();
                            SqlCommand command = new SqlCommand($@"UPDATE work SET name=@n WHERE idwork=@id", sqlConnection);
                            command.Parameters.AddWithValue("@n", (textBox3.Text));
                            command.Parameters.AddWithValue("@id", (id));
                            command.ExecuteNonQuery();
                            sqlConnection.Close();
                            dataGridView1.Enabled = true;
                            clear();
                            Work_load();
                            panel2.Visible = false;
                            id = 0;
                        }
                        else
                        {
                            MessageBox.Show("Такое место работы уже есть!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Заполните поле!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            catch { }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            for(int i = 0; i < dataGridView1.RowCount; i++)
            {
                dataGridView1.Rows[i].Selected = false;
                for (int j = 1; j < dataGridView1.ColumnCount; j++)
                {
                    if (j != 3)
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
            for (int i = 0; i < dataGridView2.RowCount; i++)
            {
                dataGridView2.Rows[i].Selected = false;
            for (int j = 1; j < dataGridView2.ColumnCount; j++)
            {
                if (j != 3)
                {
                    if (dataGridView2.Rows[i].Cells[j].Value != null)
                    {
                        if (dataGridView2.Rows[i].Cells[j].Value.ToString().ToLower().Contains(textBox1.Text.ToLower()) && textBox1.Text != "")
                        {
                            dataGridView2.Rows[i].Selected = true;
                            dataGridView2.Rows[i].DefaultCellStyle.SelectionForeColor = Color.Black;
                            dataGridView2.Rows[i].DefaultCellStyle.SelectionBackColor = Color.FromArgb(212, 236, 252);
                            break;
                        }
                        else
                        {
                            dataGridView2.Rows[i].Selected = false;
                            dataGridView2.Rows[i].DefaultCellStyle.BackColor = Color.White;


                        }
                    }
                }
            }
        }
    }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.Length == 1)
                ((TextBox)sender).Text = ((TextBox)sender).Text.ToUpper();
            ((TextBox)sender).Select(((TextBox)sender).Text.Length, 0);
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            e.Handled = !((c >= 'а' && c <= 'я') || (c >= 'А' && c <= 'Я') || (c == 'ё' || c == 8 || c == 32));
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            e.Handled = !((c >= 'а' && c <= 'я') || (c >= 'А' && c <= 'Я') || (c == 'ё' || c == 8 || c == 32 || c == ',' || c == '.' || c == '0' || c == '1' || c == '2' || c == '3' || c == '4' || c == '5' || c == '6' || c == '7' || c == '8' || c == '9'));
        }
        public void Position_load()
        { 
         sqlConnection.Open();
                DataSet dataSet = new DataSet();
                SqlDataAdapter command = new SqlDataAdapter($@"Select idposition, position.name as Должность ,harmhul as [Процент вредности],position.idwork from position inner join work on position.idwork=work.idwork where position.idwork={id}", sqlConnection);
                command.Fill(dataSet);
                dataGridView2.DataSource = dataSet.Tables[0].DefaultView;
                sqlConnection.Close();
                dataGridView2.Columns[0].Visible = false;
            dataGridView2.Columns[3].Visible = false;
            dataGridView2.AllowUserToAddRows = false;
                sqlConnection.Close();
        }
            private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                Position_load();
              
                  id = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
            }catch { }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.Length == 1)
                ((TextBox)sender).Text = ((TextBox)sender).Text.ToUpper();
            ((TextBox)sender).Select(((TextBox)sender).Text.Length, 0);
        }

        private void textBox2_TextChanged_1(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.Length == 1)
                ((TextBox)sender).Text = ((TextBox)sender).Text.ToUpper();
            ((TextBox)sender).Select(((TextBox)sender).Text.Length, 0);
        }

        private void textBox2_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            e.Handled = !((c >= 'а' && c <= 'я') || (c >= 'А' && c <= 'Я') || (c == 'ё' || c == 8 || c == 32 || c == ',' || c == '.' || c == '0' || c == '1' || c == '2' || c == '3' || c == '4' || c == '5' || c == '6' || c == '7' || c == '8' || c == '9'));
        }

        private void panel3_VisibleChanged(object sender, EventArgs e)
        {
            if (panel3.Visible == true)
            {
                panel1.Enabled = false;
                dataGridView1.Enabled= false;
                dataGridView2.Enabled= false;
            }
            else
            {
                panel1.Enabled = true;
                //dataGridView1.Enabled = true;
                //dataGridView2.Enabled = true;
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            panel1.Enabled = true;
            dataGridView1.Enabled = true;
            dataGridView2.Enabled = true;
            panel2.Visible = false;
            panel3.Visible = false;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (button3.Text == "Добавить работу")
            {
                clear();
                label6.Text = "Добавить место работы";
                button11.Text = "Добавить";
                button11.Width = 174;
                button11.Left = 128;
                //Work_load();
                //dataGridView1.Enabled = true;
                //dataGridView2.Enabled = true;
                button11.Image = new Bitmap(@"D:\College\4kurs\Praktica4Kurs\proga\Agent\Agent\Resources\pngwing.com (31).png");
                panel2.Visible = true;
                panel3.Visible = false;
                panel8.Visible = false;
                panel1.Enabled = true;
            }
            else if (button3.Text == "Удалить работу")
            {
                clear();
                //dataGridView1.Enabled = true;
                //panel2.Visible = false;
                if (id != 0)
                {
                    try
                    {
                        if (MessageBox.Show($@"Вы уверены что хотите удалить место работы {dataGridView1.CurrentRow.Cells[1].Value.ToString()}?", "Удаление", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            sqlConnection.Open();
                            string query = $@"DELETE FROM [work] WHERE [idwork] = 
                                 {dataGridView1.CurrentRow.Cells[0].Value.ToString()}";
                            SqlCommand command = new SqlCommand(query, sqlConnection);
                            command.ExecuteNonQuery();
                            sqlConnection.Close();
                            Work_load();
                            panel3.Visible = false;
                            dataGridView1.Enabled = true;
                            dataGridView2.Enabled = true;
                        }
                    }
                    catch { MessageBox.Show("Удаление невозможно, данные все еще нужны!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                }
                else { MessageBox.Show("Выберите строку для удаления!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            }
            else if (button3.Text == "Вывод работ")
            {
                dataGridView1.Enabled = true;
                dataGridView2.Enabled = true;
                clear();
                panel3.Visible = false;
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
                ExcelApp.Rows[1].Columns[1] = "Место работы";
                ExcelApp.Rows[visible + 3].Columns[1] = "Ридецкая Анна Михайловна";
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
                        Microsoft.Office.Interop.Excel.Range range = ExcelWorkSheet.Range[$"A1:A{visible + 3}"];
                        ExcelWorkSheet.Range[$"A1:A{visible + 3}"].Cells.WrapText = true;
                        range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders.Weight = Excel.XlBorderWeight.xlMedium;
                    }
                }
                ExcelApp.Visible = true;
                ExcelApp.UserControl = true;
            }
            else
            {
                if (id != 0)
                {
                    clear();
                    label6.Text = "Редактировать место работы";
                    button11.Text = "Редактировать";
                    button11.Width = 245;
                    button11.Left = 94;
                    //  Work_load();
                    dataGridView1.Enabled = false;
                    dataGridView2.Enabled = false;
                    button11.Image = new Bitmap(@"D:\College\4kurs\Praktica4Kurs\proga\Agent\Agent\Resources\pngwing.com (29).png");
                    panel2.Visible = true;
                    panel3.Visible = false;
                    panel8.Visible = false;
                    panel1.Enabled = true;
                    textBox3.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                }
                else
                {
                    MessageBox.Show("Строка не выбрана!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void button6_Click(object sender, EventArgs e)
        { 
            if (button6.Text == "Добавить должность")
            {
                clear();
                label9.Text = "Добавить должность";
                button8.Text = "Добавить";
                button8.Width = 174;
                button8.Left = 164;
                //dataGridView1.Enabled = true;
                //dataGridView2.Enabled = true;
                button8.Image = new Bitmap(@"D:\College\4kurs\Praktica4Kurs\proga\Agent\Agent\Resources\pngwing.com (31).png");
                panel8.Visible = true;
                panel3.Visible = false;
                panel2.Visible = false;
                panel1.Enabled = true;
            }
            else if (button6.Text == "Удалить должность")
            {
                clear();
                //dataGridView1.Enabled = true;
                //panel2.Visible = false;
                if (id2 != 0)
                {
                    try
                    {
                        if (MessageBox.Show($@"Вы уверены что хотите удалить должность {dataGridView2.CurrentRow.Cells[1].Value.ToString()}?", "Удаление", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            sqlConnection.Open();
                            string query = $@"DELETE FROM [position] WHERE [idposition] = 
                                 {dataGridView2.CurrentRow.Cells[0].Value.ToString()}";
                            SqlCommand command = new SqlCommand(query, sqlConnection);
                            command.ExecuteNonQuery();
                            sqlConnection.Close();
                            Position_load();
                            panel3.Visible = false;
                            dataGridView1.Enabled = true;
                            dataGridView2.Enabled = true;
                        }
                    }
                    catch { MessageBox.Show("Удаление невозможно, данные все еще нужны!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                }
                else { MessageBox.Show("Выберите строку для удаления!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            }
            else if (button6.Text == "Вывод работ и должностей")
            {

                dataGridView1.Enabled = true;
                dataGridView2.Enabled = true;
                clear();
                panel3.Visible = false;
                visible = 0;
                string query31 = $@"Select Count(idposition) from position inner join work on position.idwork=work.idwork";
                DataTable data31 = new DataTable();
                SqlDataAdapter command31 = new SqlDataAdapter(query31, sqlConnection);
                command31.Fill(data31);
                DataColumn column22 = data31.Columns[0];
                DataRow row22 = data31.Rows[0];

                for (int i = 0; i < Convert.ToInt32(row22[column22].ToString()); i++)
                {visible++;}

                    

                Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
                Excel.Worksheet ExcelWorkSheet;
                ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
                ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
                ExcelApp.Columns.NumberFormat = "General";
                ExcelWorkSheet.StandardWidth = 30;
                ExcelWorkSheet.Columns.ColumnWidth = 20;
                ExcelApp.Rows[1].Columns[2] = "Должности и места работы";
                ExcelApp.Rows[visible + 3].Columns[2] = "Ридецкая Анна Михайловна";
                    ExcelApp.Cells[2, 1] = "Должность";
                    ExcelApp.Cells[2, 2] = "Процент вредности";
                    ExcelApp.Cells[2, 3] = "Место работы";

                for (int j = 1; j < Convert.ToInt32(row22[column22].ToString())+1; j++)
                {
                    string query3 = $@"Select idposition, position.name as Должность ,harmhul as [Процент вредности], work.name from position inner join work on position.idwork=work.idwork order by work.name ";
                    DataTable data3 = new DataTable();
                    SqlDataAdapter command3 = new SqlDataAdapter(query3, sqlConnection);
                    command3.Fill(data3);
                    DataColumn column2 = data3.Columns[1];
                    DataRow row2 = data3.Rows[j-1];
                    string name = (row2[column2].ToString());
                    ExcelApp.Cells[j + 2, 1] = name;
                    DataColumn column222 = data3.Columns[2];
                    DataRow row222 = data3.Rows[j - 1];
                    string proz = (row222[column222].ToString());
                    ExcelApp.Cells[j + 2, 2] = proz;
                    DataColumn column2222 = data3.Columns[3];
                    DataRow row2222 = data3.Rows[j - 1];
                    string work = (row2222[column2222].ToString());
                    ExcelApp.Cells[j + 2, 3] = work;


                }
                for (int i = 0; i < visible; i++)
                {
                    for (int j = 0; j < Convert.ToInt32(row22[column22].ToString()); j++)
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
            else
            {
                if (id2 != 0)
                {
                    clear();
                    label9.Text = "Редактировать должность";
                    button8.Text = "Редактировать";
                    button8.Width = 245;
                    button8.Left = 128;
                    dataGridView1.Enabled = false;
                    dataGridView2.Enabled = false;
                    button8.Image = new Bitmap(@"D:\College\4kurs\Praktica4Kurs\proga\Agent\Agent\Resources\pngwing.com (29).png");
                    panel8.Visible = true;
                    panel3.Visible = false;
                    panel2.Visible = false;
                    panel1.Enabled = true;
                    textBox4.Text = dataGridView2.CurrentRow.Cells[2].Value.ToString();
                    comboBox1.Text = dataGridView2.CurrentRow.Cells[1].Value.ToString();
                    comboBox2.SelectedValue = dataGridView2.CurrentRow.Cells[3].Value.ToString();
                }
                else
                {
                    MessageBox.Show("Строка не выбрана!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && e.KeyChar != 8 && e.KeyChar != '.' && e.KeyChar != ',')
                e.Handled = true;
            if (e.KeyChar == '.')
                e.KeyChar = ',';
            if (e.KeyChar == ',')
            {
                if (((sender as TextBox).Text.IndexOf(',') != -1))
                {
                    e.Handled = true;
                    return;
                }
            }
        }

        private void comboBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            e.Handled = !((c >= 'а' && c <= 'я') || (c >= 'А' && c <= 'Я') || (c == 'ё' || c == 8 || c == 32));
           
        }

        private void button8_Click(object sender, EventArgs e)
        {
            //try
            //{
                if (button6.Text == "Добавить должность")
                {
                    if (textBox4.Text != ""&& comboBox1.Text!=""&& comboBox2.SelectedIndex != -1)
                    {
                        k = 0;
                    string query31 = $@"Select Count(idposition) from position inner join work on position.idwork=work.idwork";
                    DataTable data31 = new DataTable();
                    SqlDataAdapter command31 = new SqlDataAdapter(query31, sqlConnection);
                    command31.Fill(data31);
                    DataColumn column22 = data31.Columns[0];
                    DataRow row22 = data31.Rows[0];

                    for (int i = 0; i < Convert.ToInt32(row22[column22].ToString()); i++)
                        {  
                    string query3 = $@"Select idposition, position.name as Должность ,harmhul as [Процент вредности],position.idwork from position inner join work on position.idwork=work.idwork";
                    DataTable data3 = new DataTable();
                    SqlDataAdapter command3 = new SqlDataAdapter(query3, sqlConnection);
                    command3.Fill(data3);
                    DataColumn column2 = data3.Columns[1];
                    DataRow row2 = data3.Rows[i];
                        string name = (row2[column2].ToString());
                        DataColumn column12 = data3.Columns[3];
                        DataRow row12 = data3.Rows[i];
                        string work = (row12[column12].ToString());
                        if (comboBox1.Text+comboBox2.SelectedValue ==name +work )
                            {
                                k++;
                            }
                        }
                        if (k == 0)
                        {
                            sqlConnection.Open();
                            SqlCommand command = new SqlCommand($@"INSERT INTO [position](name,harmhul,idwork) VALUES (@n,@h,@w);", sqlConnection);
                            command.Parameters.AddWithValue("@n", (comboBox1.Text));
                            command.Parameters.AddWithValue("@h", Convert.ToDecimal(textBox4.Text));
                            command.Parameters.AddWithValue("@w", (comboBox2.SelectedValue));
                            command.ExecuteNonQuery();
                            sqlConnection.Close();
                            Position_load();
                            clear();
                            panel2.Visible = false;
                            panel8.Visible = false;
                            panel3.Visible = false;
                        }
                        else
                        {
                            MessageBox.Show("Такая должнасть уже есть!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Заполните поля!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    if (textBox4.Text != "" && comboBox1.Text!="" && comboBox2.SelectedIndex != -1)
                    {

                        k = 0;
                    string query31 = $@"Select Count(idposition) from position inner join work on position.idwork=work.idwork";
                    DataTable data31 = new DataTable();
                    SqlDataAdapter command31 = new SqlDataAdapter(query31, sqlConnection);
                    command31.Fill(data31);
                    DataColumn column22 = data31.Columns[0];
                    DataRow row22 = data31.Rows[0];

                    for (int i = 0; i < Convert.ToInt32(row22[column22].ToString()); i++)
                    {
                        string query3 = $@"Select idposition, position.name as Должность ,harmhul as [Процент вредности],position.idwork from position inner join work on position.idwork=work.idwork";
                        DataTable data3 = new DataTable();
                        SqlDataAdapter command3 = new SqlDataAdapter(query3, sqlConnection);
                        command3.Fill(data3);
                        DataColumn column2 = data3.Columns[1];
                        DataRow row2 = data3.Rows[i];
                        string name = (row2[column2].ToString());
                        DataColumn column12 = data3.Columns[3];
                        DataRow row12 = data3.Rows[i];
                        string work = (row12[column12].ToString());
                        DataColumn column122 = data3.Columns[0];
                        DataRow row122 = data3.Rows[i];
                        if (comboBox1.Text + comboBox2.SelectedValue == name + work)
                        {
                            k++;
                            j =Convert.ToInt32 (row122[column122].ToString());
                        }
                    }

                        if (k == 0 || j == id)
                        {
                        k = 0;j = 0;
                            sqlConnection.Open();
                            SqlCommand command = new SqlCommand($@"UPDATE position SET name=@n,harmhul=@h,idwork=@w WHERE idposition=@id", sqlConnection);
                        command.Parameters.AddWithValue("@n", (comboBox1.Text));
                        command.Parameters.AddWithValue("@h", Convert.ToDecimal(textBox4.Text));
                        command.Parameters.AddWithValue("@w", (comboBox2.SelectedValue));
                            command.Parameters.AddWithValue("@id", (id2));
                            command.ExecuteNonQuery();
                            sqlConnection.Close();
                            dataGridView1.Enabled = true;
                            clear();
                            Position_load();
                            panel2.Visible = false;
                            panel8.Visible = false;
                            panel3.Visible = false;
                        id2 = 0;
                        }
                        else
                        {
                        k = 0;j = 0;
                            MessageBox.Show("Такая должность уже есть!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                id2 = Convert.ToInt32(dataGridView2.CurrentRow.Cells[0].Value.ToString());
            }
            catch { }
        }

        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
            if (((ComboBox)sender).Text.Length == 1)
                ((ComboBox)sender).Text = ((ComboBox)sender).Text.ToUpper();
            ((ComboBox)sender).Select(((ComboBox)sender).Text.Length, 0);
        }

        private void panel2_VisibleChanged(object sender, EventArgs e)
        {
            if (panel2.Visible == true)
            {
                dataGridView1.Enabled = false;
                dataGridView2.Enabled = false;
            }
            else
            {
                //dataGridView1.Enabled = true;
                //dataGridView2.Enabled = true;
            }
        }

        private void panel8_VisibleChanged(object sender, EventArgs e)
        {
            if (panel8.Visible == true)
            {
                panel1.Enabled = false;
                dataGridView1.Enabled = false;
                dataGridView2.Enabled = false;
            }
            else
            {
                //dataGridView1.Enabled = true;
                //dataGridView2.Enabled = true;
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try{
                
Position_load();
                id = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
            }catch { }
        }
    }
}
