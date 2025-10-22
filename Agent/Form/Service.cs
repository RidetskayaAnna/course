using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Agent.Form
{
    public partial class Service : UserControl
    {
        datebase db = new datebase();
        public SqlConnection sqlConnection = null;
        
        Tenantak policyholder;
        public Service(Tenantak policyholder1)
        {
            InitializeComponent();
            sqlConnection = new SqlConnection(db.connection);
            policyholder = policyholder1;
        }
        Workerak insurerak;
        public Service(Workerak insurerak1)
        {
            InitializeComponent();
            sqlConnection = new SqlConnection(db.connection);
            insurerak = insurerak1;
        }
        Form1 form1;
        public Service(Form1 form)
        {
            InitializeComponent();
            sqlConnection = new SqlConnection(db.connection);
            form1 = form;
        }
        int id = 0;
        public void train()
        {
            sqlConnection.Open();
            DataSet dataSet = new DataSet();
            SqlDataAdapter command = new SqlDataAdapter($@"Select idservice, service.name as [Название],description as [Описание],cost as [Цена], service.idpost, post.name as [Работник] from service inner join post on service.idpost=post.idpost", sqlConnection);
            command.Fill(dataSet);
            dataGridView1.DataSource = dataSet.Tables[0].DefaultView;
            sqlConnection.Close();
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[4].Visible = false;
            dataGridView1.AllowUserToAddRows = false;
            sqlConnection.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            panel3.Visible = false;
            panel2.Visible = true;
            label6.Text = "Добавить услугу";
            button11.Text = "Добавить";
            button11.Width = 198;
            button11.Left = 247;
            button11.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (31).png");
            dataGridView1.Enabled = true;
            clear();
        }
        public void comboBoxposition()
        {
            try
            {
                sqlConnection.Close();
                sqlConnection.Open();
                string query = $@"Select Distinct idpost, post.name as Должность from post ";
                SqlDataAdapter sqlDbDataAdapter = new SqlDataAdapter(query, sqlConnection);
                DataSet dataSet = new DataSet();
                sqlDbDataAdapter.Fill(dataSet);
                comboBox1.DataSource = dataSet.Tables[0];
                comboBox1.DisplayMember = "Должность";
                comboBox1.ValueMember = "idpost";
                comboBox1.SelectedIndex = -1;
                sqlConnection.Close();
            }
            catch { }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            panel3.Visible = false;
            panel2.Visible = true;
            label6.Text = "Редактировать услугу";
            button11.Width = 272;
            button11.Left = 210;
            button11.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (29).png");
            button11.Text = "Редактировать";
            textBox2.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            textBox3.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            comboBox1.SelectedValue= dataGridView1.CurrentRow.Cells[4].Value.ToString();
            textBox4.Visible = true;
            textBox4.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            dataGridView1.Enabled = false;
        }
        public void clear()
        {
            textBox3.Text = "";
            comboBox1.SelectedIndex = -1;
            textBox4.Text = "";
            textBox2.Text = "";
        }
        private void Train_Load(object sender, EventArgs e)
        {
            dataGridView1.RowsDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            train();
            comboBoxposition();
            panel2.Visible = false;
            panel3.Visible = false;
            if (insurerak != null||policyholder!=null)
            {
                try
                {
                    button4.Visible = false;
                    button1.Visible = false;
                    button5.Visible = false;
                    button2.Visible = false;
                }
                catch { }
            }
            else {
                button4.Visible = true;
                button1.Visible = true;
                button5.Visible = true;
                button2.Visible = true;
            }
        }
        int k = 0; int j = 0;
        private void button11_Click(object sender, EventArgs e)
        {
            //try
            //{
                k = 0;
                j = 0;
                if (label6.Text == "Добавить услугу")
                {
                    if (textBox2.Text != "" && textBox4.Text != "" && textBox3.Text != ""&& comboBox1.SelectedIndex != -1  )
                    {
                        for (int i = 0; i < dataGridView1.RowCount; i++)
                        {
                            if (textBox4.Text.ToLower()== dataGridView1[1, i].Value.ToString().ToLower())
                            {
                                k++;
                            }
                        }
                                if (k == 0)
                                {
                                    k = 0;
                                    sqlConnection.Open();
                                    SqlCommand command = new SqlCommand($@"INSERT INTO [service](name,description,cost, idpost) VALUES (@n,@d,@c,@idp);", sqlConnection);
                                    command.Parameters.AddWithValue("@idp", (comboBox1.SelectedValue));
                                    command.Parameters.AddWithValue("@n", (textBox4.Text));
                                    command.Parameters.AddWithValue("@c", Convert.ToDecimal(textBox3.Text));
                                    command.Parameters.AddWithValue("@d", (textBox2.Text));
                                    command.ExecuteNonQuery();
                                    sqlConnection.Close();
                                    train();
                                    clear();
                                    panel2.Visible = false;
                        }
                        else
                        {
                            MessageBox.Show("Такая услуга уже есть!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }

                    }
                    else
                    {
                        MessageBox.Show("Заполните поля!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                k = 0; j = id;
                    if (textBox2.Text != "" && textBox4.Text!="" && textBox3.Text != "" && comboBox1.SelectedIndex != -1)
                    {
                        for (int i = 0; i < dataGridView1.RowCount; i++)
                        {
                            if (textBox4.Text.ToLower()  == dataGridView1[1, i].Value.ToString().ToLower())
                            {
                                k++; j = Convert.ToInt32(dataGridView1[0, i].Value.ToString());
                            }
                        }
                        if (k == 0 || j == id)
                        {  
                                    k = 0; j = 0;
                                    sqlConnection.Open();
                                    SqlCommand command = new SqlCommand($@"UPDATE service SET name=@n,description=@d,
                             cost=@c, idpost=@idp WHERE idservice=@id", sqlConnection);
                                    command.Parameters.AddWithValue("@idp", (comboBox1.SelectedValue));
                                    command.Parameters.AddWithValue("@n", (textBox4.Text));
                                    command.Parameters.AddWithValue("@d", (textBox2.Text));
                                    command.Parameters.AddWithValue("@c", Convert.ToDecimal(textBox3.Text));
                                    command.Parameters.AddWithValue("@id", (id));
                                    command.ExecuteNonQuery();
                                    sqlConnection.Close();
                                    train();

                                    panel2.Visible = false;
                                    dataGridView1.Enabled = true;

                        }
                        else
                        {
                            MessageBox.Show("Такая услуга уже есть!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        int count = 0;
        
        private void button2_Click(object sender, EventArgs e)
        {
            panel3.Visible = false;
            dataGridView1.Enabled = true;
            clear();
            panel2.Visible = false;
            if (id != 0)
            {
                try
                {
                    //Фрагмент кода кдаления данных о транспорте из БД
                    if (MessageBox.Show($@"Вы уверены что хотите удалить услугу 
                    {dataGridView1.CurrentRow.Cells[1].Value.ToString() }?",
                        "Удаление", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        sqlConnection.Open();
                        string query = $@"DELETE FROM [service] WHERE [idservice] = 
                             {dataGridView1.CurrentRow.Cells[0].Value.ToString()}";
                        SqlCommand command = new SqlCommand(query, sqlConnection);
                        command.ExecuteNonQuery();
                        sqlConnection.Close();
                        train();
                    }
                }
                catch { MessageBox.Show("Удаление невозможно, данные все еще нужны!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            }
            else { MessageBox.Show("Выберите строку для удаления!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }

        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            id = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            e.Handled = !((c >= 'а' && c <= 'я') || (c >= 'А' && c <= 'Я') || (c == 'ё' || c == '-' || c == '.' || c == ',' || c == 8 || c == 32));
        }

        private void comboBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            e.Handled = !((c >= 'а' && c <= 'я') || (c >= 'А' && c <= 'Я') || (c == 'ё' || c == '-' || c == 8 || c == 32));
        }


        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && e.KeyChar != 8)
                e.Handled = true;
        }
        int visible = 0;
        int y = 0;
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
            ExcelApp.Rows[1].Columns[3] = "Транспорт";
            ExcelApp.Rows[visible + 3].Columns[3] = "Ридецкая Анна Михайловна";
            for (int i = 1; i < dataGridView1.Columns.Count; i++)
            {
                ExcelApp.Cells[2, i] = dataGridView1.Columns[i].HeaderText;

            }
            
            for (int j = 1; j < dataGridView1.ColumnCount; j++)
            {y = 0;
                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    if (dataGridView1.Rows[i].Cells[j].Visible == true)
                    {
                      
                            ExcelApp.Cells[y + 3, j] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                        y++;
                    }
                }
            }
            for (int i = 0; i < visible; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    Microsoft.Office.Interop.Excel.Range range = ExcelWorkSheet.Range[$"A1:E{visible + 3}"];
                    ExcelWorkSheet.Range[$"A1:E{visible + 3}"].Cells.WrapText = true;
                    range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders.Weight = Excel.XlBorderWeight.xlMedium;
                }
            }
            ExcelApp.Columns["D"].Delete();
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }


        private void button7_Click(object sender, EventArgs e)
        {
            clear();
            dataGridView1.Enabled = true;
            if (panel3.Visible == true)
            {
                panel3.Visible = false;
                checkBox1.Checked = false;
                checkBox3.Checked = false;
                comboBox4.Items.Clear();
                comboBox5.Items.Clear();
                comboBox5.Text = "";
                train();
                button7.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (48).png");
            }
            else
            {
                checkBox1.Checked = false;
                checkBox3.Checked = false;
                comboBox4.Items.Clear();
                comboBox5.Items.Clear();

                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    if (comboBox4.Items.Contains(dataGridView1[1, i].Value.ToString()))
                    {
                    }
                    else
                    {

                        comboBox4.Items.Add(dataGridView1[1, i].Value.ToString());
                    }
                }
                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    if (comboBox5.Items.Contains(dataGridView1[2, i].Value.ToString()))
                    {
                    }
                    else
                    {

                        comboBox5.Items.Add(dataGridView1[2, i].Value.ToString());
                    }
                }
                comboBox4.SelectedIndex = -1;
                comboBox5.SelectedIndex = -1;
                panel3.Visible = true;
                button7.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (47).png");
            }
            panel2.Visible = false;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true && checkBox3.Checked == true)
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    dataGridView1.CurrentCell = null;
                    dataGridView1.Rows[i].Visible = false;

                    if (comboBox5.Text == dataGridView1[2, i].Value.ToString() && comboBox4.Text == dataGridView1[1, i].Value.ToString())
                    {
                        dataGridView1.Rows[i].Visible = true;
                    }
                    else
                    {
                        dataGridView1.Rows[i].Visible = false;
                    }

                }

            }
            else

            ////////////////
            ///// 2 по 1 ////

            if (checkBox1.Checked == true)
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    dataGridView1.CurrentCell = null;
                    dataGridView1.Rows[i].Visible = false;

                    if (comboBox4.Text == dataGridView1[1, i].Value.ToString())
                    {
                        dataGridView1.Rows[i].Visible = true;
                    }
                    else
                    {
                        dataGridView1.Rows[i].Visible = false;
                    }

                }

            }
            else

                if (checkBox3.Checked == true)
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    dataGridView1.CurrentCell = null;
                    dataGridView1.Rows[i].Visible = false;

                    if (comboBox5.Text == dataGridView1[2, i].Value.ToString())
                    {
                        dataGridView1.Rows[i].Visible = true;
                    }
                    else
                    {
                        dataGridView1.Rows[i].Visible = false;
                    }

                }

            }
            else if (checkBox1.Checked == false && checkBox3.Checked == false)
            {
                train();
            }
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

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            if (((System.Windows.Forms.TextBox)sender).Text.Length == 1)
                ((System.Windows.Forms.TextBox)sender).Text = ((System.Windows.Forms.TextBox)sender).Text.ToUpper();
            ((System.Windows.Forms.TextBox)sender).Select(((System.Windows.Forms.TextBox)sender).Text.Length, 0);
        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            e.Handled = !((c >= 'а' && c <= 'я') || (c >= 'А' && c <= 'Я') || (c == 'ё' || c == 8 || c == 32));
        }
    }
}
