using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace Agent.Form
{
    public partial class Worker : UserControl
    {
        datebase db = new datebase();
        public SqlConnection sqlConnection = null;
        public Worker()
        {
            InitializeComponent();
            sqlConnection = new SqlConnection(db.connection);
        }
        int id = 0;
        public void Insurer_load()
        {
            sqlConnection.Open();
            DataSet dataSet = new DataSet();
            SqlDataAdapter command = new SqlDataAdapter($@"Select idworker, firstname as Фамилия, worker.name as Имя, 
lastname as Отчество, datereception as [Дата приема], 
datelayoffs as [Дата увольнения],  worker.idpost, 
post.name as Должность, post.idwork , work.name as Отдел , 
phone as Телефон,email as Почта, login as Логин, password as Пароль 
from worker inner join post on worker.idpost=post.idpost 
inner join work on post.idwork=work.idwork", sqlConnection);
            command.Fill(dataSet);
            dataGridView1.DataSource = dataSet.Tables[0].DefaultView;
            sqlConnection.Close();
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[6].Visible = false;
            dataGridView1.Columns[8].Visible = false;
            dataGridView1.AllowUserToAddRows = false;
            sqlConnection.Close();
        }
        private void Insurer_Load(object sender, EventArgs e)
        {
            Insurer_load();
            comboBoxwork();
            panel2.Visible = false;
            panel3.Visible = false;
            dateTimePicker1.MinDate = Convert.ToDateTime("01.01.1980");
            dateTimePicker1.MaxDate = DateTime.Today;
            dateTimePicker1.MinDate = Convert.ToDateTime("01.01.1980");
            dateTimePicker2.MaxDate = DateTime.Today;
           button7.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (48).png");

        }
        public void comboBoxwork()
        {
            sqlConnection.Close();
            sqlConnection.Open();
            string query = "select idwork,name from work";
            SqlDataAdapter sqlDbDataAdapter = new SqlDataAdapter(query, sqlConnection);
            DataSet dataSet = new DataSet();
            sqlDbDataAdapter.Fill(dataSet);
            comboBox3.DataSource = dataSet.Tables[0];
            comboBox3.DisplayMember = "name";
            comboBox3.ValueMember = "idwork";
            comboBox3.SelectedIndex = -1;
            sqlConnection.Close();

        }
        public void comboBoxposition()
        {
            try
            {
                sqlConnection.Close();
                sqlConnection.Open();
                string query = $@"Select idpost, post.name as Должность from post inner join work on post.idwork=work.idwork where post.idwork={comboBox3.SelectedValue}";
                SqlDataAdapter sqlDbDataAdapter = new SqlDataAdapter(query, sqlConnection);
                DataSet dataSet = new DataSet();
                sqlDbDataAdapter.Fill(dataSet);
                comboBox4.DataSource = dataSet.Tables[0];
                comboBox4.DisplayMember = "Должность";
                comboBox4.ValueMember = "idpost";
                comboBox4.SelectedIndex = -1;
                sqlConnection.Close();
            }
            catch { }
        }
        private void button4_Click(object sender, EventArgs e)
        {
            if (panel2.Visible == false)
            {
                button7.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (48).png");
                clear();
                panel3.Visible = false;
                panel2.Visible = true;
                comboBox4.Visible = false;
                label6.Text = "Добавить работника";
                dateTimePicker2.Visible = false;
                button11.Text = "Добавить";
                button11.Width = 175;
                button11.Left = 416;
                Insurer_load();
                dataGridView1.Enabled = true;
                button11.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (31).png");
            }
            else
            {
                clear();
                panel2.Visible = false;
            }
        }
        public void clear()
        {
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            dateTimePicker1.Value = DateTime.Today;
            dateTimePicker2.Value = DateTime.Today;
            maskedTextBox2.Text = "";
                comboBox4.SelectedIndex = -1;
            comboBox3.SelectedIndex = -1;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            button7.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (48).png");
            panel3.Visible = false;
            try
            {
                if (panel2.Visible == false)
                {

                    clear();
                    if (id != 0)
                    {
                        dataGridView1.Enabled = false;
                        panel2.Visible = true;
                        label6.Text = "Редактировать работника";
                        dateTimePicker1.Value = Convert.ToDateTime(dataGridView1.CurrentRow.Cells[4].Value.ToString());

                        if (dataGridView1.CurrentRow.Cells[5].Value.ToString() != "")
                        {
                            dateTimePicker2.Value = Convert.ToDateTime(dataGridView1.CurrentRow.Cells[5].Value.ToString());
                            dateTimePicker2.Visible = true;
                        }
                        else
                        {
                            dateTimePicker2.Visible = false;
                        }
                        textBox2.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                        textBox3.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                        textBox4.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                        comboBox3.Text = dataGridView1.CurrentRow.Cells[9].Value.ToString();                    
                        comboBox4.SelectedValue = Convert.ToInt32(dataGridView1.CurrentRow.Cells[6].Value.ToString());
                        maskedTextBox2.Text = dataGridView1.CurrentRow.Cells[10].Value.ToString();
                        textBox7.Text = dataGridView1.CurrentRow.Cells[11].Value.ToString();
                        textBox5.Text = dataGridView1.CurrentRow.Cells[12].Value.ToString();
                        textBox6.Text = dataGridView1.CurrentRow.Cells[13].Value.ToString();
                        button11.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (29).png");
                        button11.Text = "Редактировать";
                        button11.Width = 245;
                        button11.Left = 377;
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
                }
            }
            catch { }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            button7.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (48).png");
            clear();
            panel3.Visible = false;
            dataGridView1.Enabled = true;
            panel2.Visible = false;
            if (id != 0)
            {
                try
                {
                    if (MessageBox.Show($@"Вы уверены что хотите удалить работника {dataGridView1.CurrentRow.Cells[1].Value.ToString() + " " + dataGridView1.CurrentRow.Cells[2].Value.ToString() + " " + dataGridView1.CurrentRow.Cells[3].Value.ToString()}?", "Удаление", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {

                        sqlConnection.Open();
                        string query = $@"DELETE FROM [worker] WHERE [idworker] = 
                             {dataGridView1.CurrentRow.Cells[0].Value.ToString()}";
                        SqlCommand command = new SqlCommand(query, sqlConnection);
                        command.ExecuteNonQuery();
                        sqlConnection.Close();
                        Insurer_load();
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
            ExcelApp.Rows[1].Columns[4] = "Работники";
            ExcelApp.Rows[visible + 3].Columns[4] = "Ридецкая Анна Михайловна";
            for (int i = 1; i < dataGridView1.Columns.Count; i++)
            {
                ExcelApp.Cells[2, i] = dataGridView1.Columns[i].HeaderText;

            }
            int y = 0;
            for (int j = 1; j < dataGridView1.ColumnCount; j++)
            {
                y = 0;
                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    if (dataGridView1.Rows[i].Cells[j].Visible == true)
                    {
                        if (j == 4)
                        {
                            ExcelApp.Cells[y + 3, j] = dataGridView1.Rows[i].Cells[j].Value.ToString().Substring(0, 10);
                            y++;
                        }
                        else
                        {
                            if (j == 5 && dataGridView1.Rows[i].Cells[j].Value.ToString() != "")
                            {
                                ExcelApp.Cells[y + 3, j] = dataGridView1.Rows[i].Cells[j].Value.ToString().Substring(0, 10);
                                y++;
                            }
                            else
                            {
                                ExcelApp.Cells[y + 3, j] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                                y++;
                            }
                        }
                    }
                }

            }
            for (int i = 0; i < visible; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    Microsoft.Office.Interop.Excel.Range range = ExcelWorkSheet.Range[$"A1:M{visible + 3}"];
                    ExcelWorkSheet.Range[$"A1:M{visible + 3}"].Cells.WrapText = true;
                    range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders.Weight = Excel.XlBorderWeight.xlMedium;
                }
            }
            ExcelApp.Columns["F"].Delete();
            ExcelApp.Columns["G"].Delete();
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
                Insurer_load();
                checkBox1.Checked = false;
                checkBox2.Checked = false;
                button7.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (48).png");
            }
            else
            {
                checkBox1.Checked = false;
                checkBox2.Checked = false;
                panel3.Visible = true;
                button7.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (47).png");
            }
            panel2.Visible = false;
        }

        private void label7_Click(object sender, EventArgs e)
        {
            if (dateTimePicker2.Visible == false)
            {
                dateTimePicker2.Visible = true;
            }
            else
            {
                dateTimePicker2.Visible = false;
            }
        }
        int k = 0;
        int j = 0;
        int count = 0;
        private void button11_Click(object sender, EventArgs e)
        {
            //try
            //{
            
                if (label6.Text == "Добавить работника")
                {
                     k=0;
                if (textBox2.Text != "" && textBox3.Text != "" && textBox4.Text != "" && textBox5.Text != "" && textBox6.Text != "" && textBox7.Text != "" && count == 2  && maskedTextBox2.Text.Length == 18 && comboBox4.SelectedIndex != -1)
                    {
                        for (int i = 0; i < dataGridView1.RowCount; i++)
                        {
                            if (textBox2.Text.ToLower() + textBox3.Text.ToLower() + textBox4.Text.ToLower() == dataGridView1[1, i].Value.ToString().ToLower() + dataGridView1[2, i].Value.ToString().ToLower() + dataGridView1[3, i].Value.ToString().ToLower())
                            {
                                k++;
                            }
                        }
                    if (k == 0)
                    {
                        Regex r1 = new Regex(@"(\(25|29|33|44)\)\S*");
                        if (r1.IsMatch(maskedTextBox2.Text))
                        {
                            k = 0;
                            for (int i = 0; i < dataGridView1.RowCount; i++)
                            {
                                if (maskedTextBox2.Text == dataGridView1[10, i].Value.ToString().ToLower())
                                {
                                    k++;
                                }
                            }
                            if (k == 0)
                            {
                                k = 0;
                            Regex r2 = new Regex(@"(\S*(@mail\.ru|@gmail\.com|@yandex\.ru))$");
                            if (r2.IsMatch(textBox7.Text))
                            {
                                for (int i = 0; i < dataGridView1.RowCount; i++)
                                {
                                    if (textBox7.Text == dataGridView1[11, i].Value.ToString())
                                    {
                                        k++;
                                    }
                                }
                                if (k == 0)
                                {
                                    if (textBox5.Text.Length > 5 && textBox5.Text.Any(char.IsLetter))

                                    {
                                        k = 0;
                                        for (int i = 0; i < dataGridView1.RowCount; i++)
                                        {
                                            if (textBox5.Text == dataGridView1[12, i].Value.ToString())
                                            {
                                                k++;
                                            }
                                        }
                                        if (k == 0)
                                        {
                                            if (textBox6.Text.Any(char.IsLower) && textBox6.Text.Any(char.IsUpper) && textBox6.Text.Length > 8)
                                            {
                                                if (dateTimePicker2.Visible == true)
                                                {
                                                    sqlConnection.Open();
                                                    SqlCommand command = new SqlCommand($@"INSERT INTO [worker](firstname,name,lastname,datereception,datelayoffs,phone,idpost,login,password,email) VALUES (@f,@n,@l,@r,@y,@ph,@idp,@lo,@p,@e);", sqlConnection);
                                                    command.Parameters.AddWithValue("@f", (textBox2.Text));
                                                    command.Parameters.AddWithValue("@n", (textBox3.Text));
                                                    command.Parameters.AddWithValue("@l", (textBox4.Text));
                                                    command.Parameters.AddWithValue("@r", (dateTimePicker1.Value));
                                                    command.Parameters.AddWithValue("@y", (dateTimePicker2.Value));
                                                    command.Parameters.AddWithValue("@ph", (maskedTextBox2.Text));
                                                    command.Parameters.AddWithValue("@idp", (comboBox4.SelectedValue));
                                                    command.Parameters.AddWithValue("lo", (textBox5.Text));
                                                    command.Parameters.AddWithValue("@p", (textBox6.Text));
                                                    command.Parameters.AddWithValue("@e", (textBox7.Text));
                                                    command.ExecuteNonQuery();
                                                    sqlConnection.Close();
                                                    Insurer_load();

                                                }
                                                else
                                                {
                                                    sqlConnection.Open();
                                                    SqlCommand command = new SqlCommand($@"INSERT INTO [worker](firstname,name,lastname,datereception,phone,idpost,login,password,email) VALUES (@f,@n,@l,@r,@ph,@idp,@lo,@p,@e);", sqlConnection);
                                                    command.Parameters.AddWithValue("@f", (textBox2.Text));
                                                    command.Parameters.AddWithValue("@n", (textBox3.Text));
                                                    command.Parameters.AddWithValue("@l", (textBox4.Text));
                                                    command.Parameters.AddWithValue("@r", (dateTimePicker1.Value));
                                                    command.Parameters.AddWithValue("@ph", (maskedTextBox2.Text));
                                                    command.Parameters.AddWithValue("@idp", (comboBox4.SelectedValue));
                                                    command.Parameters.AddWithValue("lo", (textBox5.Text));
                                                    command.Parameters.AddWithValue("@p", (textBox6.Text));
                                                    command.Parameters.AddWithValue("@e", (textBox7.Text));
                                                    command.ExecuteNonQuery();
                                                    sqlConnection.Close();
                                                    Insurer_load();

                                                }
                                                clear();
                                                panel2.Visible = false;
                                            }
                                            else
                                            {

                                                MessageBox.Show("Пароль должен быть больше 8 символов, содержать одну букву верхнего и нижнего регистра!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                            }
                                        }
                                        else
                                        {
                                            MessageBox.Show("Логин занят", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show("Логин должен быть больше 5 символов и содержать хотя бы одну букву!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    }
                                }
                                else
                                {
                                    k = 0;
                                    MessageBox.Show("Почта занята!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                            else
                            {
                                MessageBox.Show("Введите почту корректно!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                        else
                        {
                            MessageBox.Show("Телефон зарегистрирован на другого страхователя!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                        else
                        {
                            MessageBox.Show("Телефонный код введен неверно!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else
                    {
                        k = 0;
                        MessageBox.Show("Такой страховщик уже есть!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    }
                    else
                    {
                        MessageBox.Show("Заполните поля!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                k = 0;j = 0;
                    if (textBox2.Text != "" && textBox3.Text != "" && textBox4.Text != "" && textBox5.Text != "" && textBox6.Text != "" && textBox7.Text != ""&&count == 2 &&  maskedTextBox2.Text.Length == 18 &&  comboBox4.SelectedIndex != -1)
                    {
                        for (int i = 0; i < dataGridView1.RowCount; i++)
                        {
                            if (textBox2.Text.ToLower() + textBox3.Text.ToLower() + textBox4.Text.ToLower() == dataGridView1[1, i].Value.ToString().ToLower() + dataGridView1[2, i].Value.ToString().ToLower() + dataGridView1[3, i].Value.ToString().ToLower())
                            {
                                k++;
                                j = Convert.ToInt32(dataGridView1[0, i].Value.ToString());
                            }
                        }
                    if (k == 0 || j == id)
                    {
                        Regex r1 = new Regex(@"(\(25|29|33|44)\)\S*");
                        if (r1.IsMatch(maskedTextBox2.Text))
                        {
                            k = 0; j = 0;
                            for (int i = 0; i < dataGridView1.RowCount; i++)
                            {
                                if (maskedTextBox2.Text == dataGridView1[10, i].Value.ToString().ToLower())
                                {
                                    k++;
                                    j = Convert.ToInt32(dataGridView1[0, i].Value.ToString());
                                }
                            }
                            if (k == 0|| j==id)
                            {
                                k = 0; j = 0;
                            Regex r2 = new Regex(@"(\S*(@mail\.ru|@gmail\.com|@yandex\.ru))$");
                            if (r2.IsMatch(textBox7.Text) && textBox7.Text.Length > 8)
                            {
                                for (int i = 0; i < dataGridView1.RowCount; i++)
                                {
                                    if (textBox7.Text == dataGridView1[11, i].Value.ToString())
                                    {
                                        k++;
                                        j = Convert.ToInt32(dataGridView1[0, i].Value.ToString());
                                    }
                                }
                                if (k == 0 || j == id)
                                {
                                    j = 0; k = 0;
                                    if (textBox5.Text.Length > 5 && textBox5.Text.Any(char.IsLetter))
                                    {
                                        k = 0; j = 0;
                                        for (int i = 0; i < dataGridView1.RowCount; i++)
                                        {
                                            if (textBox5.Text == dataGridView1[12, i].Value.ToString())
                                            {
                                                k++;
                                                j = Convert.ToInt32(dataGridView1[0, i].Value.ToString());
                                            }
                                        }
                                        if (k == 0 || j == id)
                                            {
                                                k = 0; j = 0;
                                                if (textBox6.Text.Any(char.IsLower) && textBox6.Text.Any(char.IsUpper) && textBox6.Text.Length > 8)
                                            {
                                                if (dateTimePicker2.Visible == true)
                                                {
                                                    sqlConnection.Open();
                                                    SqlCommand command = new SqlCommand($@"UPDATE worker SET firstname=@f,name=@n ," +
                                                    $"lastname=@l,datereception=@r, datelayoffs=@y,phone=@ph,idpost=@idp,login=@lo,password=@p,email=@e WHERE idworker=@id", sqlConnection);
                                                    command.Parameters.AddWithValue("@f", (textBox2.Text));
                                                    command.Parameters.AddWithValue("@n", (textBox3.Text));
                                                    command.Parameters.AddWithValue("@l", (textBox4.Text));
                                                    command.Parameters.AddWithValue("@r", (dateTimePicker1.Value));
                                                    command.Parameters.AddWithValue("@y", (dateTimePicker2.Value));
                                                    command.Parameters.AddWithValue("@ph", (maskedTextBox2.Text));
                                                    command.Parameters.AddWithValue("@idp", (comboBox4.SelectedValue));
                                                    command.Parameters.AddWithValue("lo", (textBox5.Text));
                                                    command.Parameters.AddWithValue("@p", (textBox6.Text));
                                                    command.Parameters.AddWithValue("@e", (textBox7.Text));
                                                    command.Parameters.AddWithValue("@id", (id));
                                                    command.ExecuteNonQuery();
                                                    sqlConnection.Close();
                                                    dataGridView1.Enabled = true;

                                                }
                                                else
                                                {
                                                    sqlConnection.Open();
                                                    SqlCommand command = new SqlCommand($@"UPDATE worker SET firstname=@f,name=@n ," +
                                                    $"lastname=@l,datereception=@r, datelayoffs=Null,phone=@ph,idpost=@idp,login=@lo,password=@p,email=@e WHERE idworker=@id", sqlConnection);
                                                    command.Parameters.AddWithValue("@f", (textBox2.Text));
                                                    command.Parameters.AddWithValue("@n", (textBox3.Text));
                                                    command.Parameters.AddWithValue("@l", (textBox4.Text));
                                                    command.Parameters.AddWithValue("@r", (dateTimePicker1.Value));
                                                    command.Parameters.AddWithValue("@ph", (maskedTextBox2.Text));
                                                    command.Parameters.AddWithValue("@idp", (comboBox4.SelectedValue));
                                                    command.Parameters.AddWithValue("lo", (textBox5.Text));
                                                    command.Parameters.AddWithValue("@p", (textBox6.Text));
                                                    command.Parameters.AddWithValue("@e", (textBox7.Text));
                                                    command.Parameters.AddWithValue("@id", (id));
                                                    command.ExecuteNonQuery();
                                                    sqlConnection.Close();
                                                    dataGridView1.Enabled = true;
                                                }
                                                clear();
                                                Insurer_load();
                                                panel2.Visible = false;
                                            }
                                            else
                                            {
                                                MessageBox.Show("Пароль должен быть больше 8 символов, содержать одну букву верхнего и нижнего регистра!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                            }
                                        }
                                        else
                                            {
                                                k = 0; j = 0;
                                                MessageBox.Show("Логин занят", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show("Логин должен быть больше 5 символов и содержать хотя бы одну букву!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    }
                                }
                                else
                                {
                                    k = 0; j = 0;
                                    MessageBox.Show("Почта занята!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                            else
                            {
                                MessageBox.Show("Введите почту корректно!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                            }
                            else
                            {
                                k = 0; j = 0;
                                MessageBox.Show("Телефон зарегистрирован на другого страхователя!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                        else
                        {
                            MessageBox.Show("Телефонный код введен неверно!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else
                    {
                        k = 0; j = 0;
                        MessageBox.Show("Такой страховщик уже есть!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    }
                    else
                    {
                        MessageBox.Show("Заполните поля!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
        //}
        //    catch { }
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            dateTimePicker2.MinDate = dateTimePicker1.Value;
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

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                checkBox2.Checked = false;
            }
            else
            {
                checkBox2.Checked = true;
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                checkBox1.Checked = false;
            }
            else
            {
                checkBox1.Checked = true;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    dataGridView1.CurrentCell = null;
                    dataGridView1.Rows[i].Visible = false;

                    if (dataGridView1[5, i].Value.ToString() != "")
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
            {
                if (checkBox2.Checked == true)
                {
                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        dataGridView1.CurrentCell = null;
                        dataGridView1.Rows[i].Visible = false;

                        if (dataGridView1[5, i].Value.ToString() == "")
                        {
                            dataGridView1.Rows[i].Visible = true;
                        }
                        else
                        {
                            dataGridView1.Rows[i].Visible = false;
                        }
                    }
                }
            }
        }

        private void textBox7_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            e.Handled = !((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || (c == '@' || c == 8 || c == '.'));
            if (e.KeyChar == '@')
            {
                if (((sender as TextBox).Text.IndexOf('@') != -1))
                {
                    e.Handled = true;
                    return;
                }
            }
        }

        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            e.Handled = !((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || (c >= '0' && c <= '9') || (c == '@' || c == 8 || c == '.' || c == '_' || c == ','));

        }

        private void textBox6_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            e.Handled = !((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || (c >= '0' && c <= '9') || (c == '@' || c == 8 || c == '.' || c == '_' || c == ',' || c == '#' || c == '+' || c == '!' || c == '$' || c == ':' || c == ';' || c == '%' || c == '^' || c == '&' || c == '*' || c == ')' || c == '(' || c == '-'));

        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox3.Text != "System.Data.DataRowView" && comboBox3.Text != "")
            {
                comboBox4.Visible = true;
                comboBoxposition();
            }
        }

        private void maskedTextBox2_TextChanged(object sender, EventArgs e)
        {
            String s = maskedTextBox2.Text;
            String[] words = s.Split(' ');
            count = words.Length;
        }
    }
}
