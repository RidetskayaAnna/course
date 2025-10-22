using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Agent.Form
{
    public partial class Policyholder : UserControl
    {
        datebase db = new datebase();
        public SqlConnection sqlConnection = null;
        public Policyholder()
        {
            InitializeComponent();
            sqlConnection = new SqlConnection(db.connection);
        }
        int id = 0;
        public void Policyholder_load()
        {
            sqlConnection.Open();
            DataSet dataSet = new DataSet();
            SqlDataAdapter command = new SqlDataAdapter($@"Select idpolicyholder,policyholder.idwork,firdtname as Фамилия, policyholder.name as Имя, lastname as Отчество,dateb as [Дата рождения],policyholder.idcity, city.name as [Город прописки],address as [Адрес], passport as [Паспорт],numar as [Идентификационный номер],organ as [Орган, выдавший паспорт],datep as [Дата выдачи паспорта],phone as Телефон,work.name as [Место работы],position.name as [Должность], heal as [Группа здоровья],sport as [Занятие спортом],email as Почта,login as Логин,password as Пароль,work.idwork from policyholder inner join city on policyholder.idcity=city.idcity  inner join position on position.idposition=policyholder.idwork inner join work on work.idwork=position.idwork", sqlConnection);
            command.Fill(dataSet);
            dataGridView1.DataSource = dataSet.Tables[0].DefaultView;
            sqlConnection.Close();
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[1].Visible = false;
            dataGridView1.Columns[6].Visible = false;
            dataGridView1.Columns[21].Visible = false;
            dataGridView1.AllowUserToAddRows = false;
            sqlConnection.Close();
        }
        public void comboBoxx()
        {
            sqlConnection.Close();
            sqlConnection.Open();
            string query = "select idcity,name from city";
            SqlDataAdapter sqlDbDataAdapter = new SqlDataAdapter(query, sqlConnection);
            DataSet dataSet = new DataSet();
            sqlDbDataAdapter.Fill(dataSet);
            comboBox1.DataSource = dataSet.Tables[0];
            comboBox1.DisplayMember = "name";
            comboBox1.ValueMember = "idcity";
            comboBox1.SelectedIndex = -1;
            sqlConnection.Close();

        }
      
        public void comboBoxcity()
        {
            sqlConnection.Close();
            sqlConnection.Open();
            string query = "select idcity,name from city";
            SqlDataAdapter sqlDbDataAdapter = new SqlDataAdapter(query, sqlConnection);
            DataSet dataSet = new DataSet();
            sqlDbDataAdapter.Fill(dataSet);
            comboBox2.DataSource = dataSet.Tables[0];
            comboBox2.DisplayMember = "name";
            comboBox2.ValueMember = "idcity";
            comboBox2.SelectedIndex = -1;
            sqlConnection.Close();

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
                string query = $@"Select idposition, position.name as Должность from position inner join work on position.idwork=work.idwork where position.idwork={comboBox3.SelectedValue}";
                SqlDataAdapter sqlDbDataAdapter = new SqlDataAdapter(query, sqlConnection);
                DataSet dataSet = new DataSet();
                sqlDbDataAdapter.Fill(dataSet);
                comboBox4.DataSource = dataSet.Tables[0];
                comboBox4.DisplayMember = "Должность";
                comboBox4.ValueMember = "idposition";
                comboBox4.SelectedIndex = -1;
                sqlConnection.Close();
            }
            catch { }
        }
        public void comboBoxwork2()
        {
            sqlConnection.Close();
            sqlConnection.Open();
            string query = "select idwork,name from work";
            SqlDataAdapter sqlDbDataAdapter = new SqlDataAdapter(query, sqlConnection);
            DataSet dataSet = new DataSet();
            sqlDbDataAdapter.Fill(dataSet);
            comboBox7.DataSource = dataSet.Tables[0];
            comboBox7.DisplayMember = "name";
            comboBox7.ValueMember = "idwork";
            comboBox7.SelectedIndex = -1;
            sqlConnection.Close();

        }
        public void comboBoxposition2()
        {
            try
            {
                sqlConnection.Close();
                sqlConnection.Open();
                string query = $@"Select distinct position.name as Должность from position";
                SqlDataAdapter sqlDbDataAdapter = new SqlDataAdapter(query, sqlConnection);
                DataSet dataSet = new DataSet();
                sqlDbDataAdapter.Fill(dataSet);
                comboBox8.DataSource = dataSet.Tables[0];
                comboBox8.DisplayMember = "Должность";
                comboBox8.ValueMember = "Должность";
                comboBox8.SelectedIndex = -1;
                sqlConnection.Close();
            }
            catch { }
        }
        private void Policyholder_Load(object sender, EventArgs e)
        {
            Policyholder_load();
            comboBoxx();
            comboBoxcity();
            comboBoxwork();
            comboBoxposition2();
            comboBoxwork2();
            panel3.Visible = false;
            panel2.Visible = false;
            button7.Image = new Bitmap(@"D:\College\4kurs\Praktica4Kurs\proga\Agent\Agent\Resources\pngwing.com (48).png");
            dateTimePicker2.MaxDate = DateTime.Today.AddYears(-18).AddDays(-1);
            dateTimePicker2.MinDate = DateTime.Today.AddYears(-74).AddDays(-1);
        }
        public void clear()
        {
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            textBox8.Text = "";
            maskedTextBox1.Text = "";
            textBox9.Text = "";
            maskedTextBox2.Text = "";
            dateTimePicker1.MaxDate = DateTime.Today;
            dateTimePicker1.MinDate= DateTime.Today.AddDays(-3651);
            dateTimePicker1.Value = DateTime.Today;
            comboBox1.SelectedIndex = -1;
            checkBox2.Checked = false;
            textBox10.Text = "";
            textBox12.Text = "";
            textBox11.Text = "";
            comboBox3.SelectedIndex = -1;
            comboBox4.SelectedIndex = -1;
            comboBox5.SelectedIndex = -1;
        }
        private void button4_Click(object sender, EventArgs e)
        {
            if (panel2.Visible == false)
            {
                button7.Image = new Bitmap(@"D:\College\4kurs\Praktica4Kurs\proga\Agent\Agent\Resources\pngwing.com (48).png");
              clear();
               panel3.Visible = false;
                panel5.Visible = false;
                panel6.Visible = false;
                panel4.Visible = false;
                panel2.Visible = true;
                label6.Text = "Добавить страхователя";
                label6.Visible = true;
                button6.BringToFront();
                button6.Visible = true;
                button6.Top = 687;
                button6.Left = 909;
                button8.Visible = false;
                button11.Visible = false;
            button11.Text = "Добавить";
            Policyholder_load();
            dataGridView1.Enabled = true;
            button11.Image = new Bitmap(@"D:\College\4kurs\Praktica4Kurs\proga\Agent\Agent\Resources\pngwing.com (31).png");
            }
            else
            {
                clear();
                panel2.Visible = false;
                button6.Visible= false;
                button8.Visible = false;
                label6.Visible = false;
            }
        }
        int prod = 0;
        private void button1_Click(object sender, EventArgs e)
        {
            button7.Image = new Bitmap(@"D:\College\4kurs\Praktica4Kurs\proga\Agent\Agent\Resources\pngwing.com (48).png");
            try
            {
                if (panel2.Visible == false)
                {
                    prod = 0;
                    panel3.Visible = false;
                    clear();
                    if (id != 0)
                    {
                        dataGridView1.Enabled = false;
                        panel2.Visible = true;
                        label6.Text = "Редактировать страхователя";

                        panel5.Visible = false;
                        panel6.Visible = false;
                        panel4.Visible = false;
                        label6.Visible = true;
                        button6.Visible = true;
                        button6.BringToFront();
                        button6.Top = 687;
                        button6.Left = 909;
                        button8.Visible = false;
                        button11.Visible = false;

                        textBox2.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                        textBox3.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                        textBox4.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
                        textBox5.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString();
                        maskedTextBox1.Text = dataGridView1.CurrentRow.Cells[10].Value.ToString();
                        char[] o = { 'Р', 'О', 'В', 'Д' };
                        prod = dataGridView1.CurrentRow.Cells[11].Value.ToString().LastIndexOfAny(o);
                        dateTimePicker1.Value = Convert.ToDateTime(dataGridView1.CurrentRow.Cells[12].Value.ToString());
                        textBox8.Text = dataGridView1.CurrentRow.Cells[11].Value.ToString().Substring(0, prod - 4);
                        textBox9.Text = dataGridView1.CurrentRow.Cells[11].Value.ToString().Substring(prod - 4, 4);
                        textBox6.Text = dataGridView1.CurrentRow.Cells[9].Value.ToString().Substring(2, 7);
                        comboBox1.SelectedValue = dataGridView1.CurrentRow.Cells[6].Value.ToString();
                        maskedTextBox2.Text = dataGridView1.CurrentRow.Cells[13].Value.ToString();
                        dateTimePicker2.Value = Convert.ToDateTime(dataGridView1.CurrentRow.Cells[5].Value.ToString());
                        comboBox3.SelectedValue = dataGridView1.CurrentRow.Cells[21].Value.ToString();
                        comboBox4.SelectedValue = Convert.ToInt32(dataGridView1.CurrentRow.Cells[1].Value.ToString());
                        comboBox5.SelectedItem = dataGridView1.CurrentRow.Cells[16].Value.ToString();

                        if (dataGridView1.CurrentRow.Cells[17].Value.ToString()=="Да")
                        {
                            checkBox2.Checked = true;
                        }
                        else
                        {
                            checkBox2.Checked= false;
                        }
                        textBox11.Text = dataGridView1.CurrentRow.Cells[19].Value.ToString();
                        textBox12.Text = dataGridView1.CurrentRow.Cells[20].Value.ToString();
                        textBox10.Text = dataGridView1.CurrentRow.Cells[18].Value.ToString();
                        button11.Image = new Bitmap(@"D:\College\4kurs\Praktica4Kurs\proga\Agent\Agent\Resources\pngwing.com (29).png");
                        button11.Text = "Редактировать";
                        //button11.Width = 207;
                        //button11.Left = 553;
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
                    button6.Visible = false;
                    button8.Visible = false;
                    label6.Visible = false;
                }
            }
            catch { }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            button7.Image = new Bitmap(@"D:\College\4kurs\Praktica4Kurs\proga\Agent\Agent\Resources\pngwing.com (48).png");
            clear();
            dataGridView1.Enabled = true;
            panel2.Visible = false;
            panel3.Visible = false;
            if (id != 0)
            {
                try
                {
                    if (MessageBox.Show($@"Вы уверены что хотите удалить страхователя {dataGridView1.CurrentRow.Cells[1].Value.ToString() + " " + dataGridView1.CurrentRow.Cells[2].Value.ToString() + " " + dataGridView1.CurrentRow.Cells[3].Value.ToString()}?", "Удаление", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {

                        sqlConnection.Open();
                        string query = $@"DELETE FROM [policyholder] WHERE [idpolicyholder] = 
                             {dataGridView1.CurrentRow.Cells[0].Value.ToString()}";
                        SqlCommand command = new SqlCommand(query, sqlConnection);
                        command.ExecuteNonQuery();
                        sqlConnection.Close();
                        Policyholder_load();
                    }
                }
                catch { MessageBox.Show("Удаление невозможно, данные все еще нужны!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            }
            else { MessageBox.Show("Выберите строку для удаления!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }

        }
        int k = 0;
        int j = 0;
        int count = 0;
        int count2 = 0;
        String passport = "";
        private void button11_Click(object sender, EventArgs e)
        {
            try
            {
                if (label6.Text == "Добавить страхователя")
                {
                    if (textBox2.Text != "" && textBox3.Text != "" && textBox4.Text != ""&&textBox12.Text != "" && textBox10.Text != "" && textBox11.Text != "" && textBox5.Text != "" && textBox6.Text != "" && textBox7.Text != "" && textBox6.Text.Length == 7 && count == 2 && count2 == 1 && maskedTextBox2.Text.Length == 18 && maskedTextBox1.Text.Length == 14&&comboBox4.SelectedIndex!=-1 && comboBox5.SelectedIndex != -1)
                    {
                        Regex r1 = new Regex(@"(\(25|29|33|44)\)\S*");
                        if (r1.IsMatch(maskedTextBox2.Text))
                        {
                            k = 0;
                            for (int i = 0; i < dataGridView1.RowCount; i++)
                            {
                                if (maskedTextBox1.Text == dataGridView1[8, i].Value.ToString())
                                {
                                    k++;
                                }
                            }
                            if (k == 0)
                            {
                                ////
                                k = 0;
                                for (int i = 0; i < dataGridView1.RowCount; i++)
                                {
                                    if (passport == dataGridView1[7, i].Value.ToString())
                                    {
                                        k++;
                                    }
                                }

                                if (k == 0)
                                {
                                    ///
                                    k = 0;
                                    for (int i = 0; i < dataGridView1.RowCount; i++)
                                    {
                                        if (maskedTextBox2.Text == dataGridView1[8, i].Value.ToString().ToLower())
                                        {
                                            k++;
                                        }
                                    }
                                    if (k == 0)
                                    {
                                        
                                            k = 0;
                                            Regex r2 = new Regex(@"(\S*(@mail\.ru|@gmail\.com|@yandex\.ru))$");
                                            if (r2.IsMatch(textBox10.Text))
                                            {
                                                for (int i = 0; i < dataGridView1.RowCount; i++)
                                                {
                                                    if (textBox7.Text == dataGridView1[6, i].Value.ToString())
                                                    {
                                                        k++;
                                                    }
                                                }
                                                if (k == 0)
                                                {
                                                    if (textBox11.Text.Length > 5 && textBox11.Text.Any(char.IsLetter))

                                                    {
                                                        if (textBox12.Text.Any(char.IsLower) && textBox12.Text.Any(char.IsUpper) && textBox12.Text.Length > 8)
                                                        {

                                                            sqlConnection.Open();
                                        SqlCommand command = new SqlCommand($@"INSERT INTO [policyholder](firdtname,name,lastname,address,passport,idcity,numar,organ,phone,datep,idwork,email,heal,sport,login,password,dateb) VALUES (@f,@n,@l,@a,@p,@c,@u,@o,@h,@w,@idw,@e,@he,@s,@lo,@pa,@b);", sqlConnection);
                                        command.Parameters.AddWithValue("@f", (textBox2.Text));
                                        command.Parameters.AddWithValue("@n", (textBox3.Text));
                                        command.Parameters.AddWithValue("@l", (textBox4.Text));
                                        command.Parameters.AddWithValue("@a", (textBox5.Text));
                                        command.Parameters.AddWithValue("@p", (textBox7.Text + textBox6.Text));
                                        command.Parameters.AddWithValue("@c", (comboBox1.SelectedValue));
                                        command.Parameters.AddWithValue("@u", (maskedTextBox1.Text));
                                        command.Parameters.AddWithValue("@o", (textBox8.Text+textBox9.Text));
                                        command.Parameters.AddWithValue("@h", (maskedTextBox2.Text));
                                        command.Parameters.AddWithValue("@w", (dateTimePicker1.Value));
                                                        command.Parameters.AddWithValue("@b", (dateTimePicker2.Value));
                                                        command.Parameters.AddWithValue("@idw", (comboBox4.SelectedValue));
                                                        command.Parameters.AddWithValue("@e", (textBox10.Text));
                                                        command.Parameters.AddWithValue("@he", (comboBox5.SelectedItem));
                                                        if (checkBox2.Checked == true)
                                                        {
                                                            command.Parameters.AddWithValue("@s", "Да");
                                                        }
                                                        else
                                                        {
                                                            command.Parameters.AddWithValue("@s","Нет");
                                                        }
                                                        command.Parameters.AddWithValue("@lo", (textBox11.Text));
                                                        command.Parameters.AddWithValue("@pa", (textBox12.Text));
                                                        command.ExecuteNonQuery();
                                        sqlConnection.Close();
                                        Policyholder_load();
                                        clear();
                                                        panel2.Visible = false;
                                                        button6.Visible = false;
                                                        button8.Visible = false;
                                                        label6.Visible = false;
                                                        panel5.Visible = false;
                                                        panel6.Visible = false;
                                                        panel4.Visible = false;
                                                        button11.Visible = false;
                                                    }
                                                    else
                                                    {
                                                        MessageBox.Show("Пароль должен быть больше 8 символов, содержать одну букву верхнего и нижнего регистра!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                                    ///
                                }
                                else
                                {
                                    MessageBox.Show("Паспорт зарегистрирован на другого страхователя!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                                ////
                            }
                            else
                        {
                            MessageBox.Show("Идентификационный номер не уникален!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }      
                        }
                        else
                            {
                                MessageBox.Show("Телефонный код введен неверно!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                    }
                    else
                    {
                        MessageBox.Show("Заполните поля!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {

                if (textBox2.Text != "" && textBox3.Text != "" && textBox4.Text != "" && textBox12.Text != "" && textBox10.Text != "" && textBox11.Text != "" && textBox5.Text != "" && textBox6.Text != "" && textBox7.Text != "" && textBox6.Text.Length == 7 && count == 2 && count2 == 1 && maskedTextBox2.Text.Length == 18 && maskedTextBox1.Text.Length == 14 && comboBox4.SelectedIndex != -1 && comboBox5.SelectedIndex != -1)
                {
                    Regex r1 = new Regex(@"(\(25|29|33|44)\)\S*");
                    if (r1.IsMatch(maskedTextBox2.Text))
                    {
                        k = 0; j = 0;
                        for (int i = 0; i < dataGridView1.RowCount; i++)
                        {
                            if (maskedTextBox1.Text == dataGridView1[8, i].Value.ToString())
                            {
                                k++;
                                j = Convert.ToInt32(dataGridView1[0, i].Value.ToString());
                            }
                        }
                        if (k == 0||j==id)
                        {
                            ////
                            k = 0; j = 0;
                            for (int i = 0; i < dataGridView1.RowCount; i++)
                            {
                                if (passport == dataGridView1[7, i].Value.ToString())
                                {
                                    k++;
                                    j = Convert.ToInt32(dataGridView1[0, i].Value.ToString());
                                }
                            }

                            if (k == 0||j==id)
                            {
                                ///
                                k = 0; j = 0;
                                for (int i = 0; i < dataGridView1.RowCount; i++)
                                {
                                    if (maskedTextBox2.Text == dataGridView1[8, i].Value.ToString().ToLower())
                                    {
                                        k++;
                                        j = Convert.ToInt32(dataGridView1[0, i].Value.ToString());
                                    }
                                }
                                if (k == 0||j==id)
                                {
                                    k = 0; j = 0;
                                    Regex r2 = new Regex(@"(\S*(@mail\.ru|@gmail\.com|@yandex\.ru))$");
                                    if (r2.IsMatch(textBox10.Text))
                                    {
                                        for (int i = 0; i < dataGridView1.RowCount; i++)
                                        {
                                            if (textBox7.Text == dataGridView1[6, i].Value.ToString())
                                            {
                                                k++;
                                                j = Convert.ToInt32(dataGridView1[0, i].Value.ToString());
                                            }
                                        }
                                        if (k == 0||j==id)
                                        { k = 0;j = 0;
                                            if (textBox11.Text.Length > 5 && textBox11.Text.Any(char.IsLetter))

                                            {
                                                if (textBox12.Text.Any(char.IsLower) && textBox12.Text.Any(char.IsUpper) && textBox12.Text.Length > 8)
                                                {

                                                   
                                    sqlConnection.Open();
                                    SqlCommand command = new SqlCommand($@"UPDATE policyholder SET firdtname=@f,name=@n ," +
                                    $"lastname=@l,address=@a, passport=@p, idcity=@c,numar=@u,organ=@o,phone=@h,datep=@w,dateb=@b,login=@lo,password=@pa,sport=@s,heal=@he,email=@e,idwork=@idw WHERE idpolicyholder=@id", sqlConnection);
                                    command.Parameters.AddWithValue("@f", (textBox2.Text));
                                    command.Parameters.AddWithValue("@n", (textBox3.Text));
                                    command.Parameters.AddWithValue("@l", (textBox4.Text));
                                    command.Parameters.AddWithValue("@a", (textBox5.Text));
                                    command.Parameters.AddWithValue("@p", (textBox7.Text + textBox6.Text));
                                    command.Parameters.AddWithValue("@c", (comboBox1.SelectedValue));
                                    command.Parameters.AddWithValue("@u", (maskedTextBox1.Text));
                                    command.Parameters.AddWithValue("@o", (textBox8.Text + textBox9.Text));
                                    command.Parameters.AddWithValue("@h", (maskedTextBox2.Text));
                                        command.Parameters.AddWithValue("@w", (dateTimePicker1.Value));
                                                        command.Parameters.AddWithValue("@b", (dateTimePicker2.Value));
                                                        command.Parameters.AddWithValue("@idw", (comboBox4.SelectedValue));
                                                        command.Parameters.AddWithValue("@e", (textBox10.Text));
                                                        command.Parameters.AddWithValue("@he", (comboBox5.SelectedItem));
                                                        if (checkBox2.Checked == true)
                                                        {
                                                            command.Parameters.AddWithValue("@s", "Да");
                                                        }
                                                        else
                                                        {
                                                            command.Parameters.AddWithValue("@s", "Нет");
                                                        }
                                                        command.Parameters.AddWithValue("@lo", (textBox11.Text));
                                                        command.Parameters.AddWithValue("@pa", (textBox12.Text));
                                                        command.Parameters.AddWithValue("@id", (id));
                                    command.ExecuteNonQuery();
                                    sqlConnection.Close();
                                    dataGridView1.Enabled = true;
                                    clear();
                                    Policyholder_load();
                                                        panel2.Visible = false;
                                                        button6.Visible = false;
                                                        button8.Visible = false;
                                                        label6.Visible = false;
                                                        panel5.Visible = false;
                                                        panel6.Visible = false;
                                                        panel4.Visible = false;
                                                        button11.Visible = false;
                                                    }
                                                else
                                                {
                                                    MessageBox.Show("Пароль должен быть больше 8 символов, содержать одну букву верхнего и нижнего регистра!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                                ///
                            }
                            else
                            {
                                MessageBox.Show("Паспорт зарегистрирован на другого страхователя!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                            ////
                        }
                        else
                        {
                            MessageBox.Show("Идентификационный номер не уникален!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Телефонный код введен неверно!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
        string city = "";
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
               
                string query1 = $@"Select area from city where name='{comboBox1.Text}'";
                System.Data.DataTable data = new System.Data.DataTable();
                SqlDataAdapter command1 = new SqlDataAdapter(query1, sqlConnection);
                command1.Fill(data);
                DataColumn column = data.Columns[0];
                DataRow row = data.Rows[0];
                city = row[column].ToString();
                if (city == "Брестская область")
                {
                    textBox8.Text = comboBox1.Text;
                    textBox9.Text = " РОВД";
                    textBox7.Text = "АВ";
                }
                if (city == "Витебская область")
                {
                    textBox8.Text = comboBox1.Text;
                    textBox9.Text = " РОВД";
                    textBox7.Text = "ВМ ";
                }
                if (city == "Гомельская область")
                {
                    textBox8.Text = comboBox1.Text;
                    textBox9.Text = " РОВД";
                    textBox7.Text = "НВ";
                }
                if (city == "Гродненская область")
                {
                    textBox8.Text = comboBox1.Text;
                    textBox9.Text = " РОВД";
                    textBox7.Text = "КН";
                }
                if (city == "город Минск")
                {
                    textBox8.Text = comboBox1.Text;
                    textBox9.Text = " РОВД";
                    textBox7.Text = "МР";
                }
                if (city == "Минская область")
                {
                    textBox8.Text = comboBox1.Text;
                    textBox9.Text = " РОВД";
                    textBox7.Text = "МС";
                }
                if (city == "Могилёвская область")
                {
                    textBox8.Text = comboBox1.Text;
                    textBox9.Text = " РОВД";
                    textBox7.Text = "КВ";
                }
            }
            catch
            {

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

        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            e.Handled = !((c >= 'а' && c <= 'я') || (c >= 'А' && c <= 'Я') || (c == 'ё' || c == 8 || c == 32||c==','||c=='.'||c=='0'||c=='1'||c=='2'||c=='3'||c=='4'||c=='5'||c=='6'||c=='7'||c=='8'||c=='9'));
        }

        private void textBox6_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && e.KeyChar != 8 )
                e.Handled = true;
          
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            id = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
        }

        private void maskedTextBox2_TextChanged(object sender, EventArgs e)
        {     
                String s = maskedTextBox2.Text;
                String[] words = s.Split(' ');
                count=words.Length;
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            passport = textBox7.Text + textBox6.Text;
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            passport = textBox7.Text + textBox6.Text;
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
            ExcelApp.Rows[1].Columns[5] = "Страхователи";
            ExcelApp.Rows[visible + 3].Columns[5] = "Ридецкая Анна Михайловна";
            for (int i = 2; i < dataGridView1.Columns.Count; i++)
            {
                ExcelApp.Cells[2, i-1] = dataGridView1.Columns[i].HeaderText;

            }
            for (int j = 1; j < dataGridView1.ColumnCount; j++)
            {
                for (int i = 0; i < visible; i++)
                {
                    if (dataGridView1.Rows[i].Cells[j].Visible == true)
                    {
                        if (j==11) {
                            ExcelApp.Cells[i + 3, j - 1] = dataGridView1.Rows[i].Cells[j].Value.ToString().Substring(0,10);
                        }
                        else
                        {
                            ExcelApp.Cells[i + 3, j - 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                        }
                    }
                }
            }
            for (int i = 0; i < visible; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    Microsoft.Office.Interop.Excel.Range range = ExcelWorkSheet.Range[$"A1:R{visible + 3}"];
                    ExcelWorkSheet.Range[$"A1:R{visible + 3}"].Cells.WrapText = true;
                    range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders.Weight = Excel.XlBorderWeight.xlMedium;
                }
            }
            ExcelApp.Columns["D"].Delete();
            ExcelApp.Columns["R"].Delete();
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
                comboBox2.SelectedIndex =- 1;
                comboBox7.SelectedIndex = -1;
                comboBox8.SelectedIndex = -1;
                comboBox6.SelectedIndex = -1;
                checkBox1.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                Policyholder_load();
                button7.Image = new Bitmap(@"D:\College\4kurs\Praktica4Kurs\proga\Agent\Agent\Resources\pngwing.com (48).png");
            }
            else
            {
                panel3.Visible = true;
                button7.Image = new Bitmap(@"D:\College\4kurs\Praktica4Kurs\proga\Agent\Agent\Resources\pngwing.com (47).png");
            }
            panel2.Visible = false;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ///// 1 по 4 ////

            if (checkBox1.Checked == true && checkBox3.Checked == true && checkBox4.Checked == true && checkBox5.Checked == true)
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    dataGridView1.CurrentCell = null;
                    dataGridView1.Rows[i].Visible = false;

                    if (comboBox2.Text == dataGridView1[7,i].Value.ToString()&& comboBox7.Text == dataGridView1[14, i].Value.ToString() && comboBox8.Text == dataGridView1[15, i].Value.ToString() && comboBox6.Text == dataGridView1[16, i].Value.ToString())
                    {
                        dataGridView1.Rows[i].Visible = true;
                    }
                    else
                    {
                        dataGridView1.Rows[i].Visible = false;
                    }

                }

            } else

            ////////////////
            ///// 4 по 3 ////

            if (checkBox1.Checked == true && checkBox3.Checked == true && checkBox4.Checked == true)
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    dataGridView1.CurrentCell = null;
                    dataGridView1.Rows[i].Visible = false;

                    if (comboBox2.Text == dataGridView1[7, i].Value.ToString() && comboBox7.Text == dataGridView1[14, i].Value.ToString() && comboBox8.Text == dataGridView1[15, i].Value.ToString())
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
             if (checkBox1.Checked == true && checkBox4.Checked == true && checkBox5.Checked == true)
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    dataGridView1.CurrentCell = null;
                    dataGridView1.Rows[i].Visible = false;

                    if (comboBox2.Text == dataGridView1[7, i].Value.ToString() && comboBox8.Text == dataGridView1[15, i].Value.ToString() && comboBox6.Text == dataGridView1[16, i].Value.ToString())
                    {
                        dataGridView1.Rows[i].Visible = true;
                    }
                    else
                    {
                        dataGridView1.Rows[i].Visible = false;
                    }

                }

            }else
             if (checkBox1.Checked == true && checkBox3.Checked == true && checkBox5.Checked == true)
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    dataGridView1.CurrentCell = null;
                    dataGridView1.Rows[i].Visible = false;

                    if (comboBox2.Text == dataGridView1[7, i].Value.ToString() && comboBox7.Text == dataGridView1[14, i].Value.ToString() && comboBox6.Text == dataGridView1[16, i].Value.ToString())
                    {
                        dataGridView1.Rows[i].Visible = true;
                    }
                    else
                    {
                        dataGridView1.Rows[i].Visible = false;
                    }

                }

            }else
             if ( checkBox3.Checked == true && checkBox4.Checked == true && checkBox5.Checked == true)
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    dataGridView1.CurrentCell = null;
                    dataGridView1.Rows[i].Visible = false;

                    if ( comboBox7.Text == dataGridView1[14, i].Value.ToString() && comboBox8.Text == dataGridView1[15, i].Value.ToString() && comboBox6.Text == dataGridView1[16, i].Value.ToString())
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

            /////////////////

            //// 6 по 2 ////

            //1
            if (checkBox1.Checked == true && checkBox3.Checked == true)
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    dataGridView1.CurrentCell = null;
                    dataGridView1.Rows[i].Visible = false;

                    if (comboBox2.Text == dataGridView1[7, i].Value.ToString() && comboBox7.Text == dataGridView1[14, i].Value.ToString())
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
            //2
            if (checkBox1.Checked == true  && checkBox4.Checked == true)
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    dataGridView1.CurrentCell = null;
                    dataGridView1.Rows[i].Visible = false;

                    if (comboBox2.Text == dataGridView1[7, i].Value.ToString() && comboBox8.Text == dataGridView1[15, i].Value.ToString())
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
            //3
             if (checkBox1.Checked == true  && checkBox5.Checked == true)
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    dataGridView1.CurrentCell = null;
                    dataGridView1.Rows[i].Visible = false;

                    if (comboBox2.Text == dataGridView1[7, i].Value.ToString() && comboBox6.Text == dataGridView1[16, i].Value.ToString())
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
            //4
             if ( checkBox3.Checked == true && checkBox4.Checked == true )
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    dataGridView1.CurrentCell = null;
                    dataGridView1.Rows[i].Visible = false;

                    if (comboBox7.Text == dataGridView1[14, i].Value.ToString() && comboBox8.Text == dataGridView1[15, i].Value.ToString() )                    {
                        dataGridView1.Rows[i].Visible = true;
                    }
                    else
                    {
                        dataGridView1.Rows[i].Visible = false;
                    }

                }

            }
            else
            //5
             if ( checkBox3.Checked == true  && checkBox5.Checked == true)
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    dataGridView1.CurrentCell = null;
                    dataGridView1.Rows[i].Visible = false;

                    if ( comboBox7.Text == dataGridView1[14, i].Value.ToString() && comboBox6.Text == dataGridView1[16, i].Value.ToString())
                    {
                        dataGridView1.Rows[i].Visible = true;
                    }
                    else
                    {
                        dataGridView1.Rows[i].Visible = false;
                    }

                }

            }else
            //6
             if ( checkBox4.Checked == true && checkBox5.Checked == true)
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    dataGridView1.CurrentCell = null;
                    dataGridView1.Rows[i].Visible = false;

                    if (comboBox8.Text == dataGridView1[15, i].Value.ToString() && comboBox6.Text == dataGridView1[16, i].Value.ToString())
                    {
                        dataGridView1.Rows[i].Visible = true;
                    }
                    else
                    {
                        dataGridView1.Rows[i].Visible = false;
                    }

                }

            }else
            ///////////////// 

            //// 4 по 1 ///// 

            //1
            if (checkBox1.Checked == true)
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    dataGridView1.CurrentCell = null;
                    dataGridView1.Rows[i].Visible = false;

                    if (comboBox2.Text == dataGridView1[7, i].Value.ToString())
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
           //2
           if ( checkBox3.Checked == true )
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    dataGridView1.CurrentCell = null;
                    dataGridView1.Rows[i].Visible = false;

                    if (comboBox7.Text == dataGridView1[14, i].Value.ToString())
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
            //3
           if ( checkBox4.Checked == true )
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    dataGridView1.CurrentCell = null;
                    dataGridView1.Rows[i].Visible = false;

                    if (comboBox8.Text == dataGridView1[15, i].Value.ToString())
                    {
                        dataGridView1.Rows[i].Visible = true;
                    }
                    else
                    {
                        dataGridView1.Rows[i].Visible = false;
                    }

                }

            }else
                 //4
                 if (checkBox5.Checked == true)
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    dataGridView1.CurrentCell = null;
                    dataGridView1.Rows[i].Visible = false;

                    if (comboBox6.Text == dataGridView1[16, i].Value.ToString())
                    {
                        dataGridView1.Rows[i].Visible = true;
                    }
                    else
                    {
                        dataGridView1.Rows[i].Visible = false;
                    }

                }

            }
            /////
            else if (checkBox1.Checked==false&& checkBox3.Checked == false && checkBox4.Checked == false&&checkBox5.Checked == false )
            {
                Policyholder_load();
            }

            ///

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

        private void maskedTextBox1_MaskChanged(object sender, EventArgs e)
        {
           
        }

        private void maskedTextBox1_TextChanged(object sender, EventArgs e)
        {
         String s = maskedTextBox1.Text;
            String[] words = s.Split(' ');
            count2 = words.Length;
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (panel2.Visible == true)
            {
                panel4.Visible = true;
                panel3.Visible = false;
                panel5.Visible = false;
                panel6.Visible = false;
                panel2.Visible = false;
                button6.Left = 805;
                button6.Top = 732;
                button8.Visible = true;
                button8.Left = 341;
                button8.Top = 732;
            }
            else if(panel4.Visible == true) 
            {
                panel4.Visible = false;
                panel3.Visible = false;
                panel5.Visible = true;
                panel6.Visible = false;
                panel2.Visible = false;
                button6.Left = 744;
                button6.Top = 719;
                button8.Left = 403;
                button8.Top = 719;
            }
            else if (panel5.Visible == true)
            {
                panel4.Visible = false;
                panel3.Visible = false;
                panel6.Visible = true;
                panel5.Visible = false;
                panel2.Visible = false;
                button6.Visible = false;
                if(button11.Text== "Добавить")
                {
                button11.Width = 174;
                button11.Visible = true;
                button11.Left = 663;
                button11.Top = 707;
                button8.Left = 472;
                button8.Top = 707;
                }
                else
                {
                    button11.Width = 241;
                    button11.Visible = true;
                    button11.Left = 658;
                    button11.Top = 707;
                    button8.Left = 411;
                    button8.Top = 707;
                }

            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if(panel4.Visible == true)
            {
                panel4.Visible = false;
                panel3.Visible = false;
                panel5.Visible = false;
                panel6.Visible = false;
                panel2.Visible = true;
                button8.Visible = false; 
                button6.BringToFront();
                button6.Left = 909;
                button6.Top = 687;
            }
            else if(panel5.Visible == true)
            {
                panel4.Visible = true;
                panel3.Visible = false;
                panel5.Visible = false;
                panel6.Visible = false;
                panel2.Visible = false;
                button6.Left = 805;
                button6.Top = 732;
                button8.Left = 341;
                button8.Top = 732;
            }
            else if (panel6.Visible == true)
            {
                panel4.Visible = false;
                panel3.Visible = false;
                panel5.Visible = true;
                panel6.Visible = false;
                panel2.Visible = false;
                button11.Visible = false;
                button6.Visible = true;
          
                button6.Left = 744;
                button6.Top = 719;
                button8.Left = 403;
                button8.Top = 719;
            }
        }

        private void textBox10_KeyPress(object sender, KeyPressEventArgs e)
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

        private void textBox12_KeyPress(object sender, KeyPressEventArgs e)
        {char c = e.KeyChar;
            e.Handled = !((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || (c >= '0' && c <= '9') || (c == '@' || c == 8 || c == '.' || c == '_' || c == ',' || c == '#' || c == '+' || c == '!' || c == '$' || c == ':' || c == ';' || c == '%' || c == '^' || c == '&' || c == '*' || c == ')' || c == '(' || c == '-'));

           
        }

        private void textBox11_KeyPress(object sender, KeyPressEventArgs e)
        {
 char c = e.KeyChar;
            e.Handled = !((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || (c >= '0' && c <= '9') || (c == '@' || c == 8 || c == '.' || c == '_' || c == ','));

        }

        private void comboBox3_SelectedValueChanged(object sender, EventArgs e)
        {
           
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox3.Text!="System.Data.DataRowView"&& comboBox3.Text != "")
            {
                comboBox4.Visible = true;
                comboBoxposition();
            }
        }

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {
            
                
            
        }
    }
}
