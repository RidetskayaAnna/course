using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Data.SqlClient;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Agent.Form;

namespace Agent
{
    public partial class Policyholderak : System.Windows.Forms.Form
    {
        datebase db = new datebase();
        public SqlConnection sqlConnection = null;
        Startcs startcs;
        public Policyholderak(Startcs form2)
        {
            InitializeComponent();
            sqlConnection = new SqlConnection(db.connection);
            startcs = form2;
        }
       public int idakk = 0;
        public Boolean press = true;
        private System.Drawing.Point mouseOffset;
        private bool isMouseDown = false;
        private void Form1_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) this.Close();
        }
        int id2 = 0;
        private void Form1_MouseDown(object sender, MouseEventArgs e)
        {
            int xOffset;
            int yOffset;

            if (e.Button == MouseButtons.Left)
            {
                xOffset = -e.X - SystemInformation.FrameBorderSize.Width;
                yOffset = -e.Y - SystemInformation.CaptionHeight -
                    SystemInformation.FrameBorderSize.Height;
                mouseOffset = new System.Drawing.Point(xOffset, yOffset);
                isMouseDown = true;
            }
        }

        private void Form1_MouseMove(object sender, MouseEventArgs e)
        {
            if (isMouseDown)
            {
                System.Drawing.Point mousePos = Control.MousePosition;
                mousePos.Offset(mouseOffset.X, mouseOffset.Y);
                Location = mousePos;
            }
        }

        private void Form1_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                isMouseDown = false;
            }
        }
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
        int k = 0;
        int j = 0;
        int count = 0;
        int count2 = 0;
        String passport = "";
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
            dateTimePicker1.MinDate = DateTime.Today.AddDays(-3651);
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
        private void button11_Click(object sender, EventArgs e)
        {
            if (label10.Text == "Регистрация")
            {
                if (textBox2.Text != "" && textBox3.Text != "" && textBox4.Text != "" && textBox12.Text != "" && textBox10.Text != "" && textBox11.Text != "" && textBox5.Text != "" && textBox6.Text != "" && textBox7.Text != "" && textBox6.Text.Length == 7 && count == 2 && count2 == 1 && maskedTextBox2.Text.Length == 18 && maskedTextBox1.Text.Length == 14 && comboBox4.SelectedIndex != -1 && comboBox5.SelectedIndex != -1)
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
                                                    command.ExecuteNonQuery();
                                                    sqlConnection.Close();
                                                    Policyholder_load();
                                                    clear();
                                                    panel1.Visible = false;
                                                    Width = 1551;
                                                    Height = 638;
                                                    Policyholderak.ActiveForm.StartPosition= FormStartPosition.CenterScreen;
                                                    //try
                                                    //{
                                                        string query2 = $@" Select max(idpolicyholder) from policyholder";
                                                        System.Data.DataTable data2 = new System.Data.DataTable();
                                                        SqlDataAdapter command2 = new SqlDataAdapter(query2, sqlConnection);
                                                        command2.Fill(data2);
                                                        DataColumn column2 = data2.Columns[0];
                                                        DataRow row2 = data2.Rows[0];
                                                        idakk =Convert.ToInt32 (row2[column2].ToString());
                                                    //}
                                                    //catch { }
                                                    Form.Treaty uc = new Form.Treaty(this);
                                                    
                                                    panel14.Left = 204;
                                                    panel14.Top = button16.Top;
                                                    panel14.Height = button16.Height;
                                                    uc.dataGridView1.Visible = true;
                                                    uc.dataGridView1.RowsDefaultCellStyle.WrapMode = DataGridViewTriState.True;
                                                    uc.dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                                                    uc.dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                                                    addControll(uc);
                                                    panel6.Visible = true;panel11.Visible = true;panel14.Visible = true;
                                                    panel14.BringToFront();
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
                        if (k == 0 || j == startcs.idak)
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

                            if (k == 0 || j == startcs.idak)
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
                                if (k == 0 || j == startcs.idak)
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
                                        if (k == 0 || j == startcs.idak)
                                        {
                                            k = 0; j = 0;
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
                                                    command.Parameters.AddWithValue("@id", (idakk));
                                                    command.ExecuteNonQuery();
                                                    sqlConnection.Close();
                                                    dataGridView1.Enabled = true;
                                                    clear();
                                                    Policyholder_load();
                                                    panel1.Visible = false;
                                                    panel6.Visible = true;
                                                    Form.Treaty uc = new Form.Treaty(this);
                                                    Width = 1551;
                                                    Height = 638;
                                                    panel14.Left = 204;
                                                    panel14.Top = button16.Top;
                                                    panel14.Height = button16.Height;
                                                    uc.dataGridView1.Visible = true;
                                                    uc.dataGridView1.RowsDefaultCellStyle.WrapMode = DataGridViewTriState.True;
                                                    uc.dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                                                    uc.dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                                                    addControll(uc);

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
        private void addControll(UserControl uc)
        {

            panel6.Controls.Clear();
            panel6.Controls.Add(uc);

        }
      
        private void Policyholderak_Load(object sender, EventArgs e)
        {
            if (startcs.idak != 0)
            {
            
                idakk = startcs.idak;
                Form.Treaty uc = new Form.Treaty(this);
                panel14.Top = button16.Top;
                panel14.Left = 204;
                panel14.Height = button16.Height;
                addControll(uc);
            }
            Policyholder_load();
         comboBoxx();
              
                comboBoxwork();
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

        private void maskedTextBox2_MaskChanged(object sender, EventArgs e)
        {

        }

        private void maskedTextBox2_TextChanged(object sender, EventArgs e)
        {
            String s = maskedTextBox2.Text;
            String[] words = s.Split(' ');
            count = words.Length;
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

        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            e.Handled = !((c >= 'а' && c <= 'я') || (c >= 'А' && c <= 'Я') || (c == 'ё' || c == 8 || c == 32 || c == ',' || c == '.' || c == '0' || c == '1' || c == '2' || c == '3' || c == '4' || c == '5' || c == '6' || c == '7' || c == '8' || c == '9'));
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            passport = textBox7.Text + textBox6.Text;
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            passport = textBox7.Text + textBox6.Text;
        }

        private void textBox6_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && e.KeyChar != 8)
                e.Handled = true;
        }

        private void maskedTextBox1_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void maskedTextBox1_TextChanged(object sender, EventArgs e)
        {
            String s = maskedTextBox1.Text;
            String[] words = s.Split(' ');
            count2 = words.Length;
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox3.Text != "System.Data.DataRowView" && comboBox3.Text != "")
            {
                comboBox4.Visible = true;
                comboBoxposition();
            }
        }

        private void textBox11_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox11_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            e.Handled = !((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || (c >= '0' && c <= '9') || (c == '@' || c == 8 || c == '.' || c == '_' || c == ','));

        }

        private void textBox12_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            e.Handled = !((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || (c >= '0' && c <= '9') || (c == '@' || c == 8 || c == '.' || c == '_' || c == ',' || c == '#' || c == '+' || c == '!' || c == '$' || c == ':' || c == ';' || c == '%' || c == '^' || c == '&' || c == '*' || c == ')' || c == '(' || c == '-'));
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
        public void comboBoxvid()
        {

        }
        private void button4_Click(object sender, EventArgs e)
        {
           
        }
        public void Vid_load()
        {
           
        }
        private void button3_Click(object sender, EventArgs e)
        {
            
        }
        private void button2_Click(object sender, EventArgs e)
        {
            
        }
        int prod = 0;
        int idp = 0;
        private void button1_Click(object sender, EventArgs e)
        {
           


        }

        private void panel6_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
          
        }
        public void Objectpolicyholder_load()
        {
       
        }
        private void button6_Click(object sender, EventArgs e)
        {
            panel1.Visible = false;
            panel6.Visible = true;
            Form.Bid uc = new Form.Bid(this);
            Width = 1551;
            Height = 638;
            panel14.Left = 204;
            panel14.Top = button6.Top;
            panel14.Height = button6.Height;
            uc.dataGridView1.Visible = true;
            uc.dataGridView1.RowsDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            uc.dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            uc.dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            addControll(uc);
        }
        public void Pays_load()
        {
        }
        public void comboBoxtreaty()
        {
        }
        Decimal kk = 0;
        int hall = 0;
        string hyll = "";
        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
        }
        private void button7_Click(object sender, EventArgs e)
        {
            panel1.Visible = false;
            panel6.Visible=true;
            Form.Pays uc = new Form.Pays(this);
            panel14.Top = button7.Top;
            panel14.Left = 204;
            panel14.Height = button7.Height;
            addControll(uc);
        }

        private void comboBox6_SelectedValueChanged(object sender, EventArgs e)
        {
           
        }

        private void button8_Click(object sender, EventArgs e)
        {
           
        }

        private void button16_Click(object sender, EventArgs e)
        {
            panel1.Visible = false;
            panel6.Visible = true;
            Form.Treaty uc = new Form.Treaty(this);
            Width = 1551;
            Height = 638;
            panel14.Left = 204;
            panel14.Top = button16.Top;
            panel14.Height = button16.Height;
            uc.dataGridView1.Visible = true;
            uc.dataGridView1.RowsDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            uc.dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            uc.dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            addControll(uc);
        }

        private void button14_Click(object sender, EventArgs e)
        {
            panel14.Left = 204;
            panel14.Top = button14.Top;
            panel14.Height = button14.Height;
            panel6.Visible = false;
            panel1.Visible = true;
            panel1.BringToFront();
            idp = 0;
            string query2 = $@"select ROW_NUMBER() over (ORDER BY idpolicyholder) num,idpolicyholder
from policyholder
group by idpolicyholder";
            System.Data.DataTable data2 = new System.Data.DataTable();
            SqlDataAdapter command2 = new SqlDataAdapter(query2, sqlConnection);
            command2.Fill(data2);

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                DataColumn column2 = data2.Columns[1];
                DataRow row2 = data2.Rows[i];
                hall = Convert.ToInt32(row2[column2].ToString());
                DataColumn column22 = data2.Columns[0];
                DataRow row22 = data2.Rows[i];
                if (hall ==idakk)
                {
                    idp = Convert.ToInt32(row22[column22].ToString()) - 1;
                    break;
                }
            }
            prod = 0;
            textBox2.Text = dataGridView1[2, idp].Value.ToString();
            textBox3.Text = dataGridView1[3, idp].Value.ToString();
            textBox4.Text = dataGridView1[4, idp].Value.ToString();
            textBox5.Text = dataGridView1[8, idp].Value.ToString();
            maskedTextBox1.Text = dataGridView1[10, idp].Value.ToString();
            char[] o = { 'Р', 'О', 'В', 'Д' };
            prod = dataGridView1[11, idp].Value.ToString().LastIndexOfAny(o);
            dateTimePicker1.Value = Convert.ToDateTime(dataGridView1[12, idp].Value.ToString());
            textBox8.Text = dataGridView1[11, idp].Value.ToString().Substring(0, prod - 4);
            textBox9.Text = dataGridView1[11, idp].Value.ToString().Substring(prod - 4, 4);
            textBox6.Text = dataGridView1[9, idp].Value.ToString().Substring(2, 7);
            comboBox1.SelectedValue = dataGridView1[6, idp].Value.ToString();
            maskedTextBox2.Text = dataGridView1[13, idp].Value.ToString();
            dateTimePicker2.Value = Convert.ToDateTime(dataGridView1[5, idp].Value.ToString());
            comboBox3.SelectedValue = dataGridView1[21, idp].Value.ToString();
            comboBox4.SelectedValue = Convert.ToInt32(dataGridView1[1, idp].Value.ToString());
            comboBox5.SelectedItem = dataGridView1[16, idp].Value.ToString();

            if (dataGridView1[17, idp].Value.ToString() == "Да")
            {
                checkBox2.Checked = true;
            }
            else
            {
                checkBox2.Checked = false;
            }
            textBox11.Text = dataGridView1[19, idp].Value.ToString();
            textBox12.Text = dataGridView1[20, idp].Value.ToString();
            textBox10.Text = dataGridView1[18, idp].Value.ToString();


            metroSetControlBox2.Visible = false;
            label10.Text = "Редактировать профиль";
            button11.Text = "Редактировать";
            panel1.Left = 307;
            panel1.Top = 60;
        }

        private void button12_Click(object sender, EventArgs e)
        {
            panel1.Visible = false;
            panel6.Visible = true;
            Width = 1551;
            Height = 638;
            Form.Vid uc = new Form.Vid(this);
            panel14.Top = button12.Top;
            panel14.Left = 204;
            panel14.Height = button12.Height;
            addControll(uc);
        }
    }
}
