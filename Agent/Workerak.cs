using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Text.RegularExpressions;
using Agent.Form;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ToolBar;

namespace Agent
{
    public partial class Workerak : System.Windows.Forms.Form
    {
        datebase db = new datebase();
        public SqlConnection sqlConnection = null;
        Startcs startcs;
        public Workerak(Startcs form2)
        {
            InitializeComponent();
            sqlConnection = new SqlConnection(db.connection);
            startcs = form2;
        }
        public Boolean press = true;
        private System.Drawing.Point mouseOffset;
        private bool isMouseDown = false;
        private void Form1_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) this.Close();
        }
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
inner join work on post.idwork=work.idwork where idworker={idakk}", sqlConnection);
            command.Fill(dataSet);
            dataGridView1.DataSource = dataSet.Tables[0].DefaultView;
            sqlConnection.Close();
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[6].Visible = false;
            dataGridView1.Columns[8].Visible = false;
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            sqlConnection.Close();
        }
        public void comboBoxwork()
        {
            try
            {
                if (idakk != null)
                {
                    sqlConnection.Close();
                    sqlConnection.Open();
                    string query = $@"select work.idwork, work.name 
from work inner join post on post.idwork=work.idwork 
inner join worker on worker.idpost=post.idpost
where worker.idworker={idakk}";
                    SqlDataAdapter sqlDbDataAdapter = new SqlDataAdapter(query, sqlConnection);
                    DataSet dataSet = new DataSet();
                    sqlDbDataAdapter.Fill(dataSet);
                    comboBox3.DataSource = dataSet.Tables[0];
                    comboBox3.DisplayMember = "name";
                    comboBox3.ValueMember = "idwork";
                    comboBox3.SelectedIndex = -1;
                    sqlConnection.Close();
                }
                else
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
            }
            catch { }

        }
        public void comboBoxposition()
        {
            try
            {
                if (idakk != null)
                {
                    sqlConnection.Close();
                    sqlConnection.Open();
                    string query = $@"Select post.idpost, post.name as Должность from post inner join work on post.idwork=work.idwork 
inner join worker on worker.idpost=post.idpost where worker.idworker={idakk}";
                    SqlDataAdapter sqlDbDataAdapter = new SqlDataAdapter(query, sqlConnection);
                    DataSet dataSet = new DataSet();
                    sqlDbDataAdapter.Fill(dataSet);
                    comboBox4.DataSource = dataSet.Tables[0];
                    comboBox4.DisplayMember = "Должность";
                    comboBox4.ValueMember = "idpost";
                    comboBox4.SelectedIndex = -1;
                    sqlConnection.Close();
                }
                else {
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
            }
            catch { }
        }
        public int idakk = 0;
        private void Insurerak_Load(object sender, EventArgs e)
        {
            if (startcs.idak != 0)
            {
                idakk = startcs.idak;
                Form.Working uc = new Form.Working(this);
                panel6.Top = button7.Top;
                panel6.Left = 204;
                panel6.Height = button7.Height;
               
                addControll(uc); 
                button6.BringToFront();
                button6.Visible = false;
                button10.BringToFront();
                button10.Visible = true;
                button9.BringToFront();
                button9.Visible = false;
            }
            Insurer_load();
            dateTimePicker1.MinDate = Convert.ToDateTime("01.01.1980");
            dateTimePicker1.MaxDate = DateTime.Today;
        }
        private void addControll(UserControl uc)
        {
            panel5.Controls.Clear();
            panel5.Controls.Add(uc);

        }
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (((System.Windows.Forms.TextBox)sender).Text.Length == 1)
                ((System.Windows.Forms.TextBox)sender).Text = ((System.Windows.Forms.TextBox)sender).Text.ToUpper();
            ((System.Windows.Forms.TextBox)sender).Select(((System.Windows.Forms.TextBox)sender).Text.Length, 0);
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
        }
        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            e.Handled = !((c >= 'а' && c <= 'я') || (c >= 'А' && c <= 'Я') || (c == 'ё' || c == 8 || c == 32));
        }

        private void textBox7_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            e.Handled = !((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || (c == '@' || c == 8 || c == '.'));
            if (e.KeyChar == '@')
            {
                if (((sender as System.Windows.Forms.TextBox).Text.IndexOf('@') != -1))
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
        int k = 0;
        int j = 0;
        int count = 0;
        private void button11_Click(object sender, EventArgs e)
        {
            try
            {
                if (label6.Text == "Регистрация")
                {
                    k = 0;
                    if (textBox2.Text != "" && textBox3.Text != "" && textBox4.Text != "" && textBox5.Text != "" && count == 2 && textBox6.Text != "" && textBox7.Text != "" && maskedTextBox2.Text.Length == 18 && comboBox4.SelectedIndex != -1)
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

                                                        clear();
                                                        panel2.Visible = false;
                                                        panel1.Visible = false;
                                                        panel2.Visible = true;
                                                        panel5.Visible = true;
                                                        string query2 = $@" Select max(idworker) from worker";
                                                        System.Data.DataTable data2 = new System.Data.DataTable();
                                                        SqlDataAdapter command2 = new SqlDataAdapter(query2, sqlConnection);
                                                        command2.Fill(data2);
                                                        DataColumn column2 = data2.Columns[0];
                                                        DataRow row2 = data2.Rows[0];
                                                        idakk = Convert.ToInt32(row2[column2].ToString());
                                                        Form.Working uc = new Form.Working(this);
                                                        panel6.BringToFront();
                                                        panel6.Visible = true;
                                                        panel6.Top = button7.Top;
                                                        panel6.Left = 204;
                                                        panel6.Height = button7.Height;
                                                        addControll(uc);

                                                        metroSetControlBox2.Visible = true;
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
                    k = 0; j = 0;
                    if (textBox2.Text != "" && textBox3.Text != "" && textBox4.Text != "" && textBox5.Text != "" && textBox6.Text != "" && textBox7.Text != "" && count == 2 && maskedTextBox2.Text.Length == 18 && comboBox4.SelectedIndex != -1)
                    {
                        for (int i = 0; i < dataGridView1.RowCount; i++)
                        {
                            if (textBox2.Text.ToLower() + textBox3.Text.ToLower() + textBox4.Text.ToLower() == dataGridView1[1, i].Value.ToString().ToLower() + dataGridView1[2, i].Value.ToString().ToLower() + dataGridView1[3, i].Value.ToString().ToLower())
                            {
                                k++;
                                j = Convert.ToInt32(dataGridView1[0, i].Value.ToString());
                            }
                        }
                        if (k == 0 || j == startcs.idak)
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
                                if (k == 0 || j == startcs.idak)
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
                                        if (k == 0 || j == startcs.idak)
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
                                                if (k == 0 || j == startcs.idak)
                                                {
                                                    k = 0; j = 0;
                                                    if (textBox6.Text.Any(char.IsLower) && textBox6.Text.Any(char.IsUpper) && textBox6.Text.Length > 8)
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
                                                        command.Parameters.AddWithValue("@id", (idakk));
                                                        command.ExecuteNonQuery();
                                                        sqlConnection.Close();
                                                        dataGridView1.Enabled = true;


                                                        
                                                        Insurer_load();
                                                        
                                                        MessageBox.Show("Изменения сохранены", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            }
            catch { }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            button9.Visible = false;
            button10.Visible = false;
            panel6.BringToFront();
            panel6.Visible = true;
            panel1.Visible = false;
            panel5.Visible = true;
            Form.Working uc = new Form.Working(this);
          
            panel6.Left = 204;
            panel6.Top = button7.Top;
            panel6.Height = button7.Height;
            uc.dataGridView1.Visible = true;
            uc.dataGridView1.RowsDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            uc.dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            addControll(uc); 
            button6.Visible = false;
            button6.BringToFront();
            button9.Visible = false;
            button9.BringToFront();
            button10.Visible = true;
            button10.BringToFront();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            panel6.BringToFront();
            panel6.Visible = true;
            panel1.Visible = false;
            panel5.Visible = true;
            button6.Visible = false;
            button9.Visible = false;
            button10.Visible = false;
            Form.Object uc = new Form.Object(this);
            panel6.Top = button1.Top;
            panel6.Left = 204;
            panel6.Height = button1.Height;
            addControll(uc);
        }

       

        private void button3_Click(object sender, EventArgs e)
        {
            panel6.BringToFront();
            panel6.Visible = true;
            panel1.Visible = false;
            panel5.Visible = true;
            button6.Visible = false;
            button9.Visible = false;
            button10.Visible = false;
            Form.Service uc = new Form.Service(this);
            panel6.Top = button3.Top;
            panel6.Left = 204;
            panel6.Height = button3.Height;
            addControll(uc);
        }

      

        private void button5_Click(object sender, EventArgs e)
        {
          
           
            panel6.BringToFront();
            panel6.Visible = true;
            panel1.Visible = false;
            panel5.Visible = true; ;
            button6.Visible = false;
            button9.Visible = false;
            button10.Visible = false;
            Notification uc = new Notification(this);
            panel6.Top = button5.Top;
            panel6.Left = 204;
            panel6.Height = button5.Height;
            addControll(uc);
        }
        
        private void button2_Click(object sender, EventArgs e)
        {
            button6.Visible = false;
            button9.Visible = false;
            button10.Visible = false;

            panel6.BringToFront();
            panel6.Visible = true;
            panel5.Visible = false;
           
            metroSetControlBox1.Visible = false;
            panel1.Visible = true;
            panel1.Visible = true;
            panel1.BringToFront();

            textBox2.Text = dataGridView1[1, 0].Value.ToString();
            textBox3.Text = dataGridView1[2, 0].Value.ToString();
            textBox4.Text = dataGridView1[3, 0].Value.ToString();
            comboBox3.Text = dataGridView1[9, 0].Value.ToString();
            comboBox4.SelectedValue = Convert.ToInt32(dataGridView1[6, 0].Value.ToString());
            maskedTextBox2.Text = dataGridView1[10, 0].Value.ToString();
            textBox7.Text = dataGridView1[11, 0].Value.ToString();
            textBox5.Text = dataGridView1[12, 0].Value.ToString();
            textBox6.Text = dataGridView1[13, 0].Value.ToString();
            label6.Text = "Редактировать профиль";
            button11.Text = "Редактировать";
            panel1.Left = 228;
            panel1.Top = 117;
            panel6.Top = button2.Top;
            panel6.Left = 204;
            panel6.Height = button2.Height;
            comboBoxposition();
            comboBoxwork();
            comboBox4.Visible = true;
            comboBox3.SelectedIndex = 0;
            comboBox4.SelectedIndex = 0;
        }

        private void Insurerak_FormClosed(object sender, FormClosedEventArgs e)
        {
           
        }

     

        private void maskedTextBox2_TextChanged(object sender, EventArgs e)
        {
            String s = maskedTextBox2.Text;
            String[] words = s.Split(' ');
            count = words.Length;

        }

        private void textBox2_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            e.Handled = !((c >= 'а' && c <= 'я') || (c >= 'А' && c <= 'Я') || (c == 'ё' || c == 8 || c == 32));

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            if (((System.Windows.Forms.TextBox)sender).Text.Length == 1)
                ((System.Windows.Forms.TextBox)sender).Text = ((System.Windows.Forms.TextBox)sender).Text.ToUpper();
            ((System.Windows.Forms.TextBox)sender).Select(((System.Windows.Forms.TextBox)sender).Text.Length, 0);
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            e.Handled = !((c >= 'а' && c <= 'я') || (c >= 'А' && c <= 'Я') || (c == 'ё' || c == 8 || c == 32));

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

        private void textBox7_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            e.Handled = !((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || (c == '@' || c == 8 || c == '.'));
            if (e.KeyChar == '@')
            {
                if (((sender as System.Windows.Forms.TextBox).Text.IndexOf('@') != -1))
                {
                    e.Handled = true;
                    return;
                }
            }
        }

        private void textBox5_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            e.Handled = !((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || (c >= '0' && c <= '9') || (c == '@' || c == 8 || c == '.' || c == '_' || c == ','));

        }

        private void textBox6_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            e.Handled = !((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || (c >= '0' && c <= '9') || (c == '@' || c == 8 || c == '.' || c == '_' || c == ',' || c == '#' || c == '+' || c == '!' || c == '$' || c == ':' || c == ';' || c == '%' || c == '^' || c == '&' || c == '*' || c == ')' || c == '(' || c == '-'));

        }

        private void button10_Click(object sender, EventArgs e)
        {
            Form.Bid uc = new Form.Bid(this);
           
            button7.BringToFront();
            button7.Visible = true;
            button10.BringToFront();
            button10.Visible = true;
            button9.BringToFront();
            button9.Visible = true;
            panel6.Top = button10.Top;
            panel6.Left = 1205;
            panel6.Height = button10.Height;
            addControll(uc);
        }

        private void button9_Click_1(object sender, EventArgs e)
        {
            panel6.BringToFront();
            panel6.Visible = true;
            panel1.Visible = false;
            panel5.Visible = true;
            button7.Visible = true;
            button9.Visible = true;
           
            button10.Visible = true;
            Form.Back uc = new Form.Back(this);
            panel6.Top = button9.Top;
            panel6.Left = 1342;
            panel6.Height = button9.Height;
            panel6.BringToFront();
            addControll(uc);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Form.Pays uc = new Form.Pays(this);
            panel6.Top = button6.Top;
            panel6.Left = 1486;
            panel6.Height = button6.Height;
            button7.BringToFront();
            panel6.BringToFront();
            button7.Visible = true;
            button9.Visible = true;
         
            button10.Visible = true;
            addControll(uc);
        }

        private void button13_Click(object sender, EventArgs e)
        {
           

            panel6.BringToFront();
            panel6.Visible = true;
            panel1.Visible = false;
            panel5.Visible = true;
            button6.Visible = false;
            button9.Visible = false;
            button10.Visible = false;
            Form.Calendar uc = new Form.Calendar();
            panel6.Top = button13.Top;
            panel6.Left = 204;
            panel6.Height = button13.Height;
            addControll(uc);
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox3.Text != "System.Data.DataRowView" && comboBox3.Text != "")
            {
                comboBox4.Visible = true;
                comboBoxposition();
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Form.Static staticc = new Form.Static();
            addControll(staticc);
            panel4.Height = button6.Height;
            panel4.Top = button6.Top;
            panel4.Left = 204;
            button10.Visible = false;
           
            button7.Visible = false;
            button9.Visible = false;
        }
    }
}
