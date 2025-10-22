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
            SqlDataAdapter command = new SqlDataAdapter($@"Select idtenant,firstname as Фамилия,name as Имя, lastname as Отчество,dateb as [Дата рождения], phone as Телефон,email as Почта,login as Логин,password as Пароль from tenant", sqlConnection);
            command.Fill(dataSet);
            dataGridView1.DataSource = dataSet.Tables[0].DefaultView;
            sqlConnection.Close();
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.AllowUserToAddRows = false;
            sqlConnection.Close();
        }
        
        
        int k = 0;
        int j = 0;
       
        public void clear()
        {
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            maskedTextBox2.Text = "";
            textBox10.Text = "";
            textBox12.Text = "";
            textBox11.Text = "";
        }
        private void button11_Click(object sender, EventArgs e)
        {
            if (label10.Text == "Регистрация")
            {

                k = 0;
                if (textBox2.Text != "" && textBox3.Text != "" && textBox4.Text != "" && textBox12.Text != "" && textBox10.Text != "" && textBox11.Text != ""  && maskedTextBox2.Text.Length == 18)
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
                                if (maskedTextBox2.Text == dataGridView1[5, i].Value.ToString().ToLower())
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
                                        if (textBox10.Text == dataGridView1[6, i].Value.ToString())
                                        {
                                            k++;
                                        }
                                    }
                                    if (k == 0)
                                    {
                                        if (textBox11.Text.Length > 5 && textBox11.Text.Any(char.IsLetter))

                                        {
                                            k = 0;
                                            for (int i = 0; i < dataGridView1.RowCount; i++)
                                            {
                                                if (textBox11.Text == dataGridView1[7, i].Value.ToString())
                                                {
                                                    k++;
                                                }
                                            }
                                            if (k == 0)
                                            {
                                                if (textBox12.Text.Any(char.IsLower) && textBox12.Text.Any(char.IsUpper) && textBox12.Text.Length > 8)
                                                {

                                                    sqlConnection.Open();
                                                    SqlCommand command = new SqlCommand($@"INSERT INTO [tenant](firstname,name,lastname,phone,dateb,email,login,password) VALUES (@f,@n,@l,@h,@b,@e,@lo,@pa);", sqlConnection);
                                                    command.Parameters.AddWithValue("@f", (textBox2.Text));
                                                    command.Parameters.AddWithValue("@n", (textBox3.Text));
                                                    command.Parameters.AddWithValue("@l", (textBox4.Text));

                                                    command.Parameters.AddWithValue("@h", (maskedTextBox2.Text));

                                                    command.Parameters.AddWithValue("@b", (dateTimePicker2.Value));

                                                    command.Parameters.AddWithValue("@e", (textBox10.Text));


                                                    command.Parameters.AddWithValue("@lo", (textBox11.Text));
                                                    command.Parameters.AddWithValue("@pa", (textBox12.Text));
                                                    command.ExecuteNonQuery();
                                                    sqlConnection.Close();
                                                    Policyholder_load();
                                                    clear();
                                                    panel1.Visible = false;
                                                   
                                                    Policyholderak.ActiveForm.StartPosition = FormStartPosition.CenterScreen;
                                                    try
                                                    {
                                                        string query2 = $@" Select max(idtenant) from tenant";
                                                        System.Data.DataTable data2 = new System.Data.DataTable();
                                                        SqlDataAdapter command2 = new SqlDataAdapter(query2, sqlConnection);
                                                        command2.Fill(data2);
                                                        DataColumn column2 = data2.Columns[0];
                                                        DataRow row2 = data2.Rows[0];
                                                        idakk = Convert.ToInt32(row2[column2].ToString());
                                                    }
                                                    catch { }
                                                    Form.Working uc = new Form.Working(this);

                                                    panel14.Left = 204;
                                                    panel14.Top = button16.Top;
                                                    panel14.Height = button16.Height;
                                                    uc.dataGridView1.Visible = true;
                                                    uc.dataGridView1.RowsDefaultCellStyle.WrapMode = DataGridViewTriState.True;
                                                    uc.dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                                                    uc.dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                                                    addControll(uc);
                                                    panel6.Visible = true; panel11.Visible = true; panel14.Visible = true;
                                                    panel14.BringToFront();
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
                        MessageBox.Show("Такой страхователь уже есть!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                if (textBox2.Text != "" && textBox3.Text != "" && textBox4.Text != "" && textBox12.Text != "" && textBox10.Text != "" && textBox11.Text != "" && maskedTextBox2.Text.Length == 18)
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
                                if (maskedTextBox2.Text == dataGridView1[5, i].Value.ToString().ToLower())
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
                                        if (textBox10.Text == dataGridView1[6, i].Value.ToString())
                                        {
                                            k++;
                                            j = Convert.ToInt32(dataGridView1[0, i].Value.ToString());
                                        }
                                    }
                                    if (k == 0 || j == startcs.idak)
                                    {
                                        if (textBox11.Text.Length > 5 && textBox11.Text.Any(char.IsLetter))

                                        {
                                            k = 0; j = 0;
                                            for (int i = 0; i < dataGridView1.RowCount; i++)
                                            {
                                                if (textBox11.Text == dataGridView1[7, i].Value.ToString())
                                                {
                                                    k++;
                                                    j = Convert.ToInt32(dataGridView1[0, i].Value.ToString());
                                                }
                                            }
                                            if (k == 0 || j == startcs.idak)
                                            {
                                                if (textBox12.Text.Any(char.IsLower) && textBox12.Text.Any(char.IsUpper) && textBox12.Text.Length > 8)
                                                {


                                                    sqlConnection.Open();
                                                    SqlCommand command = new SqlCommand($@"UPDATE tenant SET firstname=@f,name=@n ," +
                                                    $"lastname=@l,phone=@h,dateb=@b,login=@lo,password=@pa,email=@e WHERE idtenant=@id", sqlConnection);
                                                    command.Parameters.AddWithValue("@f", (textBox2.Text));
                                                    command.Parameters.AddWithValue("@n", (textBox3.Text));
                                                    command.Parameters.AddWithValue("@l", (textBox4.Text));

                                                    command.Parameters.AddWithValue("@h", (maskedTextBox2.Text));

                                                    command.Parameters.AddWithValue("@b", (dateTimePicker2.Value));

                                                    command.Parameters.AddWithValue("@e", (textBox10.Text));

                                                    command.Parameters.AddWithValue("@lo", (textBox11.Text));
                                                    command.Parameters.AddWithValue("@pa", (textBox12.Text));
                                                    command.Parameters.AddWithValue("@id", (idakk));
                                                    command.ExecuteNonQuery();
                                                    sqlConnection.Close();
                                                    dataGridView1.Enabled = true;
                                                    clear();
                                                    Policyholder_load();
                                                    panel1.Visible = false;
                                                  
                                                    Form.Working uc = new Form.Working(this);
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
                        MessageBox.Show("Такой страхователь уже есть!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                Form.Working uc = new Form.Working(this);
                panel14.Top = button16.Top;
                panel14.Left = 204;
                panel14.Height = button16.Height;
                addControll(uc);
            }
            Policyholder_load();
           
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

        private void maskedTextBox2_TextChanged(object sender, EventArgs e)
        {
          
        }
        string city = "";
        

        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            e.Handled = !((c >= 'а' && c <= 'я') || (c >= 'А' && c <= 'Я') || (c == 'ё' || c == 8 || c == 32 || c == ',' || c == '.' || c == '0' || c == '1' || c == '2' || c == '3' || c == '4' || c == '5' || c == '6' || c == '7' || c == '8' || c == '9'));
        }

       

        private void textBox6_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && e.KeyChar != 8)
                e.Handled = true;
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

        int prod = 0;
        int idp = 0;

        private void button6_Click(object sender, EventArgs e)
        {

            panel1.Visible = false;
            panel6.Visible = true;
            Form.Back uc = new Form.Back(this);

            panel14.Left = 204;
            panel14.Top = button6.Top;
            panel14.Height = button6.Height;
           
            addControll(uc);
        }
       
        int hall = 0;
        string hyll = "";

        private void button7_Click(object sender, EventArgs e)
        {
            panel1.Visible = false;
            panel6.Visible = true;
            Form.Pays uc = new Form.Pays(this);
            panel14.Top = button7.Top;
            panel14.Left = 204;
            panel14.Height = button7.Height;
            addControll(uc);
        }

        private void button16_Click(object sender, EventArgs e)
        {

            panel1.Visible = false;
            panel6.Visible = true;
            Form.Bid uc = new Form.Bid(this);

            panel14.Left = 204;
            panel14.Top = button6.Top;
            panel14.Height = button6.Height;
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
            string query2 = $@"select ROW_NUMBER() over (ORDER BY idtenant) num,idtenant
from tenant
group by idtenant";
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
                if (hall == idakk)
                {
                    idp = Convert.ToInt32(row22[column22].ToString()) - 1;
                    break;
                }
            }
            prod = 0;
            textBox2.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            textBox3.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            textBox4.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();

            maskedTextBox2.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
            dateTimePicker2.Value = Convert.ToDateTime(dataGridView1.CurrentRow.Cells[4].Value.ToString());

            textBox11.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString();
            textBox12.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString();
            textBox10.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();

            label10.Text = "Редактировать профиль";
            button11.Text = "Редактировать";
            panel1.Left = 407;
            panel1.Top = 60;
        }

        private void button12_Click(object sender, EventArgs e)
        {
            
            panel1.Visible = false;
            panel6.Visible = true;
            
            Form.Object uc = new Form.Object(this);
            panel14.Top = button12.Top;
            panel14.Left = 204;
            panel14.Height = button12.Height;
            addControll(uc);
        }

        private void Policyholderak_FormClosed(object sender, FormClosedEventArgs e)
        {
           
        }

        private void panel6_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            panel1.Visible = false;
            panel6.Visible = true;

            Form.Service uc = new Form.Service(this);
            panel14.Top = button3.Top;
            panel14.Left = 204;
            panel14.Height = button3.Height;
            addControll(uc);
        }
    }
}
