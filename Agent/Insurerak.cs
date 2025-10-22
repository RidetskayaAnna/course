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
    public partial class Insurerak : System.Windows.Forms.Form
    {
        datebase db = new datebase();
        public SqlConnection sqlConnection = null;
        Startcs startcs;
        public Insurerak(Startcs form2)
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
            SqlDataAdapter command = new SqlDataAdapter($@"Select idinsurer, firstname as Фамилия, name as Имя, lastname as Отчество, datereception as [Дата приема], datelayoffs as [Дата увольнения], email as Почта, login as Логин, password as Пароль from insurer", sqlConnection);
            command.Fill(dataSet);
            dataGridView1.DataSource = dataSet.Tables[0].DefaultView;
            sqlConnection.Close();
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.AllowUserToAddRows = false;
            sqlConnection.Close();
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
        private void button11_Click(object sender, EventArgs e)
        {
            try
            {
                if (label6.Text == "Регистрация")
                {
                    if (textBox2.Text != "" && textBox3.Text != "" && textBox4.Text != "" && textBox5.Text != "" && textBox6.Text != "" && textBox7.Text != "")
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
                            k = 0;
                            Regex r1 = new Regex(@"(\S*(@mail\.ru|@gmail\.com|@yandex\.ru))$");
                            if (r1.IsMatch(textBox7.Text))
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
                                    if (textBox5.Text.Length > 5 && textBox5.Text.Any(char.IsLetter))

                                    {
                                        k = 0;
                                        for (int i = 0; i < dataGridView1.RowCount; i++)
                                        {
                                            if (textBox5.Text == dataGridView1[7, i].Value.ToString())
                                            {
                                                k++;
                                            }
                                        }
                                        if (k == 0)
                                        {
                                            if (textBox6.Text.Any(char.IsLower) && textBox6.Text.Any(char.IsUpper) && textBox6.Text.Length > 8)
                                        {

                                            sqlConnection.Open();
                                            SqlCommand command = new SqlCommand($@"INSERT INTO [insurer](firstname,name,lastname,datereception,login,password,email) VALUES (@f,@n,@l,@r,@lo,@p,@e);", sqlConnection);
                                            command.Parameters.AddWithValue("@f", (textBox2.Text));
                                            command.Parameters.AddWithValue("@n", (textBox3.Text));
                                            command.Parameters.AddWithValue("@l", (textBox4.Text));
                                            command.Parameters.AddWithValue("@r", (dateTimePicker1.Value));
                                            command.Parameters.AddWithValue("lo", (textBox5.Text));
                                            command.Parameters.AddWithValue("@p", (textBox6.Text));
                                            command.Parameters.AddWithValue("@e", (textBox7.Text));
                                            command.ExecuteNonQuery();
                                            sqlConnection.Close();
                                            Insurer_load();
                                            clear();
                                            panel1.Visible = false;
                                            panel2.Visible = true;
                                            panel5.Visible = true;
                                            string query2 = $@" Select max(idinsurer) from insurer";
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
                                            Width = 1551;
                                            Height = 866;
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
                    if (textBox2.Text != "" && textBox3.Text != "" && textBox4.Text != "" && textBox5.Text != "" && textBox6.Text != "" && textBox7.Text != "")
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
                            k = 0; j = 0;
                            Regex r1 = new Regex(@"(\S*(@mail\.ru|@gmail\.com|@yandex\.ru))$");
                            if (r1.IsMatch(textBox7.Text) && textBox7.Text.Length > 8)
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
                                    j = 0; k = 0;
                                    if (textBox5.Text.Length > 5 && textBox5.Text.Any(char.IsLetter))

                                    {
                                        k = 0;j = 0;
                                        for (int i = 0; i < dataGridView1.RowCount; i++)
                                        {
                                            if (textBox5.Text == dataGridView1[7, i].Value.ToString())
                                            {
                                                k++;
                                                j = Convert.ToInt32(dataGridView1[0, i].Value.ToString());
                                            }
                                        }
                                        if (k == 0||j==0)
                                        {
                                            if (textBox6.Text.Any(char.IsLower) && textBox6.Text.Any(char.IsUpper) && textBox6.Text.Length > 8)
                                        {

                                            sqlConnection.Open();
                                            SqlCommand command = new SqlCommand($@"UPDATE insurer SET firstname=@f,name=@n ," +
                                            $"lastname=@l,datereception=@r, datelayoffs=Null,login=@lo,password=@p,email=@e WHERE idinsurer=@id", sqlConnection);
                                            command.Parameters.AddWithValue("@f", (textBox2.Text));
                                            command.Parameters.AddWithValue("@n", (textBox3.Text));
                                            command.Parameters.AddWithValue("@l", (textBox4.Text));
                                            command.Parameters.AddWithValue("@r", (dateTimePicker1.Value));
                                            command.Parameters.AddWithValue("lo", (textBox5.Text));
                                            command.Parameters.AddWithValue("@p", (textBox6.Text));
                                            command.Parameters.AddWithValue("@e", (textBox7.Text));
                                            command.Parameters.AddWithValue("@id", (idakk));
                                            command.ExecuteNonQuery();
                                            sqlConnection.Close();
                                            dataGridView1.Enabled = true;
                                            clear();
                                            Insurer_load();
                                            panel6.BringToFront();
                                            panel1.Visible = false;
                                            panel6.Visible = true;
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
            sqlConnection.Open();
            string query = $@"DELETE FROM [sostav] WHERE [idtreaty] is null";
            SqlCommand command = new SqlCommand(query, sqlConnection);
            command.ExecuteNonQuery();
            sqlConnection.Close();
            panel6.BringToFront();
            panel6.Visible = true;
            panel1.Visible = false;
            panel5.Visible = true;
            Form.Working uc = new Form.Working(this);
            Width = 1651;
            Height = 866;
            panel6.Left = 204;
            panel6.Top = button7.Top;
            panel6.Height = button7.Height;
            uc.dataGridView1.Visible = true;
            uc.dataGridView1.RowsDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            uc.dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            uc.dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            addControll(uc);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            button9.Visible = true;
            sqlConnection.Open();
            string query = $@"DELETE FROM [sostav] WHERE [idtreaty] is null";
            SqlCommand command = new SqlCommand(query, sqlConnection);
            command.ExecuteNonQuery();
            sqlConnection.Close();
            panel6.BringToFront();
            panel6.Visible = true;
            panel1.Visible = false;
            panel5.Visible = true;
            Width = 1651;
            Height = 866;
            Form.Object uc = new Form.Object(this);
            panel6.Top = button1.Top;
            panel6.Left = 204;
            panel6.Height = button1.Height;
            addControll(uc);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            button9.Visible = false;
            sqlConnection.Open();
            string query = $@"DELETE FROM [sostav] WHERE [idtreaty] is null";
            SqlCommand command = new SqlCommand(query, sqlConnection);
            command.ExecuteNonQuery();
            sqlConnection.Close();
            panel6.BringToFront();
            panel6.Visible = true;
            panel1.Visible = false;
            panel5.Visible = true;
            Width = 1651;
            Height = 866;
            Form.Service uc = new Form.Service();
            panel6.Top = button4.Top;
            panel6.Left = 204;
            panel6.Height = button4.Height;
            panel6.BringToFront();
            addControll(uc);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            button9.Visible = false;
            sqlConnection.Open();
            string query = $@"DELETE FROM [sostav] WHERE [idtreaty] is null";
            SqlCommand command = new SqlCommand(query, sqlConnection);
            command.ExecuteNonQuery();
            sqlConnection.Close();
            panel6.BringToFront();
            panel6.Visible = true;
            panel1.Visible = false;
            panel5.Visible = true;
            Width = 1651;
            Height = 866;
            Form.Calendar uc = new Form.Calendar();
            panel6.Top = button3.Top;
            panel6.Left = 204;
            panel6.Height = button3.Height;
            addControll(uc);
        }

        private void button8_Click(object sender, EventArgs e)
        {
            button9.Visible = false;
            sqlConnection.Open();
            string query = $@"DELETE FROM [sostav] WHERE [idtreaty] is null";
            SqlCommand command = new SqlCommand(query, sqlConnection);
            command.ExecuteNonQuery();
            sqlConnection.Close();
            panel6.BringToFront();
            panel6.Visible = true;
            panel1.Visible = false;
            panel5.Visible = true;
            Width = 1651;
            Height = 866;
            Form.Work uc = new Form.Work();
            panel6.Top = button8.Top;
            panel6.Height = button8.Height;
            panel6.Left = 204;
            addControll(uc);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            button9.Visible = false;
            sqlConnection.Open();
            string query = $@"DELETE FROM [sostav] WHERE [idtreaty] is null";
            SqlCommand command = new SqlCommand(query, sqlConnection);
            command.ExecuteNonQuery();
            sqlConnection.Close();
            panel6.BringToFront();
            panel6.Visible = true;
            panel1.Visible = false;
            panel5.Visible = true; ;
            Width = 1651;
            Height = 866;
            Notification uc = new Notification(this);
            panel6.Top = button5.Top;
            panel6.Left = 204;
            panel6.Height = button5.Height;
            addControll(uc);
        }
        
        int hall = 0;
        int idp = 0;
        private void button2_Click(object sender, EventArgs e)
        {
            button9.Visible = false;
            sqlConnection.Open();
            string query = $@"DELETE FROM [sostav] WHERE [idtreaty] is null";
            SqlCommand command = new SqlCommand(query, sqlConnection);
            command.ExecuteNonQuery();
            sqlConnection.Close();
            panel6.BringToFront();
            panel6.Visible = true;
            panel5.Visible = false;
            Width = 768;
            Height = 866;
            metroSetControlBox1.Visible = false;
            panel1.Visible = true;
            panel1.Visible = true;
            panel1.BringToFront();
            idp = 0;
            string query2 = $@"select ROW_NUMBER() over (ORDER BY idinsurer) num,idinsurer
from insurer
group by idinsurer";
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
            
            dateTimePicker1.Value = Convert.ToDateTime(dataGridView1[4, idp].Value.ToString());
            textBox2.Text = dataGridView1[1, idp].Value.ToString();
            textBox3.Text = dataGridView1[2, idp].Value.ToString();
            textBox4.Text = dataGridView1[3, idp].Value.ToString();
            textBox7.Text = dataGridView1[6, idp].Value.ToString();
            textBox5.Text = dataGridView1[7, idp].Value.ToString();
            textBox6.Text = dataGridView1[8, idp].Value.ToString();
            label6.Text = "Редактировать профиль";
            button11.Text = "Редактировать";
            panel1.Left = 228;
            panel1.Top = 117;
            panel6.Top = button2.Top;
            panel6.Left = 204;
            panel6.Height = button2.Height;
        }

        private void Insurerak_FormClosed(object sender, FormClosedEventArgs e)
        {
            sqlConnection.Open();
            string query = $@"DELETE FROM [sostav] WHERE [idtreaty] is null";
            SqlCommand command = new SqlCommand(query, sqlConnection);
            command.ExecuteNonQuery();
            sqlConnection.Close();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            sqlConnection.Open();
            string query = $@"DELETE FROM [sostav] WHERE [idtreaty] is null";
            SqlCommand command = new SqlCommand(query, sqlConnection);
            command.ExecuteNonQuery();
            sqlConnection.Close();
            panel6.BringToFront();
            panel6.Visible = true;
            panel1.Visible = false;
            panel5.Visible = true;
            Width = 1651;
            Height = 866;
            Form.Back uc = new Form.Back();
            panel6.Top = button9.Top;
            panel6.Left = 1286;
            panel6.Height = button9.Height;
            panel6.BringToFront();
            addControll(uc);
        }
    }
}
