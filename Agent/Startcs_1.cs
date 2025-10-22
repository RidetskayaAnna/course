using Agent.Form;
using Agent.Properties;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net.Mail;
using System.Data.SqlClient;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ToolBar;
using static MetroFramework.Drawing.MetroPaint.BorderColor;
using System.Diagnostics.Eventing.Reader;

namespace Agent
{
    public partial class Startcs : System.Windows.Forms.Form
    {
        datebase db = new datebase();
        public SqlConnection sqlConnection = null;
        public Startcs()
        {
            InitializeComponent();
            sqlConnection = new SqlConnection(db.connection);
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
        public int idak=0;
        private void button1_Click(object sender, EventArgs e)
        {
            
            if (comboBox1.SelectedIndex == 1)
            {
                try
                {
                    string query1 = $@"Select idpolicyholder from policyholder where password='{textBox1.Text}' and login='{textBox3.Text}'";
                    DataTable data = new DataTable();
                    SqlDataAdapter command1 = new SqlDataAdapter(query1, sqlConnection);
                    command1.Fill(data);
                    DataColumn column = data.Columns[0];
                    DataRow row = data.Rows[0];
                    idak = Convert.ToInt32(row[column].ToString());
                Policyholderak policyholderak = new Policyholderak(this);
                policyholderak.Width = 1551;
                policyholderak.Height = 638;
                policyholderak.panel6.Visible = true;
                policyholderak.panel1.Visible = false;
                policyholderak.ShowDialog();
                    textBox1.Text = "";
                    textBox3.Text = "";
                }
                catch
                {
                    MessageBox.Show("Неверный логин или пароль!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else if (comboBox1.SelectedIndex == 2)          
                {

                try
                {
                    string query1 = $@"Select idinsurer from insurer where password='{textBox1.Text}' and login='{textBox3.Text}'";
                    DataTable data = new DataTable();
                    SqlDataAdapter command1 = new SqlDataAdapter(query1, sqlConnection);
                    command1.Fill(data);
                    DataColumn column = data.Columns[0];
                    DataRow row = data.Rows[0];
                    idak = Convert.ToInt32(row[column].ToString());

                    string query2 = $@"Select datelayoffs from insurer where idinsurer={idak}";
                    DataTable data2 = new DataTable();
                    SqlDataAdapter command2 = new SqlDataAdapter(query2, sqlConnection);
                    command2.Fill(data2);
                    DataColumn column2 = data2.Columns[0];
                    DataRow row2 = data2.Rows[0];
                    if (row2[column2].ToString() != null && row2[column2].ToString() != "")
                    {
                        MessageBox.Show("Нет доступа. Причина: вы уволены!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    }
                    else
                    {
                        Insurerak insurerak = new Insurerak(this);
                        insurerak.Width = 1551;
                        insurerak.Height = 866;
                        insurerak.panel1.Visible = false;
                        insurerak.ShowDialog();
                        textBox1.Text = "";
                        textBox3.Text = "";
                    }

                }
                catch
                {
                    MessageBox.Show("Неверный логин или пароль!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
             
            }
            else
            { if (textBox1.Text == "1205" && textBox3.Text == "1205")
                {
                Form1 form1 = new Form1();
                form1.ShowDialog();
                textBox1.Text = "";
                textBox3.Text = "";
                } else
                {
                    MessageBox.Show("Неверный логи или пароль!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
           
           
        }

        private void label5_Click(object sender, EventArgs e)
        {
            if (textBox1.UseSystemPasswordChar == true)
            {
                textBox1.UseSystemPasswordChar = false;
                label5.Image = Image.FromFile("D:/Diplom/proga/Agent/Agent/Resources/eye.png");
            }
            else
            {
                textBox1.UseSystemPasswordChar = true;
                label5.Image = Image.FromFile("D:/Diplom/proga/Agent/Agent/Resources/1.png");
            }
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            panel1.Visible = true;
            panel2.Visible = false;
            panel3.Visible = false;
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
           
            label9.Text = "Забыли пароль или логин";
            button4.Visible = true;
            button3.Visible = true;
            button11.Visible= false;
            label8.Visible= false;
            label7.Visible= false;
            comboBox2.Visible= false;
            textBox10.Visible= false;
            panel2.Visible = true;
            panel1.Visible = false;
            panel3.Visible = false;
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            panel3.Visible = true;
            panel1.Visible = false;
            panel2.Visible = false;
        }

        private void Startcs_Load(object sender, EventArgs e)
        {
            panel2.Visible = false;
            panel3.Visible = false;
            panel1.Visible = false;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (comboBox3.SelectedIndex==0)
            {
                Policyholderak policyholderak = new Policyholderak(this);
                policyholderak.panel1.Visible = true;
                policyholderak.panel1.Left = 12;
                policyholderak.panel1.Top = 0;
                policyholderak.Width =1164;
                policyholderak.Height = 524;
                policyholderak.panel6.Visible = false;
                policyholderak.panel11.Visible = false;
                policyholderak.panel14.Visible= false;
                policyholderak.label10.Text = "Регистрация";
                policyholderak.button11.Text ="Зарегистрироваться";
                
                policyholderak.panel6.Visible = false;
                policyholderak.metroSetControlBox2.Visible = true;
                policyholderak.ShowDialog();

            }
            else
            {
                Insurerak insurerak = new Insurerak(this);
                insurerak.panel1.Visible = true;
                insurerak.panel1.Left = 0;
                insurerak.panel1.Top = 0;
                insurerak.Width = 540;
                insurerak.Height = 404;
                insurerak.panel2.Visible = false;
                insurerak.panel6.Visible = false;
                insurerak.button11.Width= 280;
                insurerak.metroSetControlBox2.Visible = false;
                insurerak.label6.Text = "Регистрация";
                insurerak.button11.Text = "Зарегистрироваться";
                insurerak.metroSetControlBox1.Visible = true;
                insurerak.ShowDialog();
            }
        }
        string kod = "";
        int id = 0;
        string email = "";
        private void button11_Click(object sender, EventArgs e)
        {
           
            if ( comboBox2.SelectedIndex != -1) {

                if (comboBox2.SelectedIndex == 0)
                {
                     if (button11.Text == "Отправить код")
                        { textBox10.MaxLength = 6;
                            Regex r2 = new Regex(@"(\S*(@mail\.ru|@gmail\.com|@yandex\.ru))$");
                            if (r2.IsMatch(textBox10.Text) && textBox10.TextLength > 8)
                            {
                            try
                            {
                                string query1 = $@"Select idpolicyholder from policyholder where email='{textBox10.Text}'";
                                DataTable data = new DataTable();
                                SqlDataAdapter command1 = new SqlDataAdapter(query1, sqlConnection);
                                command1.Fill(data);
                                DataColumn column = data.Columns[0];
                                DataRow row = data.Rows[0];
                                id = Convert.ToInt32(row[column].ToString());
                                email = textBox10.Text;
                            }   
                         catch
                        {
                            MessageBox.Show("Нет страхователя с такой почтой!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }

                        try
                                {

                                    MailAddress fromadress = new MailAddress("mih2023@mail.ru", "Admin");

                                    MailAddress toadress = new MailAddress(textBox10.Text, comboBox2.Text);
                                    MailMessage Message = new MailMessage(fromadress, toadress);
                                    Random random = new Random();
                                    kod = Convert.ToString(random.Next(100000, 1000000));
                                if (text == "password")
                                {
                                    Message.Subject = "Забыли пароль Insurance";
                                    Message.Body = "Код для смены пароля: " + kod;
                                }
                                else
                                {
                                    Message.Subject = "Забыли логин Insurance";
                                    Message.Body = "Код для смены логина: " + kod;
                                }
                                    SmtpClient smtpClient = new SmtpClient();
                                    smtpClient.Host = "smtp.mail.ru";
                                    smtpClient.Port = 587;
                                    smtpClient.EnableSsl = true;
                                    smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
                                    smtpClient.UseDefaultCredentials = false;
                                    smtpClient.Credentials = new NetworkCredential("mih2023@mail.ru", "rR4eE7xCy0chrmZxg4mu");
                                    smtpClient.Send(Message);
                                    MessageBox.Show("Проверьте почту", "Отправлено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    textBox10.Text = "";
                                    label9.Text = "Введите код";
                                    label8.Text = "Код";
                                    label7.Visible = false;
                                    comboBox2.Visible = false;
                                    button11.Text = "Проверить";


                                }
                                catch { MessageBox.Show("Ошибка!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                            }
                            else
                            {
                                MessageBox.Show("Некорректная почта!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                        else if (button11.Text == "Проверить")
                        {
                            textBox10.MaxLength = 30;
                            if (kod == textBox10.Text)
                            {
                            if (text == "password")
                            {
                                label8.Text = "Пароль";
                                label9.Text = "Новый пароль";
                                button11.Text = "Сменить пароль";
                                textBox10.Text = "";
                            }
                            else
                            {
                                label8.Text = "Логин";
                                label9.Text = "Новый логин";
                                button11.Text = "Сменить логин";
                                textBox10.Text = "";
                            }
                        }
                        else
                        { 
                            MessageBox.Show("Код неверный!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            label8.Text = "Почта";
                            if (text == "password")
                            {
                                label9.Text = "Забыли пароль";
                            }
                            else
                            {
                                label9.Text = "Забыли логин";
                            }
                            button11.Text = "Отправить код";
                            comboBox2.SelectedIndex = -1;
                            textBox10.Text = email;
                            label7.Visible = true;
                            comboBox2.Visible = true;
                            

                        }
                    }
                        else if (button11.Text == "Сменить пароль")
                        {
                           textBox10.MaxLength = 30;
                            if (textBox10.Text.Any(char.IsLower) && textBox10.Text.Any(char.IsUpper) && textBox10.Text.Length > 8)
                            {
                                sqlConnection.Open();
                                SqlCommand command = new SqlCommand($@"UPDATE policyholder SET password=@pa WHERE idpolicyholder=@id", sqlConnection);
                                command.Parameters.AddWithValue("@pa", (textBox10.Text));
                                command.Parameters.AddWithValue("@id", (id));
                                command.ExecuteNonQuery();
                                sqlConnection.Close();
                            MessageBox.Show("Пароль изменен!", "Пароль", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            label8.Text = "Почта";
                                label9.Text = "Забыли пароль";
                                button11.Text = "Отправить код";
                                comboBox2.SelectedIndex = -1;
                                textBox10.Text = "";
                            label8.Visible = false;
                            textBox10.Visible = false;
                            button3.Visible = true;
                            button4.Visible = true;
                            panel2.Visible = false;
                                panel1.Visible = true;
                            }
                            else
                            {
                                MessageBox.Show("Пароль должен быть больше 8 символов, содержать одну букву верхнего и нижнего регистра!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }else if (button11.Text == "Сменить логин")
                    {
                        textBox10.MaxLength = 30;
                        if (textBox10.Text.Length > 5 && textBox10.Text.Any(char.IsLetter))
                        {
                            sqlConnection.Open();
                            SqlCommand command = new SqlCommand($@"UPDATE policyholder SET login=@pa WHERE idpolicyholder=@id", sqlConnection);
                            command.Parameters.AddWithValue("@pa", (textBox10.Text));
                            command.Parameters.AddWithValue("@id", (id));
                            command.ExecuteNonQuery();
                            sqlConnection.Close();
                            MessageBox.Show("Логин изменен!", "Логин", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            label8.Text = "Почта";
                            label9.Text = "Забыли пароль или логин";
                            button11.Text = "Отправить код";
                            comboBox2.SelectedIndex = -1;
                            textBox10.Text = "";
                            label8.Visible = false;
                            textBox10.Visible = false;
                            button3.Visible = true;
                            button4.Visible=true;
                            panel2.Visible = false;
                            panel1.Visible = true;
                        }
                        else
                        {
                            MessageBox.Show("Логин должен быть больше 5 символов и содержать хотя бы одну букву!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }



                } else
                {
                    if (button11.Text == "Отправить код")
                    {
                        textBox10.MaxLength = 6;
                        Regex r2 = new Regex(@"(\S*(@mail\.ru|@gmail\.com|@yandex\.ru))$");
                        if (r2.IsMatch(textBox10.Text) && textBox10.TextLength > 8)
                        {
                            try
                            {
                                string query1 = $@"Select idinsurer from insurer where email='{textBox10.Text}'";
                                DataTable data = new DataTable();
                                SqlDataAdapter command1 = new SqlDataAdapter(query1, sqlConnection);
                                command1.Fill(data);
                                DataColumn column = data.Columns[0];
                                DataRow row = data.Rows[0];
                                id = Convert.ToInt32(row[column].ToString());
                                email = textBox10.Text;

                                string query2 = $@"Select datelayoffs from insurer where idinsurer={id}";
                                DataTable data2 = new DataTable();
                                SqlDataAdapter command2 = new SqlDataAdapter(query2, sqlConnection);
                                command2.Fill(data2);
                                DataColumn column2 = data2.Columns[0];
                                DataRow row2 = data2.Rows[0];
                                if (row2[column2].ToString() != null && row2[column2].ToString() != "")
                                {
                                    MessageBox.Show("Нет доступа. Причина: вы уволены!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

                                }
                                else
                                {
                                    try
                                    {

                                        MailAddress fromadress = new MailAddress("mih2023@mail.ru", "Admin");

                                        MailAddress toadress = new MailAddress(textBox10.Text, comboBox2.Text);
                                        MailMessage Message = new MailMessage(fromadress, toadress);
                                        Random random = new Random();
                                        kod = Convert.ToString(random.Next(100000, 1000000));
                                        if (text == "password")
                                        {
                                            Message.Subject = "Забыли пароль Insurance";
                                            Message.Body = "Код для смены пароля: " + kod;
                                        }
                                        else
                                        {
                                            Message.Subject = "Забыли логин Insurance";
                                            Message.Body = "Код для смены логина: " + kod;
                                        }
                                        SmtpClient smtpClient = new SmtpClient();
                                        smtpClient.Host = "smtp.mail.ru";
                                        smtpClient.Port = 587;
                                        smtpClient.EnableSsl = true;
                                        smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
                                        smtpClient.UseDefaultCredentials = false;
                                        smtpClient.Credentials = new NetworkCredential("mih2023@mail.ru", "rR4eE7xCy0chrmZxg4mu");
                                        smtpClient.Send(Message);
                                        MessageBox.Show("Проверьте почту", "Отправлено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                        textBox10.Text = "";
                                        label9.Text = "Введите код";
                                        label8.Text = "Код";
                                        label7.Visible = false;
                                        comboBox2.Visible = false;
                                        button11.Text = "Проверить";


                                    }
                                    catch { MessageBox.Show("Ошибка!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                                }

                            }
                            catch
                            {
                                MessageBox.Show("Нет страховщика с такой почтой!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                        else
                        {
                            MessageBox.Show("Некорректная почта!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else if (button11.Text == "Проверить")
                    {
                        textBox10.MaxLength = 30;
                        if (kod == textBox10.Text)
                        {
                            if (text == "password")
                            {
                                label8.Text = "Пароль";
                                label9.Text = "Новый пароль";
                                button11.Text = "Сменить пароль";
                                textBox10.Text = "";
                            }
                            else
                            {
                                label8.Text = "Логин";
                                label9.Text = "Новый логин";
                                button11.Text = "Сменить логин";
                                textBox10.Text = "";
                            }
                        }
                        else
                        {
                            MessageBox.Show("Код неверный!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            label8.Text = "Почта";
                            if (text == "password")
                            {
                                label9.Text = "Забыли пароль";
                            }
                            else
                            {
                                label9.Text = "Забыли логин";
                            }
                            button11.Text = "Отправить код";
                            comboBox2.SelectedIndex = -1;
                            textBox10.Text = email;
                            label7.Visible = true;
                            comboBox2.Visible = true;


                        }
                    }
                    else if (button11.Text == "Сменить пароль")
                    {
                        textBox10.MaxLength = 30;
                        if (textBox10.Text.Any(char.IsLower) && textBox10.Text.Any(char.IsUpper) && textBox10.Text.Length > 8)
                        {
                            sqlConnection.Open();
                            SqlCommand command = new SqlCommand($@"UPDATE insurer SET password=@pa WHERE idinsurer=@id", sqlConnection);
                            command.Parameters.AddWithValue("@pa", (textBox10.Text));
                            command.Parameters.AddWithValue("@id", (id));
                            command.ExecuteNonQuery();
                            sqlConnection.Close();
                            MessageBox.Show("Пароль изменен!", "Пароль", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            label8.Text = "Почта";
                            label9.Text = "Забыли пароль";
                            button11.Text = "Отправить код";
                            comboBox2.SelectedIndex = -1;
                            textBox10.Text = "";
                            label8.Visible = false;
                            textBox10.Visible = false;
                            button3.Visible = true;
                            button4.Visible = true;
                            panel2.Visible = false;
                            panel1.Visible = true;
                        }
                        else
                        {
                            MessageBox.Show("Пароль должен быть больше 8 символов, содержать одну букву верхнего и нижнего регистра!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else if (button11.Text == "Сменить логин")
                    {
                        textBox10.MaxLength = 30;
                        if (textBox10.Text.Length > 5 && textBox10.Text.Any(char.IsLetter))
                        {
                            sqlConnection.Open();
                            SqlCommand command = new SqlCommand($@"UPDATE insurer SET login=@pa WHERE idinsurer=@id", sqlConnection);
                            command.Parameters.AddWithValue("@pa", (textBox10.Text));
                            command.Parameters.AddWithValue("@id", (id));
                            command.ExecuteNonQuery();
                            sqlConnection.Close();
                            MessageBox.Show("Логин изменен!", "Логин", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            label8.Text = "Почта";
                            label9.Text = "Забыли пароль или логин";
                            button11.Text = "Отправить код";
                            comboBox2.SelectedIndex = -1;
                            textBox10.Text = "";
                            label8.Visible = false;
                            textBox10.Visible = false;
                            button3.Visible = true;
                            button4.Visible = true;
                            panel2.Visible = false;
                            panel1.Visible = true;
                        }
                        else
                        {
                            MessageBox.Show("Логин должен быть больше 5 символов и содержать хотя бы одну букву!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }

            }
        }

        private void textBox10_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (button11.Text == "Отправить код")
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
            }else if (label9.Text == "Забыли логин")
            {
                char c = e.KeyChar;
                e.Handled = !((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || (c >= '0' && c <= '9') || (c == '@' || c == 8 || c == '.' || c == '_' || c == ','));
            }
            else if (button11.Text == "Проверить")
            {
                if ((e.KeyChar <= 47 || e.KeyChar >= 58) && e.KeyChar != 8)
                    e.Handled = true;
            }
            else
            {
                char c = e.KeyChar;
                e.Handled = !((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || (c >= '0' && c <= '9') || (c == '@' || c == 8 || c == '.' || c == '_' || c == ',' || c == '#' || c == '+' || c == '!' || c == '$' || c == ':' || c == ';' || c == '%' || c == '^' || c == '&' || c == '*' || c == ')' || c == '(' || c == '-'));
            }
            
        }
        string text = "";
        private void button4_Click(object sender, EventArgs e)
        {
            label9.Text = "Забыли логин";
            button11.Visible = true;
            label8.Visible =true;
            label7.Visible =true;
            comboBox2.Visible =true;
            textBox10.Visible =true;
            text = "login";
            button3.Visible = false;
            button4.Visible = false;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            label9.Text = "Забыли пароль";
            button11.Visible = true;
            label8.Visible = true;
            label7.Visible = true;
            comboBox2.Visible = true;
            textBox10.Visible = true;
            text = "password";
            button3.Visible = false;
            button4.Visible = false;
        }
    }
}
