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
    public partial class Tenant : UserControl
    {
        datebase db = new datebase();
        public SqlConnection sqlConnection = null;
        public Tenant()
        {
            InitializeComponent();
            sqlConnection = new SqlConnection(db.connection);
        }
        int id = 0;
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
        
       
        private void Policyholder_Load(object sender, EventArgs e)
        {
            Policyholder_load();
            panel2.Visible = false;
            //Ограничение на максимальный и минимальный ввод даты 
            dateTimePicker2.MaxDate = DateTime.Today.AddYears(-18).AddDays(-1);
            dateTimePicker2.MinDate = DateTime.Today.AddYears(-74).AddDays(-1);
          
        }
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
        private void button4_Click(object sender, EventArgs e)
        {
            if (panel2.Visible == false)
            {
                
              clear();
                panel2.Visible = true;
                label6.Text = "Добавить пользователя";
                label6.Visible = true;
            button11.Text = "Добавить";
                button11.Width = 180;
                button11.Left = 376;
                button11.Visible = true;
            Policyholder_load();
            dataGridView1.Enabled = true;
                button11.Visible = true;
            button11.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (31).png");
            }
            else
            {
                clear();
                panel2.Visible = false;
                label6.Visible = false;
            }
        }
        
        private void button1_Click(object sender, EventArgs e)
        {
            
            try
            {
                if (panel2.Visible == false)
                {
                    
                    
                    clear();
                    if (id != 0)
                    {button11.Visible = true;
                        dataGridView1.Enabled = false;
                        panel2.Visible = true;
                        label6.Text = "Редактировать пользователя";
                        label6.Visible = true;
                        button11.Width = 241;
                        button11.Left = 376;
                        textBox2.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                        textBox3.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                        textBox4.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
        
                        maskedTextBox2.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
                        dateTimePicker2.Value = Convert.ToDateTime(dataGridView1.CurrentRow.Cells[4].Value.ToString());
                        
                        textBox11.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString();
                        textBox12.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString();
                        textBox10.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();
                        button11.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (29).png");
                        button11.Text = "Редактировать";
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
                    label6.Visible = false;
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
                    if (MessageBox.Show($@"Вы уверены что хотите удалить пользователя {dataGridView1.CurrentRow.Cells[1].Value.ToString() + " " + dataGridView1.CurrentRow.Cells[2].Value.ToString() + " " + dataGridView1.CurrentRow.Cells[3].Value.ToString()}?", "Удаление", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {

                        sqlConnection.Open();
                        string query = $@"DELETE FROM [tenant] WHERE [idtenant] = 
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
        String passport = "";
        private void button11_Click(object sender, EventArgs e)
        {
            //try
            //{
                if (label6.Text == "Добавить пользователя")
                {
                k = 0;
                    if (textBox2.Text != "" && textBox3.Text != "" && textBox4.Text != "" && textBox12.Text != "" && textBox10.Text != "" && textBox11.Text != "" && count == 2 &&  maskedTextBox2.Text.Length == 18)
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
                                                        panel2.Visible = false;
                                                        label6.Visible = false;
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
                if (textBox2.Text != "" && textBox3.Text != "" && textBox4.Text != "" && textBox12.Text != "" && textBox10.Text != "" && textBox11.Text != "" && count == 2 &&  maskedTextBox2.Text.Length == 18)
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
                                k = 0;j= 0;
                                for (int i = 0; i < dataGridView1.RowCount; i++)
                                {
                                    if (maskedTextBox2.Text == dataGridView1[5, i].Value.ToString().ToLower())
                                    {
                                        k++;
                                    j = Convert.ToInt32(dataGridView1[0, i].Value.ToString());
                                }
                                }
                                if (k == 0|| j==id)
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
                                        if (k == 0 || j == id)
                                        {
                                            if (textBox11.Text.Length > 5 && textBox11.Text.Any(char.IsLetter))

                                            {
                                                k = 0;j = 0;
                                                for (int i = 0; i < dataGridView1.RowCount; i++)
                                                {
                                                    if (textBox11.Text == dataGridView1[7, i].Value.ToString())
                                                    {
                                                        k++;
                                                    j = Convert.ToInt32(dataGridView1[0, i].Value.ToString());
                                                }
                                                }
                                                if (k == 0 || j == id)
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
                                                        command.Parameters.AddWithValue("@id", (id));
                                                        command.ExecuteNonQuery();
                                                        sqlConnection.Close();
                                                        dataGridView1.Enabled = true;
                                                        clear();
                                                        Policyholder_load();
                                                        panel2.Visible = false;
                                                        label6.Visible = false;
                                                    k = 0;j= 0;
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
            //}
            //catch { }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.Length == 1)
                ((TextBox)sender).Text = ((TextBox)sender).Text.ToUpper();
            ((TextBox)sender).Select(((TextBox)sender).Text.Length, 0);
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            //Ограничение на ввод только букв русского алфавита
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
           // Ограничение на ввод только цифр
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && e.KeyChar != 8 )
                e.Handled = true;
          
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            id = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
            clear();
            panel2.Visible = false;
            label6.Visible = false;
            button11.Visible = true;
        }

        private void maskedTextBox2_TextChanged(object sender, EventArgs e)
        {     
                String s = maskedTextBox2.Text;
                String[] words = s.Split(' ');
                count=words.Length;
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
            ExcelApp.Rows[1].Columns[6] = "Пользователи";
            ExcelApp.Rows[visible + 3].Columns[6] = "Ридецкая Анна Михайловна";
            for (int i = 2; i < dataGridView1.Columns.Count; i++)
            {
                ExcelApp.Cells[2, i-1] = dataGridView1.Columns[i].HeaderText;

            }
            int y = 0;
            for (int j = 1; j < dataGridView1.ColumnCount; j++)
            {y = 0;
                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    
                    if (dataGridView1.Rows[i].Cells[j].Visible == true)
                    {
                        if (j == 4)
                        {
                            ExcelApp.Cells[y + 3, j - 1] = dataGridView1.Rows[i].Cells[j].Value.ToString().Substring(0, 10);
                            y++;
                        }
                        else
                        {
                            ExcelApp.Cells[y + 3, j - 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                            y++;
                        }
                    }
                }
            }
            for (int i = 0; i < visible; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    Microsoft.Office.Interop.Excel.Range range = ExcelWorkSheet.Range[$"A1:S{visible + 3}"];
                    ExcelWorkSheet.Range[$"A1:S{visible + 3}"].Cells.WrapText = true;
                    range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders.Weight = Excel.XlBorderWeight.xlMedium;
                }
            }

            ExcelApp.Columns["E"].Delete();
            ExcelApp.Columns["S"].Delete();
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
        { 
            //Ограничение на ввод только букв английского алфавита и специальных символов
            char c = e.KeyChar;
            e.Handled = !((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || 
            (c >= '0' && c <= '9') || (c == '@' || c == 8 || c == '.' || c == '_' 
            || c == ',' || c == '#' || c == '+' || c == '!' || c == '$' || c == ':' 
            || c == ';' || c == '%' || c == '^' || c == '&' || c == '*' || c == ')' || c == '(' || c == '-'));

           
        }

        private void textBox11_KeyPress(object sender, KeyPressEventArgs e)
        {
 char c = e.KeyChar;
            e.Handled = !((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || (c >= '0' && c <= '9') || (c == '@' || c == 8 || c == '.' || c == '_' || c == ','));

        }


      
    }
}
