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
    public partial class Bid : UserControl
    {    datebase db = new datebase();
        public SqlConnection sqlConnection = null;
        Policyholderak policyholder;
        public Bid(Policyholderak policyholder1)
        {
            InitializeComponent();
            sqlConnection = new SqlConnection(db.connection);
            policyholder = policyholder1;
        }
        Form1 form1;
        public Bid(Form1 form)
        {
            InitializeComponent();
            sqlConnection = new SqlConnection(db.connection);
            form1 = form;
        }
        int id = 0;int id2 = 0;
        public void Objectpolicyholder_load()
        {
            sqlConnection.Open();
            DataSet dataSet = new DataSet();
            SqlDataAdapter command = new SqlDataAdapter($@"Select idbid as [Номер заявки], date as [Дата заявки], bid.idvida, vid.name as [Вид страхования], status as Статус , bid.idpolicyholder, (policyholder.firdtname +' '+policyholder.name+' '+policyholder.lastname) as Страхователь, bid.note as Пожелания from bid inner join vid on bid.idvida=vid.idvida inner join policyholder on bid.idpolicyholder=policyholder.idpolicyholder", sqlConnection);
            command.Fill(dataSet);
            dataGridView1.DataSource = dataSet.Tables[0].DefaultView;
            sqlConnection.Close();
            dataGridView1.Columns[2].Visible = false;
            dataGridView1.Columns[5].Visible = false;
            dataGridView1.AllowUserToAddRows = false;
            sqlConnection.Close();
        }
        public void Objectpolicyholder_load2()
        {
            sqlConnection.Open();
            DataSet dataSet = new DataSet();
            SqlDataAdapter command = new SqlDataAdapter($@"Select idbid as [Номер заявки], date as [Дата заявки], bid.idvida, vid.name as [Вид страхования], status as Статус , bid.idpolicyholder, (policyholder.firdtname +' '+policyholder.name+' '+policyholder.lastname) as Страхователь, bid.note as Пожелания from bid inner join vid on bid.idvida=vid.idvida inner join policyholder on bid.idpolicyholder=policyholder.idpolicyholder where bid.idpolicyholder={policyholder.idakk}", sqlConnection);
            command.Fill(dataSet);
            dataGridView1.DataSource = dataSet.Tables[0].DefaultView;
            sqlConnection.Close();
            dataGridView1.Columns[2].Visible = false;
            dataGridView1.Columns[5].Visible = false;
            dataGridView1.AllowUserToAddRows = false;
            sqlConnection.Close();
        }
        public void comboBoxcity()
        {
            sqlConnection.Close();
            sqlConnection.Open();
            string query = "select idpolicyholder,(firdtname+' '+name+' '+lastname) as p from policyholder";
            SqlDataAdapter sqlDbDataAdapter = new SqlDataAdapter(query, sqlConnection);
            DataSet dataSet = new DataSet();
            sqlDbDataAdapter.Fill(dataSet);
            comboBox1.DataSource = dataSet.Tables[0];
            comboBox1.DisplayMember = "p";
            comboBox1.ValueMember = "idpolicyholder";
            comboBox1.SelectedIndex = -1;
            sqlConnection.Close();

        }
        public void comboBoxcity3()
        {
            sqlConnection.Close();
            sqlConnection.Open();
            string query = $@"select idpolicyholder,(firdtname+' '+name+' '+lastname) as p from policyholder where idpolicyholder={policyholder.idakk}";
            SqlDataAdapter sqlDbDataAdapter = new SqlDataAdapter(query, sqlConnection);
            DataSet dataSet = new DataSet();
            sqlDbDataAdapter.Fill(dataSet);
            comboBox1.DataSource = dataSet.Tables[0];
            comboBox1.DisplayMember = "p";
            comboBox1.ValueMember = "idpolicyholder";
            comboBox1.SelectedIndex = 0;
            sqlConnection.Close();

        }
        public void comboBoxcity2()
        {
            sqlConnection.Close();
            sqlConnection.Open();
            string query = "select idpolicyholder,(firdtname+' '+name+' '+lastname) as p from policyholder";
            SqlDataAdapter sqlDbDataAdapter = new SqlDataAdapter(query, sqlConnection);
            DataSet dataSet = new DataSet();
            sqlDbDataAdapter.Fill(dataSet);
            comboBox3.DataSource = dataSet.Tables[0];
            comboBox3.DisplayMember = "p";
            comboBox3.ValueMember = "idpolicyholder";
            comboBox3.SelectedIndex = -1;
            sqlConnection.Close();

        }
        public void clear()
        {
            textBox2.Text = "";
            comboBox1.SelectedIndex = -1;
        }
        private void objectpolicyholder_Load(object sender, EventArgs e)
        {if (policyholder != null)
            {
                try
                {
                    
                    button1.Visible = false;
                    button2.Visible = false;
                    button6.Visible = false;
                    button7.Visible = false;
                    comboBoxcity3();
                    Objectpolicyholder_load2();
                }
                catch { }
            }
            else
            { Objectpolicyholder_load();
            comboBoxcity2();
            comboBoxcity();
               
                button1.Visible = true;
                button2.Visible = true;
                button6.Visible = true;
                button7.Visible = true;
            }
            dataGridView2.RowsDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView1.RowsDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
           
            dateTimePicker1.MaxDate= DateTime.Today;
            dateTimePicker1.MinDate= DateTime.Today;
            panel3.Visible = false;
            panel2.Visible = false;
            
        }

        private void button4_Click(object sender, EventArgs e)
        {

            if (panel2.Visible == false)
            {
                button7.Image = new Bitmap(@"D:\College\4kurs\Praktica4Kurs\proga\Agent\Agent\Resources\pngwing.com (48).png");
                clear();
                panel3.Visible = false;
                panel2.Visible = true;
                label6.Text = "Добавить заявку";
                button11.Text = "Добавить";
                button11.Width = 174;
                button11.Left = 161;
                if (policyholder != null)
                {
                    try
                    {
                        dataGridView1.Visible = false;
                        panel2.Top = 150;
                        Objectpolicyholder_load2();
                    }
                    catch { }
                }
                else
                {
                    dataGridView1.Visible = true;
                    panel2.Top = 471;
                    Objectpolicyholder_load();
                }
            dataGridView1.Enabled = true;
            button11.Image = new Bitmap(@"D:\College\4kurs\Praktica4Kurs\proga\Agent\Agent\Resources\pngwing.com (31).png");
            }
            else
            {
                dataGridView1.Visible = true;
                dataGridView1.Enabled = true;
                clear();
                panel2.Visible = false;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
         
            try
            {
                if (panel2.Visible == false)
                {
                    button7.Image = new Bitmap(@"D:\College\4kurs\Praktica4Kurs\proga\Agent\Agent\Resources\pngwing.com (48).png");
                panel3.Visible = false;
                clear();
                if (id != 0)
                {
                        if (dataGridView1.CurrentRow.Cells[4].Value.ToString() != "Оформлен")
                        {
                            dataGridView1.Enabled = false;
                            panel2.Visible = true;
                            label6.Text = "Редактировать заявку";
                            comboBox1.SelectedValue = dataGridView1.CurrentRow.Cells[5].Value.ToString();
                            dateTimePicker1.Value = Convert.ToDateTime(dataGridView1.CurrentRow.Cells[1].Value.ToString());
                            textBox2.Text= dataGridView1.CurrentRow.Cells[7].Value.ToString();


                            for (int i = 0; i < dataGridView2.RowCount; i++)
                            {
                                dataGridView2.Rows[i].Selected = false;
                            }

                            for (int i = 0; i < dataGridView2.RowCount; i++)
                            {
                             
                                if (dataGridView1.CurrentRow.Cells[2].Value.ToString() == dataGridView2[0, i].Value.ToString())
                                {
                                    dataGridView2.Rows[i].Selected = true;
                                    dataGridView2.Rows[i].DefaultCellStyle.SelectionForeColor = Color.Black;
                                    dataGridView2.Rows[i].DefaultCellStyle.SelectionBackColor = Color.FromArgb(212, 236, 252);
                                    break;
                                }
                            }
                            button11.Image = new Bitmap(@"D:\College\4kurs\Praktica4Kurs\proga\Agent\Agent\Resources\pngwing.com (29).png");
                            button11.Text = "Редактировать";
                            button11.Width = 241;
                            button11.Left = 127;
                        }
                        else
                        {
                            MessageBox.Show("Заявка уже рассмотрена!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
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
           
            button7.Image = new Bitmap(@"D:\College\4kurs\Praktica4Kurs\proga\Agent\Agent\Resources\pngwing.com (48).png");
            clear();
            dataGridView1.Enabled = true;
            panel2.Visible = false;
            panel3.Visible = false;
            if (id != 0)
            {
                try
                {
                    if (MessageBox.Show($@"Вы уверены что хотите удалить объект страхователя {dataGridView1.CurrentRow.Cells[1].Value.ToString()}?", "Удаление", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        sqlConnection.Open();
                        string query = $@"DELETE FROM [objectpolicyholder] WHERE [idobject] = 
                             {dataGridView1.CurrentRow.Cells[0].Value.ToString()}";
                        SqlCommand command = new SqlCommand(query, sqlConnection);
                        command.ExecuteNonQuery();
                        sqlConnection.Close();
                        Objectpolicyholder_load();
                    }
                }
                catch { MessageBox.Show("Удаление невозможно, данные все еще нужны!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            }
            else { MessageBox.Show("Выберите строку для удаления!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }

        }
       
        private void button11_Click(object sender, EventArgs e)
        {
            try
            {
                if (label6.Text == "Добавить заявку")
                {
                    if (comboBox1.SelectedIndex != -1&&id2!=0)
                    {
                        
                            sqlConnection.Open();
                            SqlCommand command = new SqlCommand($@"INSERT INTO [bid](date,idvida,status,idpolicyholder,note) VALUES (@d,@i,@s,@ip,@n);", sqlConnection);
                            command.Parameters.AddWithValue("@d", (dateTimePicker1.Value));
                            command.Parameters.AddWithValue("@i", (id2));
                            command.Parameters.AddWithValue("@s", ("Ожидание"));
                        command.Parameters.AddWithValue("@n", textBox2.Text);
                        command.Parameters.AddWithValue("@ip", (comboBox1.SelectedValue));
                            command.ExecuteNonQuery();
                            sqlConnection.Close();
                        if (policyholder != null)
                        {
                            try
                            {
                                dataGridView1.Visible = true;
                                panel2.Visible = false;
                                Objectpolicyholder_load2();
                            }
                            catch { }
                        }
                        else
                        {
                            dataGridView1.Visible = true;
                            Objectpolicyholder_load();
                        }
                            clear();
                            panel2.Visible = false;
                            id2 = 0;
                       
                    }
                    else
                    {
                        MessageBox.Show("Заполните поля!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    if (comboBox1.SelectedIndex != -1 && id2 != 0)
                    {
                        sqlConnection.Open();
                        SqlCommand command = new SqlCommand($@"UPDATE bid SET date=@d,idvida=@i," +
                        $"idpolicyholder=@ip,note=@n WHERE idbid=@id", sqlConnection);
                        command.Parameters.AddWithValue("@d", (dateTimePicker1.Value));
                        command.Parameters.AddWithValue("@i", (id2));
                        command.Parameters.AddWithValue("@ip", (comboBox1.SelectedValue));
                        command.Parameters.AddWithValue("@n", textBox2.Text);
                        command.Parameters.AddWithValue("@id", (id));
                        command.ExecuteNonQuery();
                        sqlConnection.Close();
                        dataGridView1.Enabled = true;
                        clear();
                        Objectpolicyholder_load();
                        panel2.Visible = false;
                        id2 = 0;
                    }
                    else
                    {
                        MessageBox.Show("Заполните поля!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            catch { }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                id = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
            }
            catch { }
        }
        int visible = 0;
        private void button5_Click(object sender, EventArgs e)
        {
            

            
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

       
        private void button7_Click(object sender, EventArgs e)
        {
           
            clear();
            dataGridView1.Enabled = true;
            if (panel3.Visible == true)
            {
                panel3.Visible = false;
               // comboBox2.SelectedIndex = -1;
                Objectpolicyholder_load();
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
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                dataGridView1.CurrentCell = null;
                dataGridView1.Rows[i].Visible = false;

                if (dataGridView1[6, i].Value.ToString() == comboBox3.Text)
                {
                    dataGridView1.Rows[i].Visible = true;
                }
                else
                {
                    dataGridView1.Rows[i].Visible = false;
                }
            }
        }

        private void label7_Click(object sender, EventArgs e)
        {

        }
        public void comboBoxvid()
        {
            try
            {
               
                sqlConnection.Open();
                DataSet dataSet = new DataSet();
                SqlDataAdapter command = new SqlDataAdapter($@"select idvida, vid.name as Название, note as Описание from vid, policyholder,position,work
where policyholder.idwork=position.idposition and work.idwork=position.idwork 
and idpolicyholder={comboBox1.SelectedValue}
and 
((vid.name NOT Like '%гражданской%' and vid.name NOT Like '%пассажиров%'  and work.name NOT Like '%бжд%' ) 
or (work.name  Like '%бжд%' and vid.name NOT Like '%профессиональной%')
or (work.name  Like '%бжд%' and position.name  Like '%начальник%' ))", sqlConnection);
                command.Fill(dataSet);
                dataGridView2.DataSource = dataSet.Tables[0].DefaultView;
                sqlConnection.Close();
                dataGridView2.Columns[0].Visible = false;
              
                dataGridView2.AllowUserToAddRows = false;
                sqlConnection.Close();

                
            }
            catch { }

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text != "System.Data.DataRowView" && comboBox1.Text != "")
            {
               
                comboBoxvid();
            }
        }

     

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }
        
        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            id2 =Convert.ToInt32( dataGridView2.CurrentRow.Cells[0].Value.ToString());
        }

        private void button9_Click(object sender, EventArgs e)
        {
            dataGridView1.Enabled = true;
          
        }

        private void button6_Click(object sender, EventArgs e)
        {
            dataGridView1.Enabled = true;
            clear();
          
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
            ExcelApp.Rows[1].Columns[4] = "Заявки";
            ExcelApp.Rows[visible + 3].Columns[4] = "Ридецкая Анна Михайловна";
            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                ExcelApp.Cells[2, i+1] = dataGridView1.Columns[i].HeaderText;
            }
            for (int j = 0; j < dataGridView1.ColumnCount; j++)
            {
                for (int i = 0; i < visible; i++)
                {
                    if (dataGridView1.Rows[i].Cells[j].Visible == true)
                    {
                        if (j == 1)
                        {
                            ExcelApp.Cells[i + 3, j+1] = dataGridView1.Rows[i].Cells[j].Value.ToString().Substring(0,10);
                        }
                        else
                        {
                            ExcelApp.Cells[i + 3, j+1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                        }
                        }
                }
            }
            for (int i = 0; i < visible; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    Microsoft.Office.Interop.Excel.Range range = ExcelWorkSheet.Range[$"A1:I{visible + 3}"];
                    ExcelWorkSheet.Range[$"A1:I{visible + 3}"].Cells.WrapText = true;
                    range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders.Weight = Excel.XlBorderWeight.xlMedium;
                }
            }
            ExcelApp.Columns["C"].Delete();
            ExcelApp.Columns["E"].Delete();
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }
    }
}
