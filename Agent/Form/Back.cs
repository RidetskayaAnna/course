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
    public partial class Back : UserControl
    {
        datebase db = new datebase();
        public SqlConnection sqlConnection = null;
      
        Tenantak policyholder;
        public Back(Tenantak policyholder1)
        {
            InitializeComponent();
            sqlConnection = new SqlConnection(db.connection);
            policyholder = policyholder1;
        }
        Form1 form1;
        public Back(Form1 form)
        {
            InitializeComponent();
            sqlConnection = new SqlConnection(db.connection);
            form1 = form;
        }
        Workerak insurerak;
        public Back(Workerak insurerak1)
        {
            InitializeComponent();
            sqlConnection = new SqlConnection(db.connection);
            insurerak = insurerak1;
        }
        int id = 0;int id2 = 0;
        public void Back_load()
        {
            if (policyholder != null)
            {
                sqlConnection.Open();
                DataSet dataSet = new DataSet();
                SqlDataAdapter command = new SqlDataAdapter($@"Select  idback,back.idbid as [Номер заявки], (service.name+'- '+description) as Услуга, back.date as [Дата отзыва],grade as Оценка,com as [Комментарий] from back inner join bid on back.idbid=bid.idbid inner join service on service.idservice=bid.idservice where bid.idtenant={policyholder.idakk}", sqlConnection);
                command.Fill(dataSet);
                dataGridView1.DataSource = dataSet.Tables[0].DefaultView;
                sqlConnection.Close();
                dataGridView1.Columns[0].Visible = false;
                dataGridView1.AllowUserToAddRows = false;
                sqlConnection.Close();
            }
            else if (insurerak != null)
            {
                sqlConnection.Open();
                DataSet dataSet = new DataSet();
                SqlDataAdapter command = new SqlDataAdapter($@"Select  idback,back.idbid as [Номер заявки], (service.name+'- '+description) as Услуга, back.date as [Дата отзыва],grade as Оценка,com as [Комментарий] from back inner join bid on back.idbid=bid.idbid inner join service on service.idservice=bid.idservice where bid.idworker={insurerak.idakk}", sqlConnection);
                command.Fill(dataSet);
                dataGridView1.DataSource = dataSet.Tables[0].DefaultView;
                sqlConnection.Close();
                dataGridView1.Columns[0].Visible = false;
                dataGridView1.AllowUserToAddRows = false;
                sqlConnection.Close();
            }
            else
            {
                sqlConnection.Open();
                DataSet dataSet = new DataSet();
                SqlDataAdapter command = new SqlDataAdapter($@"Select  idback,back.idbid as [Номер заявки], (service.name+'- '+description) as Услуга, back.date as [Дата отзыва],grade as Оценка,com as [Комментарий] from back inner join bid on back.idbid=bid.idbid inner join service on service.idservice=bid.idservice", sqlConnection);
                command.Fill(dataSet);
                dataGridView1.DataSource = dataSet.Tables[0].DefaultView;
                sqlConnection.Close();
                dataGridView1.Columns[0].Visible = false;
                dataGridView1.AllowUserToAddRows = false;
                sqlConnection.Close();
            }
        }

       
        public void bid_load()
        {
            if (policyholder != null)
            {
                sqlConnection.Open();
                DataSet dataSet = new DataSet();
                SqlDataAdapter command = new SqlDataAdapter($@"Select bid.idbid as [Номер заявки],  (service.name+'- '+description) as Услуга, 
(type+', '+address) as Объект, poz as Комментарий
from bid inner join object on bid.idobject=object.idobject inner join tenant on bid.idtenant=tenant.idtenant 
inner join worker on bid.idworker=worker.idworker inner join service on bid.idservice=service.idservice
left join back on back.idbid=bid.idbid
where status='Выполнено' and idback is null and bid.idtenant={policyholder.idakk}", sqlConnection);
                command.Fill(dataSet);
                dataGridView2.DataSource = dataSet.Tables[0].DefaultView;
                sqlConnection.Close();
                dataGridView2.AllowUserToAddRows = false;
                sqlConnection.Close();
            }
            else
            {
                sqlConnection.Open();
                DataSet dataSet = new DataSet();
                SqlDataAdapter command = new SqlDataAdapter($@"Select bid.idbid as [Номер заявки],  (service.name+'- '+description) as Услуга, 
(type+', '+address) as Объект, poz as Комментарий
from bid inner join object on bid.idobject=object.idobject inner join tenant on bid.idtenant=tenant.idtenant 
inner join worker on bid.idworker=worker.idworker inner join service on bid.idservice=service.idservice
left join back on back.idbid=bid.idbid
where status='Выполнено' and idback is null", sqlConnection);
                command.Fill(dataSet);
                dataGridView2.DataSource = dataSet.Tables[0].DefaultView;
                sqlConnection.Close();
                dataGridView2.AllowUserToAddRows = false;
                sqlConnection.Close();
            }
        }

        public void clear()
        {
            textBox2.Text = "";
            label7.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (5).png");
            label8.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (5).png");
            label9.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (5).png");
            label10.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (5).png");
            label11.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (5).png");

            //textBox3.Text = "";
            //textBox4.Text = "";
            //comboBox1.SelectedIndex = -1;
        }
        public void vid_load()
        {
        //    sqlConnection.Open();
        //    DataSet dataSet = new DataSet();
        //    SqlDataAdapter command = new SqlDataAdapter($@"Select idvida,vid.name as [Вид страхования] from vid", sqlConnection);
        //    command.Fill(dataSet);
        //    dataGridView2.DataSource = dataSet.Tables[0].DefaultView;
        //    sqlConnection.Close();
        //    dataGridView2.Columns[0].Visible = false;
        //    dataGridView2.AllowUserToAddRows = false;
        //    sqlConnection.Close();
        }
        private void Correctionfactor_Load(object sender, EventArgs e)
        {
            dataGridView2.RowsDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView1.RowsDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            clear();
            panel2.Visible = false;

           
     
           if (insurerak!=null) { 
                button1.Visible = false;
                button2.Visible = false;
                button4.Visible = false;
                button5.Visible = false;
                insurerak.button6.Visible = true;
                insurerak.button10.Visible = true;
                insurerak.button9.Visible = true;
            }
            else
            {  
                button1.Visible = true;
                button2.Visible = true;
                button4.Visible = true;
                button5.Visible = true;
            }
            Back_load();
            bid_load();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (panel2.Visible == false)
            {
                
                if (dataGridView2.RowCount != 0)
                {
                    Back_load();
                    bid_load();
                    clear();
                    panel2.Visible = true;
                    label6.Text = "Добавить отзыв";
                    button11.Text = "Добавить";
                    button11.Width = 174;
                    button11.Left = 137;
                    
                    dataGridView1.Enabled = true;
                    button11.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (31).png");
                }
                else
                {
                    MessageBox.Show("Все заявки уже получили отзыв!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
            }
            else
            {
                clear();
                panel2.Visible = false;
            }
            }
        int znak = 0;
        int countt = 0;
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (panel2.Visible == false)
                {
                    clear();
                    if (id != 0)
                    {
                        dataGridView1.Enabled = false;
                        panel2.Visible = true;
                        label6.Text = "Редактировать отзыв";
                        textBox2.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
                        
                            sqlConnection.Open();
                            DataSet dataSet = new DataSet();

                            SqlDataAdapter command = new SqlDataAdapter($@"Select idbid as [Номер заявки],  (service.name+'- '+description) as Услуга, 
(type+', '+address) as Объект, poz as Комментарий
from bid inner join object on bid.idobject=object.idobject inner join tenant on bid.idtenant=tenant.idtenant 
inner join worker on bid.idworker=worker.idworker inner join service on bid.idservice=service.idservice
where status='Выполнено' and bid.idbid='{id2}'", sqlConnection);
                            command.Fill(dataSet);
                            dataGridView2.DataSource = dataSet.Tables[0].DefaultView;
                            sqlConnection.Close();
                            dataGridView2.AllowUserToAddRows = false;
                            sqlConnection.Close();
                        


                        if (dataGridView1.CurrentRow.Cells[4].Value.ToString() == "\u2605 \u2605 \u2605 \u2605 \u2605")
                        {
                            star = "\u2605 \u2605 \u2605 \u2605 \u2605";
                            label7.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (6).png");
                            label8.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (6).png");
                            label9.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (6).png");
                            label10.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (6).png");
                            label11.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (6).png");
                        } else if (dataGridView1.CurrentRow.Cells[4].Value.ToString() == "\u2605 \u2605 \u2605 \u2605")
                        {
                            star = "\u2605 \u2605 \u2605 \u2605";
                            label7.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (6).png");
                            label8.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (6).png");
                            label9.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (6).png");
                            label10.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (6).png");
                            label11.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (5).png");
                        }
                        else if (dataGridView1.CurrentRow.Cells[4].Value.ToString() == "\u2605 \u2605 \u2605")
                        {
                            star = "\u2605 \u2605 \u2605";
                            label7.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (6).png");
                            label8.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (6).png");
                            label9.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (6).png");
                            label10.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (5).png");
                            label11.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (5).png");
                        }
                        else if (dataGridView1.CurrentRow.Cells[4].Value.ToString() == "\u2605 \u2605")
                        {
                            star = "\u2605 \u2605";
                            label7.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (6).png");
                            label8.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (6).png");
                            label9.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (5).png");
                            label10.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (5).png");
                            label11.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (5).png");
                        }
                        else
                        {
                            star = "\u2605";
                            label7.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (6).png");
                            label8.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (5).png");
                            label9.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (5).png");
                            label10.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (5).png");
                            label11.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (5).png");
                        }


                        //for (int i = 0; i < dataGridView2.RowCount; i++)
                        //{
                        //    dataGridView2.Rows[i].Selected = false;
                        //}
                        //    for (int i = 0; i < dataGridView2.RowCount; i++)
                        //{

                        //    if (dataGridView1.CurrentRow.Cells[2].Value.ToString() == dataGridView2[0, i].Value.ToString())
                        //    {
                        //        dataGridView2.Rows[i].Selected = true;
                        //        dataGridView2.Rows[i].DefaultCellStyle.SelectionForeColor = Color.Black;
                        //        dataGridView2.Rows[i].DefaultCellStyle.SelectionBackColor = Color.FromArgb(212, 236, 252);
                        //        break;
                        //    }
                        //}

                        button11.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (29).png");
                        button11.Text = "Редактировать";
                        button11.Width = 207;
                        button11.Left = 121;
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
                    dataGridView1.Enabled = true;
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
                    if (MessageBox.Show($@"Вы уверены что хотите удалить отзыв ?", "Удаление", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {

                        sqlConnection.Open();
                        string query = $@"DELETE FROM [back] WHERE [idback] = 
                             {dataGridView1.CurrentRow.Cells[0].Value.ToString()}";
                        SqlCommand command = new SqlCommand(query, sqlConnection);
                        command.ExecuteNonQuery();
                        sqlConnection.Close();
                        Back_load();
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
            //Фрагмент кода импорта данных в "Microsoft Excel" 
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Excel.Worksheet ExcelWorkSheet;
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
            ExcelApp.Columns.NumberFormat = "General";
            ExcelWorkSheet.StandardWidth = 30;
            ExcelWorkSheet.Columns.ColumnWidth = 20;
            ExcelApp.Rows[1].Columns[2] = "Корректировоные коэффициенты";
            ExcelApp.Rows[visible + 3].Columns[2] = "Ридецкая Анна Михайловна";
            for (int i = 1; i < dataGridView1.Columns.Count; i++)
            {
                ExcelApp.Cells[2, i] = dataGridView1.Columns[i].HeaderText;
            }
            for (int j = 1; j < dataGridView1.ColumnCount; j++)
            {
                for (int i = 0; i < visible; i++)
                {
                    if (dataGridView1.Rows[i].Cells[j].Visible == true)
                    {
                        ExcelApp.Cells[i + 3, j] = dataGridView1.Rows[i].Cells[j].Value.ToString();
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
        int k = 0;
        int j = 0;
        Decimal cof = 0;
        private void button11_Click(object sender, EventArgs e)
        {
            //try
            //{

                if (label6.Text == "Добавить отзыв")
                {
                    //if (textBox2.Text != "" && textBox3.Text != "" && textBox4.Text!="")
                    //{


                            if (k == 0)
                            {

                                sqlConnection.Open();
                                SqlCommand command = new SqlCommand($@"INSERT INTO [back](idbid,date,grade,com) VALUES (@idb,@d,@g,@c);", sqlConnection);
                                command.Parameters.AddWithValue("@idb", (id2));
                                command.Parameters.AddWithValue("@d", (dateTimePicker1.Value));
                    command.Parameters.AddWithValue("@g", star);
                    command.Parameters.AddWithValue("@c", (textBox2.Text));
                    command.ExecuteNonQuery();
                                sqlConnection.Close();
                    Back_load();
                    clear();
                                panel2.Visible = false;
                                dataGridView1.Enabled = true;
                                id2 = 0;
                            }
                            else
                            {
                                MessageBox.Show("Измените примечание!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        //}
                        //else
                        //{
                        //    MessageBox.Show("Такой корректировочный коэффициент уже есть!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        //}
                //}
                //else
                //{
                //    MessageBox.Show("Заполните поля!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //}
            }
            else
                {
                    //if (textBox2.Text != "" && textBox3.Text != "" && textBox4.Text != "")
                    //{

                        k = 0;
                        for (int i = 0; i < dataGridView1.RowCount; i++)
                        {
                            //if (textBox2.Text + Convert.ToString(cof) + textBox4.Text + id2 == dataGridView1[1, i].Value.ToString() + dataGridView1[2, i].Value.ToString() + dataGridView1[3, i].Value.ToString() + dataGridView1[4, i].Value.ToString())
                            //{
                            //    k++;
                            //    j = Convert.ToInt32(dataGridView1[0, i].Value.ToString());
                            //}
                        }
                        if (k == 0 || j == id)
                        {
                            k = 0;j = 0;

                            if (k == 0||j==id)
                            {
                                sqlConnection.Open();
                            SqlCommand command = new SqlCommand($@"UPDATE back SET idbid=@idb,date=@d ," +
                            $"grade=@g, com=@c WHERE idback=@id", sqlConnection);
                        command.Parameters.AddWithValue("@idb", (id2));
                        command.Parameters.AddWithValue("@d", (dateTimePicker1.Value));
                        command.Parameters.AddWithValue("@g", star);
                        command.Parameters.AddWithValue("@c", (textBox2.Text));
                        command.Parameters.AddWithValue("@id", (id));
                            command.ExecuteNonQuery();
                            sqlConnection.Close();
                            dataGridView1.Enabled = true;
                            clear();
                        Back_load();
                        panel2.Visible = false;
                            id2 = 0;id = 0;
                            }
                            else
                            {
                                MessageBox.Show("Измените примечание!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                        else
                        {
                            MessageBox.Show("Такой корректировочный коэффициент уже есть!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
            //}
            //        else
            //{
            //    MessageBox.Show("Заполните поля!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }
        //}
        //    catch { }
    }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            id = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
            id2 = Convert.ToInt32(dataGridView1.CurrentRow.Cells[1].Value.ToString());
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            //try
            //{
            //    cof = Convert.ToDecimal(textBox3.Text) + Convert.ToDecimal(0.01) - +Convert.ToDecimal(0.01);
            //}
            //catch { }
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && e.KeyChar != 8 && e.KeyChar != '.' && e.KeyChar != ',')
                e.Handled = true;
            if (e.KeyChar == '.')
                e.KeyChar = ',';
            if (e.KeyChar == ',')
            {
                if (((sender as TextBox).Text.IndexOf(',') != -1) || (sender as TextBox).Text.Length == 0)
                {
                    e.Handled = true;
                    return;
                }
            }
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            e.Handled = !((c >= 'а' && c <= 'я') || (c >= 'А' && c <= 'Я') || (c == 'ё' || c == 8 || c == 32 || c == ',' || c == '.' || c == '0' || c == '1' || c == '2' || c == '3' || c == '4' || c == '5' || c == '6' || c == '7' || c == '8' || c == '9'));

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.Length == 1)
                ((TextBox)sender).Text = ((TextBox)sender).Text.ToUpper();
            ((TextBox)sender).Select(((TextBox)sender).Text.Length, 0);
        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            //if (comboBox1.Text!="")
            //{
            //    char c = e.KeyChar;
            //    e.Handled = !((c == 8 || c == ',' || c == '.' || c == '0' || c == '1' || c == '2' || c == '3' || c == '4' || c == '5' || c == '6' || c == '7' || c == '8' || c == '9'));
            //}
            //else
            //{
            //    char c = e.KeyChar;
            //    e.Handled = !((c >= 'а' && c <= 'я') || (c >= 'А' && c <= 'Я') || (c == 'ё' || c == 8 || c == 32 || c == ',' || c == '.' || c == '0' || c == '1' || c == '2' || c == '3' || c == '4' || c == '5' || c == '6' || c == '7' || c == '8' || c == '9'));
            //}
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            id2 = Convert.ToInt32(dataGridView2.CurrentRow.Cells[0].Value.ToString());
           
            }
        int labe7=0;
        int labe8 = 0;
        string star = "";
        private void label7_Click(object sender, EventArgs e)
        {
            if (labe7==labe8)
            {
                star = "\u2605";
                labe8 = 1;
                label7.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (6).png");
                label8.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (5).png");
                label9.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (5).png");
                label10.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (5).png");
                label11.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (5).png");
            }
            else { labe8 = 0; label7.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (5).png");star = ""; }
        }

        private void label8_Click(object sender, EventArgs e)
        {
            star = "\u2605 \u2605";
            labe8 = 0;
            label7.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (6).png");
            label8.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (6).png");
            label9.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (5).png");
            label10.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (5).png");
            label11.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (5).png");
        }

        private void label9_Click(object sender, EventArgs e)
        {
            star = "\u2605 \u2605 \u2605";
            label7.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (6).png");
            label8.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (6).png");
            label9.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (6).png");
            label10.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (5).png");
            label11.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (5).png");
        }

        private void label10_Click(object sender, EventArgs e)
        {
            star = "\u2605 \u2605 \u2605 \u2605";
            label7.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (6).png");
            label8.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (6).png");
            label9.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (6).png");
            label10.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (6).png");
            label11.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (5).png");
        }

        private void label11_Click(object sender, EventArgs e)
        {
            star = "\u2605 \u2605 \u2605 \u2605 \u2605";
            label7.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (6).png");
            label8.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (6).png");
            label9.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (6).png");
            label10.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (6).png");
            label11.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (6).png");
        }
    }
}
