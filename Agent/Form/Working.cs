using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.ComponentModel.Design.ObjectSelectorEditor;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace Agent.Form
{
    public partial class Working : UserControl
    {
        datebase db = new datebase();
        public SqlConnection sqlConnection = null;
        Workerak insurerak;
        public Working(Workerak insurerak1)
        {
            InitializeComponent();
            sqlConnection = new SqlConnection(db.connection);
            insurerak = insurerak1;
        }
        Form1 form1;
        public Working(Form1 form)
        {
            InitializeComponent();
            sqlConnection = new SqlConnection(db.connection);
            form1 = form;
        }
        Startcs startcs;
        public Working(Startcs startcs1)
        {
            InitializeComponent();
            sqlConnection = new SqlConnection(db.connection);
            startcs = startcs1;
        }
        Tenantak policyholder;
        public Working(Tenantak policyholder1)
        {
            InitializeComponent();
            sqlConnection = new SqlConnection(db.connection);
            policyholder = policyholder1;
        }
        int id2 = 0;
        int id3 = 0;

 
        public void Tread_load()
        {

            string zap = $@"Select working.idworking, working.idobject, working.name as Работа,(type+', '+address) as Объект , working.date as Дата, status as Статус, working.idworker, (firstname+' '+ worker.name+' '+lastname) as Ответственный from worker, working, object where object.idobject=working.idobject and worker.idworker=working.idworker";

            sqlConnection.Close();
            sqlConnection.Open();
            DataSet dataSet = new DataSet();
            SqlDataAdapter command = new SqlDataAdapter($@"{zap}", sqlConnection);
            command.Fill(dataSet);
            dataGridView1.DataSource = dataSet.Tables[0].DefaultView;
            sqlConnection.Close();
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[1].Visible = false;
            dataGridView1.Columns[6].Visible = false;
            dataGridView1.AllowUserToAddRows = false;
            
            sqlConnection.Close();
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            //dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            yes = true;
            pod = true;
        }

        public void Tread_load2()
        {       
            try
            {
                    string zap = $@"Select working.idworking, working.idobject, working.name as Работа,(type+', '+address) as Объект , 
working.date as Дата, status as Статус, working.idworker, (firstname+' '+ worker.name+' '+lastname) as Ответственный from worker, 
working, object where object.idobject=working.idobject and worker.idworker=working.idworker and 
working.idworker={insurerak.idakk}";

                    sqlConnection.Close();
                    sqlConnection.Open();
                    DataSet dataSet = new DataSet();
                    SqlDataAdapter command = new SqlDataAdapter($@"{zap}", sqlConnection);
                    command.Fill(dataSet);
                    dataGridView1.DataSource = dataSet.Tables[0].DefaultView;
                    sqlConnection.Close();
                    dataGridView1.Columns[0].Visible = false;
                    dataGridView1.Columns[1].Visible = false;
                    dataGridView1.Columns[6].Visible = false;
                    dataGridView1.AllowUserToAddRows = false;

                    sqlConnection.Close();
                    dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                    //dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                    yes = true;
                    pod = true;
                
            }
            catch
            {

            }
        }
        public void comboBoxinsurer()
        {
            sqlConnection.Open();
            string query = "select idworker,(firstname+' '+name+' '+lastname) as i from worker where datelayoffs is null";
            SqlDataAdapter sqlDbDataAdapter = new SqlDataAdapter(query, sqlConnection);
            DataSet dataSet = new DataSet();
            sqlDbDataAdapter.Fill(dataSet);
            comboBox1.DataSource = dataSet.Tables[0];
            comboBox1.DisplayMember = "i";
            comboBox1.ValueMember = "idworker";
            comboBox1.SelectedIndex = -1;
            sqlConnection.Close();
        }
        public void comboBoxpolicyholder()
        {
            sqlConnection.Open();
            string query = "select worker.idworker,(firstname+' '+name+' '+lastname) as p from worker";
            SqlDataAdapter sqlDbDataAdapter = new SqlDataAdapter(query, sqlConnection);
            DataSet dataSet = new DataSet();
            sqlDbDataAdapter.Fill(dataSet);
            comboBox1.DataSource = dataSet.Tables[0];
            comboBox1.DisplayMember = "p";
            comboBox1.ValueMember = "idworker";
            comboBox1.SelectedIndex = -1;
            sqlConnection.Close();
        }
        public void comboBoxinsurer2()
        {

            sqlConnection.Open();
            string query = "select idinsurer,(firstname+' '+name+' '+lastname) as i from insurer ";
            SqlDataAdapter sqlDbDataAdapter = new SqlDataAdapter(query, sqlConnection);
            DataSet dataSet = new DataSet();
            sqlDbDataAdapter.Fill(dataSet);
            comboBox7.DataSource = dataSet.Tables[0];
            comboBox7.DisplayMember = "i";
            comboBox7.ValueMember = "idinsurer";
            comboBox7.SelectedIndex = -1;
            sqlConnection.Close();
        }
        public void objecti()
        {
            sqlConnection.Close();
            sqlConnection.Open();
            DataSet dataSet = new DataSet();
            SqlDataAdapter command = new SqlDataAdapter($@"Select idobject, type as [Тип объекта],address as Адрес, square as Площадь, com as Описание from object", sqlConnection);
            command.Fill(dataSet);
            dataGridView3.DataSource = dataSet.Tables[0].DefaultView;
            sqlConnection.Close();
            dataGridView3.Columns[0].Visible = false;
            dataGridView3.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView3.AllowUserToAddRows = false;
            dataGridView3.Columns[1].Width = 120;
            dataGridView3.Columns[2].Width = 100;
            dataGridView3.Columns[3].Width = 80;
            dataGridView3.Columns[4].Width = 500;
            sqlConnection.Close();
        }

        public void clear()
        {
            comboBox1.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;
            dateTimePicker1.Value = DateTime.Today;


        }
        public Boolean pod = false;
        Boolean no = false;
        private void Treaty_Load(object sender, EventArgs e)
        {
            //try
            //{
            //    dataGridView1.RowsDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            //    dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            //    dataGridView3.RowsDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            //    dataGridView3.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            if (insurerak != null)
            {
                Tread_load2();
            }
            else 
            {
                try
                {
                    Tread_load();
                    
                }
                catch { }
            }
            //    else
            //    {
            //        Tread_load();
            //        button4.Visible = true;
            //        button3.Visible = true;
            //        button5.Visible = true;
            //        button7.Visible = true;
            //    }
            //    pod = true;
            //    comboBoxinsurer();
            //    comboBoxpolicyholder();
            //    panel3.Visible = false;
            //    comboBoxinsurer2();
            //    comboBoxpolicyholder2();
            //    string query = $@"Select Min(dateconclusion) from treaty";
            //    System.Data.DataTable data = new System.Data.DataTable();
            //    SqlDataAdapter command = new SqlDataAdapter(query, sqlConnection);
            //    command.Fill(data);
            //    DataColumn column = data.Columns[0];
            //    DataRow row = data.Rows[0];
            //    dateTimePicker8.MinDate = Convert.ToDateTime(row[column].ToString());
            //    dateTimePicker8.MaxDate = DateTime.Today;
            //    dateTimePicker9.MaxDate = DateTime.Today;
            //    textBox1_TextChanged(sender, e);
            //}
            //catch { }
            //no = true;
        }

        Boolean yes = false;
        private void button4_Click(object sender, EventArgs e)
        {

            panel4.Visible = false;

            //sqlConnection.Open();
            //string query = $@"select idinsurer,(firstname+' '+name+' '+lastname) as i from insurer where datelayoffs is null and idinsurer={insurerak.idakk}";
            //SqlDataAdapter sqlDbDataAdapter = new SqlDataAdapter(query, sqlConnection);
            //DataSet dataSet = new DataSet();
            //sqlDbDataAdapter.Fill(dataSet);
            //comboBox1.DataSource = dataSet.Tables[0];
            //comboBox1.DisplayMember = "i";
            //comboBox1.ValueMember = "idinsurer";
            //comboBox1.SelectedIndex = 0;
            //sqlConnection.Close();

            comboBox1.Enabled = true;
            comboBox2.Enabled = true;
           

            try
            {
                comboBoxpolicyholder();
            }
            catch { }
            dataGridView1.Enabled = true;
            dateTimePicker1.MaxDate = DateTime.Today.AddMonths(1); ;
            dateTimePicker1.MinDate = DateTime.Today;
            panel6.Visible = true;

            panel3.Visible = false;
            objecti();
            id2 = 0;
            id3 = 0;
            clear();
     

        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {

        }

       


        int idtre = 0;
        private void button11_Click(object sender, EventArgs e)
        {
            //try
            //{
            if (comboBox1.SelectedIndex != -1 && comboBox2.SelectedIndex != -1)
            {
                sqlConnection.Close();
                //Фрагмент кода добавления договора страхования в БД
                sqlConnection.Open();
                SqlCommand command = new SqlCommand($@"INSERT INTO [working](
                            idobject,date,status,name,idworker)
                            VALUES (@ido,@d,@s,@n,@idwo);", sqlConnection);
                command.Parameters.AddWithValue("@ido", (id2));
                command.Parameters.AddWithValue("@idwo", (comboBox1.SelectedValue));
                command.Parameters.AddWithValue("@d", (dateTimePicker1.Value));
                command.Parameters.AddWithValue("@s", ("Назначена"));
                command.Parameters.AddWithValue("@n", (comboBox2.Text));
         
                //command.Parameters.AddWithValue("@i", (id2));
                command.ExecuteNonQuery();
            
                sqlConnection.Close();
                //sqlConnection.Open();
                //SqlCommand command2 = new SqlCommand($@"UPDATE bid SET status=@s WHERE idbid=@id", sqlConnection);
                //command2.Parameters.AddWithValue("@s", ("Оформлен"));
                //command2.Parameters.AddWithValue("@id", (id2));
                //command2.ExecuteNonQuery();
                //sqlConnection.Close();
                //}
                // six();
                clear();
                Tread_load();
                panel6.Visible = false;
                button11.Visible = false;

            }
            //else
            //{
            //    MessageBox.Show("Заполните поля!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}

            else
            {
                //if (comboBox6.SelectedIndex != -1 && comboBox1.SelectedIndex != -1 && comboBox2.SelectedIndex != -1 && comboBox5.SelectedIndex != -1 && textBox13.Text != "")
                //{
                //    if (textBox8.Text != "")
                ////    {
                //sqlConnection.Open();
                //SqlCommand command = new SqlCommand($@"INSERT INTO [treaty](idinsurer,
                //            dateconclusion,term,suminsured,datestart,datefinish,vznos,idbid)
                //           VALUES (@in,@date,@t,@s,@ds,@df,@vz,@ib);", sqlConnection);
                //command.Parameters.AddWithValue("@in", (comboBox1.SelectedValue));
                //command.Parameters.AddWithValue("@date", (dateTimePicker1.Value));

                //command.Parameters.AddWithValue("@s", Convert.ToDecimal(sumid));


                //command.Parameters.AddWithValue("@ib", (id2));
                //command.ExecuteNonQuery();
                //sqlConnection.Close();
                //sqlConnection.Open();
                //SqlCommand command2 = new SqlCommand($@"UPDATE bid SET status=@s WHERE idbid=@id", sqlConnection);
                //command2.Parameters.AddWithValue("@s", ("Оформлен"));
                //command2.Parameters.AddWithValue("@id", (id2));
                //command2.ExecuteNonQuery();
                //sqlConnection.Close();
            }
            //six();

            button11.Visible = false;

            //}
            //else
            //{
            //    MessageBox.Show("Заполните поля!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
            //    }
            //}
            //catch { }
        }
        public void six()
        {
            Tread_load();
        }

        private void button11_MouseUp(object sender, MouseEventArgs e)
        {
            //sqlConnection.Close();
            //string query2 = $@"Select max(idtreaty) from treaty";
            //System.Data.DataTable data2 = new System.Data.DataTable();
            //SqlDataAdapter command3 = new SqlDataAdapter(query2, sqlConnection);
            //command3.Fill(data2);
            //DataColumn column2 = data2.Columns[0];
            //DataRow row2 = data2.Rows[0];
            //idtre = Convert.ToInt32(row2[column2].ToString());
            //if (comboBox1.SelectedIndex != -1 && comboBox1.SelectedIndex != -1 && comboBox2.SelectedIndex != -1)
            //{
            //    sqlConnection.Open();
            //    SqlCommand command2 = new SqlCommand($@"INSERT INTO [pay](idtreaty,datepay,summa,vidpay) VALUES (@idt,@date,@s,@v);", sqlConnection);
            //    command2.Parameters.AddWithValue("@idt", (idtre));
            //    command2.Parameters.AddWithValue("@date", (dateTimePicker1.Value));
            //    command2.Parameters.AddWithValue("@s", Convert.ToDecimal(textBox1.Text));
            
            //    command2.ExecuteNonQuery();
            //    sqlConnection.Close();
            //    clear();
            //    six();
            //    sqlConnection.Close();
            //    sqlConnection.Open();
            //    SqlCommand command4 = new SqlCommand($@"UPDATE sostav SET idtreaty=@s WHERE idtreaty is null", sqlConnection);
            //    command4.Parameters.AddWithValue("@s", (idtre));
            //    command4.ExecuteNonQuery();
            //    sqlConnection.Close();

            //}
        }

        public void textBox1_TextChanged(object sender, EventArgs e)
        {
            //Код на поиск данных из "dataGridView"
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                dataGridView1.Rows[i].Selected = false;
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    if (dataGridView1.Rows[i].Cells[j].Value != null)
                    {
                        if (dataGridView1.Rows[i].Cells[j].Value.ToString().ToLower().Contains(textBox1.Text.ToLower())
                            && textBox1.Text != "")
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
            if (textBox1.Text == "")
            {
                podsvetka();
            }

        }
        public void podsvetka()
        {
//            if (pod == true)
//            {
//                int hall = 0;
//                string hyll = "";
//                DateTime min = DateTime.MinValue;
//                DateTime max = DateTime.MaxValue;
//                for (int i = 0; i < dataGridView1.RowCount; i++)
//                {
//                    try
//                    {
//                        string query2 = $@"select count(pay.idtreaty),treaty.idtreaty,dateconclusion,
//(CASE WHEN(treaty.term='В два срока')
//THEN  DATEADD(MONTH,6, datestart)
//else (CASE WHEN(treaty.term='Ежеквартально')
//THEN  DATEADD(MONTH,3, datestart)
//else null end) end) as [Дата второго взноса],
//(CASE WHEN(treaty.term='Ежеквартально')
//THEN  DATEADD(MONTH,6, datestart)
//else null end
//)as [Дата третьего взноса],
//(CASE WHEN(treaty.term='Ежеквартально')
//THEN   DATEADD(MONTH,9, datestart)
//else null end
//)as [Дата 4 взноса],
//(CASE WHEN(treaty.term='В два срока')
//THEN 2-count(pay.idtreaty) else 4-count(pay.idtreaty) end) as t,
//treaty.datestart
//from treaty inner join pay  on treaty.idtreaty=pay.idtreaty
//where treaty.term!='Единовременно' 
//and treaty.idtreaty={dataGridView1[0, i].Value.ToString()}
//group by treaty.idtreaty,treaty.term,treaty.datestart,datefinish,dateconclusion
//having (treaty.term='В два срока' and 2-count(pay.idtreaty)!=0 and treaty.term!='Ежеквартально') 
//or (4-count(pay.idtreaty)!=0 and treaty.term='Ежеквартально') ";

//                        System.Data.DataTable data2 = new System.Data.DataTable();
//                        SqlDataAdapter command2 = new SqlDataAdapter(query2, sqlConnection);
//                        command2.Fill(data2);
//                        DataColumn column2 = data2.Columns[6];
//                        DataRow row2 = data2.Rows[0];
//                        hall = Convert.ToInt32(row2[column2].ToString());
//                        DataColumn column24 = data2.Columns[4];
//                        DataRow row24 = data2.Rows[0];
//                        hyll = (row24[column24].ToString());
//                        if (hyll == "")
//                        {
//                            DataColumn column3 = data2.Columns[2];
//                            DataRow row3 = data2.Rows[0];
//                            min = Convert.ToDateTime(row3[column3].ToString()); ;
//                            DataColumn column4 = data2.Columns[3];
//                            DataRow row4 = data2.Rows[0];
//                            max = Convert.ToDateTime(row4[column4].ToString());
//                        }
//                        else if (hall == 3)
//                        {
//                            DataColumn column3 = data2.Columns[2];
//                            DataRow row3 = data2.Rows[0];
//                            min = Convert.ToDateTime(row3[column3].ToString()); ;
//                            DataColumn column4 = data2.Columns[3];
//                            DataRow row4 = data2.Rows[0];
//                            max = Convert.ToDateTime(row4[column4].ToString());
//                        }
//                        else if (hall == 2)
//                        {
//                            DataColumn column3 = data2.Columns[3];
//                            DataRow row3 = data2.Rows[0];
//                            min = Convert.ToDateTime(row3[column3].ToString());
//                            DataColumn column4 = data2.Columns[4];
//                            DataRow row4 = data2.Rows[0];
//                            max = Convert.ToDateTime(row4[column4].ToString());
//                        }
//                        else if (hall == 1 && hyll != "")
//                        {
//                            DataColumn column3 = data2.Columns[4];
//                            DataRow row3 = data2.Rows[0];
//                            min = Convert.ToDateTime(row3[column3].ToString());
//                            DataColumn column4 = data2.Columns[5];
//                            DataRow row4 = data2.Rows[0];
//                            max = Convert.ToDateTime(row4[column4].ToString());
//                        }

//                        if (Convert.ToInt32(DateTime.Today.ToString().Substring(3, 2)) + 1 == Convert.ToInt32(max.ToString().Substring(3, 2)))
//                        {
//                            dataGridView1.Rows[i].Selected = true;
//                            dataGridView1.Rows[i].DefaultCellStyle.SelectionForeColor = Color.Black;
//                            dataGridView1.Rows[i].DefaultCellStyle.SelectionBackColor = Color.FromArgb(250, 170, 209);
//                        }
//                    }
//                    catch { }

//                    if (Convert.ToDateTime(dataGridView1[10, i].Value.ToString()) <= DateTime.Today.AddMonths(1) && Convert.ToDateTime(dataGridView1[10, i].Value.ToString()) >= DateTime.Today)
//                    {
//                        dataGridView1.Rows[i].Selected = true;
//                        dataGridView1.Rows[i].DefaultCellStyle.SelectionForeColor = Color.Black;
//                        dataGridView1.Rows[i].DefaultCellStyle.SelectionBackColor = Color.FromArgb(252, 141, 143);
//                    }
//                }
           //}
        }
        private void dateTimePicker8_ValueChanged(object sender, EventArgs e)
        {
            dateTimePicker9.MinDate = dateTimePicker8.Value;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            ///// 1 по 3 ////
            if (checkBox1.Checked == true && checkBox2.Checked == true && checkBox3.Checked == true)
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    dataGridView1.CurrentCell = null;
                    dataGridView1.Rows[i].Visible = false;

                    if ((Convert.ToInt32(dataGridView1[1, i].Value.ToString()) == Convert.ToInt32(comboBox7.SelectedValue)) && Convert.ToInt32(dataGridView1[3, i].Value.ToString()) == Convert.ToInt32(comboBox4.SelectedValue) && (Convert.ToDateTime(dataGridView1[7, i].Value.ToString()) >= Convert.ToDateTime(dateTimePicker8.Value.ToString())) && (Convert.ToDateTime(dataGridView1[7, i].Value.ToString()) <= Convert.ToDateTime(dateTimePicker9.Value.ToString())))
                    {
                        dataGridView1.Rows[i].Visible = true;
                    }
                    else
                    {
                        dataGridView1.Rows[i].Visible = false;
                    }
                }
            }
            /////////////////
            //// 3 по 2 ////
            //1
            if (checkBox1.Checked == true && checkBox2.Checked == true)
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    dataGridView1.CurrentCell = null;
                    dataGridView1.Rows[i].Visible = false;
                    if ((Convert.ToInt32(dataGridView1[1, i].Value.ToString()) == Convert.ToInt32(comboBox7.SelectedValue)) && Convert.ToInt32(dataGridView1[3, i].Value.ToString()) == Convert.ToInt32(comboBox4.SelectedValue))
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
           if (checkBox1.Checked == true && checkBox3.Checked == true)
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    dataGridView1.CurrentCell = null;
                    dataGridView1.Rows[i].Visible = false;

                    if ((Convert.ToInt32(dataGridView1[1, i].Value.ToString()) == Convert.ToInt32(comboBox7.SelectedValue)) && (Convert.ToDateTime(dataGridView1[7, i].Value.ToString()) >= Convert.ToDateTime(dateTimePicker8.Value.ToString())) && (Convert.ToDateTime(dataGridView1[7, i].Value.ToString()) <= Convert.ToDateTime(dateTimePicker9.Value.ToString())))
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
            if (checkBox3.Checked == true && checkBox2.Checked == true)
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    dataGridView1.CurrentCell = null;
                    dataGridView1.Rows[i].Visible = false;

                    if ((Convert.ToInt32(dataGridView1[3, i].Value.ToString()) == Convert.ToInt32(comboBox4.SelectedValue)) && (Convert.ToDateTime(dataGridView1[7, i].Value.ToString()) >= Convert.ToDateTime(dateTimePicker8.Value.ToString())) && (Convert.ToDateTime(dataGridView1[7, i].Value.ToString()) <= Convert.ToDateTime(dateTimePicker9.Value.ToString())))
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

             //// 3 по 1 ///// 

             //1
             if (checkBox1.Checked == true)
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    dataGridView1.CurrentCell = null;
                    dataGridView1.Rows[i].Visible = false;
                    if ((Convert.ToInt32(dataGridView1[1, i].Value.ToString()) == Convert.ToInt32(comboBox7.SelectedValue)))
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
           if (checkBox2.Checked == true)
            {
                //Фрагмент кода фильтрации 
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    dataGridView1.CurrentCell = null;
                    dataGridView1.Rows[i].Visible = false;

                    if ((Convert.ToInt32(dataGridView1[3, i].Value.ToString()) == Convert.ToInt32(comboBox4.SelectedValue)))
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
            if (checkBox3.Checked == true)
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    dataGridView1.CurrentCell = null;
                    dataGridView1.Rows[i].Visible = false;
                    if ((Convert.ToDateTime(dataGridView1[7, i].Value.ToString()) >= Convert.ToDateTime(dateTimePicker8.Value.ToString())) && (Convert.ToDateTime(dataGridView1[7, i].Value.ToString()) <= Convert.ToDateTime(dateTimePicker9.Value.ToString())))
                    {
                        dataGridView1.Rows[i].Visible = true;
                    }
                    else
                    {
                        dataGridView1.Rows[i].Visible = false;
                    }
                }
            }
            else if (checkBox3.Checked == false && checkBox2.Checked == false && checkBox1.Checked == false)
            {
                Tread_load();
            }
            textBox1_TextChanged(sender, e);
            /////
        }

        private void button7_Click(object sender, EventArgs e)
        {
            try
            {
                sqlConnection.Open();
                string query = $@"DELETE FROM [sostav] WHERE [idtreaty] is null";
                SqlCommand command = new SqlCommand(query, sqlConnection);
                command.ExecuteNonQuery();
                sqlConnection.Close();

            }
            catch { }
            panel4.Visible = false; ;
            four();
            if (panel3.Visible == true)
            {
                if (insurerak.button2.Text == "Профиль")
                {
                    checkBox1.Enabled = false;
                    comboBox7.Enabled = false;
                    Tread_load();
                }
                else
                {
                    checkBox1.Enabled = true;
                    comboBox7.Enabled = true;
                    Tread_load();
                }
                panel3.Visible = false;
                button7.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (48).png");
            }
            else
            {
                if (insurerak != null)
                {
                    checkBox1.Enabled = false;
                    comboBox7.Enabled = false;
                }
                else
                {
                    checkBox1.Enabled = true;
                    comboBox7.Enabled = true;
                }
                panel3.Visible = true;
                dataGridView1.Height = 407;
                dateTimePicker8.Value = DateTime.Today;
                dateTimePicker9.Value = DateTime.Today;
                comboBox7.SelectedItem = -1;
                comboBox4.SelectedItem = -1;
                checkBox1.Checked = false;
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                button7.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (47).png");
            }
        }
        private void ReplaceWordStubs(string stubToReplace, string text, Word.Document WordDoc)
        {
            var range = WordDoc.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: stubToReplace, ReplaceWith: text);
        }
        string strax = "";
        string strax2 = "";
        string pas = "";
        string vznos = "";
        public void word()
        {
            //            panel3.Visible = false;
            //            panel6.Visible = false;
            //            if (id != 0)
            //            {
            //                if (dataGridView1.CurrentRow.Cells[6].Value.ToString() == "Медицинское страхование")
            //                {
            //                    strax = "";
            //                    strax2 = "";
            //                    pas = "";
            //                    vznos = "";
            //                    Word.Application wordApplication = new Word.Application();
            //                    string PathToNote = @"/Diplom/proga/Agent/treaty/med/med.docx";
            //                    Word.Application application = new Word.Application();
            //                    application.Visible = false;
            //                    Word.Document document = application.Documents.Open(PathToNote);
            //                    ReplaceWordStubs("{номер}", dataGridView1.CurrentRow.Cells[0].Value.ToString(), document);
            //                    ReplaceWordStubs("{число}", dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(0, 2), document);
            //                    ReplaceWordStubs("{месяц}", dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(3, 2), document);
            //                    ReplaceWordStubs("{год}", dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(6, 4), document);
            //                    ReplaceWordStubs("{страховщик}", dataGridView1.CurrentRow.Cells[2].Value.ToString(), document);
            //                    strax = dataGridView1.CurrentRow.Cells[4].Value.ToString();
            //                    string query45 = $@"select date from treaty 
            //                    inner join bid on bid.idbid=treaty.idbid where treaty.idbid={dataGridView1.CurrentRow.Cells[24].Value.ToString()}";
            //                    System.Data.DataTable data45 = new System.Data.DataTable();
            //                    SqlDataAdapter command45 = new SqlDataAdapter(query45, sqlConnection);
            //                    command45.Fill(data45);
            //                    DataColumn column45 = data45.Columns[0];
            //                    DataRow row45 = data45.Rows[0];
            //                    ReplaceWordStubs("{датазаявки}", row45[column45].ToString().Substring(0, 10), document);
            //                    string query1 = $@"Select dateb as [Дата рождения],city.indexcity,
            //                    city.name as [Город прописки],address as [Адрес], passport as [Паспорт],
            //                    numar as [Идентификационный номер],organ as [Орган, выдавший паспорт],
            //                    datep as [Дата выдачи паспорта],phone as Телефон,work.name as [Место работы],
            //                    position.name as [Должность], heal as [Группа здоровья],sport as [Занятие спортом],
            //                    email as Почта
            //                    from policyholder inner join city on policyholder.idcity=city.idcity 
            //                    inner join position on position.idposition=policyholder.idwork 
            //inner join work on work.idwork=position.idwork
            // where idpolicyholder={dataGridView1.CurrentRow.Cells[3].Value.ToString()}";
            //                    System.Data.DataTable data = new System.Data.DataTable();
            //                    SqlDataAdapter command1 = new SqlDataAdapter(query1, sqlConnection);
            //                    command1.Fill(data);
            //                    for (int i = 0; i < 10; i++)
            //                    {
            //                        DataColumn column = data.Columns[i];
            //                        DataRow row = data.Rows[0];
            //                        if (i == 1)
            //                        {
            //                            strax = strax + "; " + row[column].ToString();

            //                            pas = pas + row[column].ToString();
            //                        }
            //                        else
            //                        if (i == 2)
            //                        {
            //                            strax = strax + ", г." + row[column].ToString();
            //                            pas = pas + ", г." + row[column].ToString();
            //                        }
            //                        else
            //                        if (i == 3)
            //                        {
            //                            strax = strax + ", " + row[column].ToString();
            //                            pas = pas + ", " + row[column].ToString();
            //                        }
            //                        else
            //                        if (i == 4)
            //                        {
            //                            strax = strax + ";";
            //                            strax2 = strax2 + " паспорт " + row[column].ToString();
            //                        }
            //                        else
            //                        if (i == 5)
            //                        {
            //                            strax2 = strax2 + ", идентификационный номер " + row[column].ToString();
            //                        }
            //                        else
            //                        if (i == 6)
            //                        {
            //                            strax2 = strax2 + ", выдан  " + row[column].ToString();
            //                        }
            //                        else
            //                        if (i == 7)
            //                        {
            //                            strax2 = strax2 + "; " + row[column].ToString().Substring(0, 10);
            //                        }
            //                        else if (i == 8)
            //                        {
            //                            pas = pas + ", " + row[column].ToString();
            //                        }
            //                    }

            //                    ReplaceWordStubs("{страхователь}", strax, document);
            //                    ReplaceWordStubs("{паспорт}", strax2, document);
            //                    ReplaceWordStubs("{фио}", dataGridView1.CurrentRow.Cells[4].Value.ToString(), document);
            //                    DataColumn column5 = data.Columns[0];
            //                    DataRow row5 = data.Rows[0];
            //                    ReplaceWordStubs("{датаро}", row5[column5].ToString().Substring(0, 10), document);
            //                    ReplaceWordStubs("{паспорт}", strax2 + pas, document);
            //                    ReplaceWordStubs("{страхс}", dataGridView1.CurrentRow.Cells[11].Value.ToString(), document);
            //                    ReplaceWordStubs("{страхп}", dataGridView1.CurrentRow.Cells[15].Value.ToString(), document);

            //                    DataColumn column10 = data.Columns[11];
            //                    DataRow row10 = data.Rows[0];
            //                    if (row10[column10].ToString() == "Д-1")
            //                    {
            //                        ReplaceWordStubs("{х}", "X", document);
            //                        ReplaceWordStubs("{х}", "", document);
            //                        ReplaceWordStubs("{х}", "", document);
            //                        ReplaceWordStubs("{х}", "", document);
            //                    }
            //                    else if (row10[column10].ToString() == "Д-2")
            //                    {
            //                        ReplaceWordStubs("{х}", "", document);
            //                        ReplaceWordStubs("{х}", "X", document);
            //                        ReplaceWordStubs("{х}", "", document);
            //                        ReplaceWordStubs("{х}", "", document);
            //                    }
            //                    else if (row10[column10].ToString() == "Д-3")
            //                    {
            //                        ReplaceWordStubs("{х}", "", document);
            //                        ReplaceWordStubs("{х}", "", document);
            //                        ReplaceWordStubs("{х}", "X", document);
            //                        ReplaceWordStubs("{х}", "", document);
            //                    }
            //                    else
            //                    {
            //                        ReplaceWordStubs("{х}", "", document);
            //                        ReplaceWordStubs("{х}", "", document);
            //                        ReplaceWordStubs("{х}", "", document);
            //                        ReplaceWordStubs("{х}", "X", document);
            //                    }

            //                    ReplaceWordStubs("{страхс}", dataGridView1.CurrentRow.Cells[11].Value.ToString(), document);
            //                    ReplaceWordStubs("{датан}", dataGridView1.CurrentRow.Cells[9].Value.ToString().Substring(0, 10), document);
            //                    ReplaceWordStubs("{датак}", dataGridView1.CurrentRow.Cells[10].Value.ToString().Substring(0, 10), document);
            //                    ReplaceWordStubs("{коэф}", dataGridView1.CurrentRow.Cells[12].Value.ToString(), document);
            //                    ReplaceWordStubs("{страхп}", dataGridView1.CurrentRow.Cells[15].Value.ToString(), document);

            //                    string query2 = $@"Select  pay.vidpay,datepay,summa
            //from pay inner join  treaty on treaty.idtreaty=pay.idtreaty
            // where treaty.idtreaty={dataGridView1.CurrentRow.Cells[0].Value.ToString()}";
            //                    System.Data.DataTable data2 = new System.Data.DataTable();
            //                    SqlDataAdapter command2 = new SqlDataAdapter(query2, sqlConnection);
            //                    command2.Fill(data2);
            //                    DataColumn column1 = data2.Columns[0];
            //                    DataRow row1 = data2.Rows[0];
            //                    ReplaceWordStubs("{оплата}", row1[column1].ToString(), document);

            //                    if (dataGridView1.CurrentRow.Cells[8].Value.ToString() == "Единовременно")
            //                    {
            //                        ReplaceWordStubs("{х}", "X", document);
            //                        ReplaceWordStubs("{х}", "", document);
            //                        ReplaceWordStubs("{х}", "", document);
            //                    }
            //                    else if (dataGridView1.CurrentRow.Cells[8].Value.ToString() == "В два срока")
            //                    {
            //                        ReplaceWordStubs("{х}", "", document);
            //                        ReplaceWordStubs("{х}", "X", document);
            //                        ReplaceWordStubs("{х}", "", document);
            //                    }
            //                    else
            //                    {
            //                        ReplaceWordStubs("{х}", "", document);
            //                        ReplaceWordStubs("{х}", "", document);
            //                        ReplaceWordStubs("{х}", "X", document);
            //                    }
            //                    Random random = new Random();
            //                    ReplaceWordStubs("{страх}", dataGridView1.CurrentRow.Cells[4].Value.ToString(), document);
            //                    ReplaceWordStubs("{взнос}", dataGridView1.CurrentRow.Cells[17].Value.ToString(), document);
            //                    ReplaceWordStubs("{дата}", dataGridView1.CurrentRow.Cells[16].Value.ToString().Substring(0, 10), document);
            //                    ReplaceWordStubs("{платеж}", row1[column1].ToString() + " №" + random.Next(0, 100), document);

            //                    if (dataGridView1.CurrentRow.Cells[18].Value.ToString() != "")
            //                    {
            //                        vznos = vznos + " " + dataGridView1.CurrentRow.Cells[18].Value.ToString().Substring(0, 10) + "; " + dataGridView1.CurrentRow.Cells[19].Value.ToString() + ";";
            //                        if (dataGridView1.CurrentRow.Cells[20].Value.ToString() != "")
            //                        {
            //                            vznos = vznos + " " + dataGridView1.CurrentRow.Cells[20].Value.ToString().Substring(0, 10) + "; " + dataGridView1.CurrentRow.Cells[21].Value.ToString() + ";";
            //                            vznos = vznos + " " + dataGridView1.CurrentRow.Cells[22].Value.ToString().Substring(0, 10) + "; " + dataGridView1.CurrentRow.Cells[23].Value.ToString() + ";";
            //                        }
            //                    }
            //                    else { vznos = ""; }
            //                    ReplaceWordStubs("{взносы}", vznos, document);
            //                    ReplaceWordStubs("{страхов}", dataGridView1.CurrentRow.Cells[2].Value.ToString(), document);
            //                    DataColumn column11 = data.Columns[8];
            //                    DataRow row11 = data.Rows[0];
            //                    ReplaceWordStubs("{тел}", row11[column11].ToString(), document);
            //                    application.Visible = true;

            //                }
            //                else if (dataGridView1.CurrentRow.Cells[6].Value.ToString() == "Страхование от несчатного случая")
            //                {
            //                    strax = "";
            //                    strax2 = "";
            //                    pas = "";
            //                    vznos = "";
            //                    //Фрагмент кода на вывод «Договора страхования» 
            //                    Word.Application wordApplication = new Word.Application();
            //                    string PathToNote = @"/Diplom/proga/Agent/treaty/nes/nes.docx";
            //                    Word.Application application = new Word.Application();
            //                    application.Visible = false;
            //                    Word.Document document = application.Documents.Open(PathToNote);
            //                    ReplaceWordStubs("{номер}", dataGridView1.CurrentRow.Cells[0].Value.ToString(), document);
            //                    ReplaceWordStubs("{страховщик}", dataGridView1.CurrentRow.Cells[2].Value.ToString(), document);
            //                    strax = dataGridView1.CurrentRow.Cells[4].Value.ToString();
            //                    string query1 = $@"Select dateb as [Дата рождения],city.indexcity,
            //                    city.name as [Город прописки],address as [Адрес], passport as [Паспорт],
            //                    numar as [Идентификационный номер],organ as [Орган, выдавший паспорт],
            //                    datep as [Дата выдачи паспорта],phone as Телефон,work.name as [Место работы],
            //                    position.name as [Должность], heal as [Группа здоровья],sport as [Занятие спортом],
            //                    email as Почта,position.harmhul
            //                    from policyholder inner join city on policyholder.idcity=city.idcity 
            //                    inner join position on position.idposition=policyholder.idwork 
            //                    inner join work on work.idwork=position.idwork
            //                    where idpolicyholder={dataGridView1.CurrentRow.Cells[3].Value.ToString()}";
            //                    System.Data.DataTable data = new System.Data.DataTable();
            //                    SqlDataAdapter command1 = new SqlDataAdapter(query1, sqlConnection);
            //                    command1.Fill(data);
            //                    for (int i = 0; i < 10; i++)
            //                    {
            //                        DataColumn column = data.Columns[i];
            //                        DataRow row = data.Rows[0];
            //                        if (i == 1)
            //                        {
            //                            strax = strax + "; " + row[column].ToString();
            //                            pas = pas + row[column].ToString();
            //                        }
            //                        else
            //                        if (i == 2)
            //                        {
            //                            strax = strax + ", г." + row[column].ToString();
            //                            pas = pas + ", г." + row[column].ToString();
            //                        }
            //                        else
            //                        if (i == 3)
            //                        {
            //                            strax = strax + ", " + row[column].ToString();
            //                            pas = pas + ", " + row[column].ToString();
            //                        }
            //                        else
            //                        if (i == 4)
            //                        {
            //                            strax = strax + ";";
            //                            strax2 = strax2 + " паспорт " + row[column].ToString();
            //                        }
            //                        else
            //                        if (i == 5)
            //                        {
            //                            strax2 = strax2 + ", идентификационный номер " + row[column].ToString();
            //                        }
            //                        else
            //                        if (i == 6)
            //                        {
            //                            strax2 = strax2 + ", выдан  " + row[column].ToString();
            //                        }
            //                        else
            //                        if (i == 7)
            //                        {
            //                            strax2 = strax2 + "; " + row[column].ToString().Substring(0, 10);
            //                        }
            //                        else if (i == 8)
            //                        {
            //                            pas = pas + ", " + row[column].ToString();
            //                        }
            //                    }
            //                    ReplaceWordStubs("{страхователь}", strax, document);
            //                    ReplaceWordStubs("{паспорт}", strax2, document);
            //                    string query45 = $@"select date from treaty 
            //inner join bid on bid.idbid=treaty.idbid where treaty.idbid={dataGridView1.CurrentRow.Cells[24].Value.ToString()}";
            //                    System.Data.DataTable data45 = new System.Data.DataTable();
            //                    SqlDataAdapter command45 = new SqlDataAdapter(query45, sqlConnection);
            //                    command45.Fill(data45);
            //                    DataColumn column45 = data45.Columns[0];
            //                    DataRow row45 = data45.Rows[0];
            //                    ReplaceWordStubs("{датез}", row45[column45].ToString().Substring(0, 10), document);
            //                    ReplaceWordStubs("{страхователь}", dataGridView1.CurrentRow.Cells[4].Value.ToString(), document);
            //                    DataColumn column5 = data.Columns[0];
            //                    DataRow row5 = data.Rows[0];
            //                    ReplaceWordStubs("{датаро}", row5[column5].ToString().Substring(0, 10), document);
            //                    DataColumn column15 = data.Columns[9];
            //                    DataRow row15 = data.Rows[0];
            //                    DataColumn column16 = data.Columns[10];
            //                    DataRow row16 = data.Rows[0];
            //                    ReplaceWordStubs("{работа}", row15[column15].ToString() + "; " + row16[column16].ToString(), document);
            //                    DataColumn column10 = data.Columns[11];
            //                    DataRow row10 = data.Rows[0];
            //                    if (row10[column10].ToString() == "Д-1")
            //                    {
            //                        ReplaceWordStubs("{й}", "", document);
            //                        ReplaceWordStubs("{ц}", "Х", document);
            //                    }
            //                    else if (row10[column10].ToString() == "Д-2")
            //                    {
            //                        ReplaceWordStubs("{й}", "", document);
            //                        ReplaceWordStubs("{ц}", "Х", document);
            //                    }
            //                    else if (row10[column10].ToString() == "Д-3")
            //                    {
            //                        ReplaceWordStubs("{й}", "", document);
            //                        ReplaceWordStubs("{ц}", "Х", document);
            //                    }
            //                    else
            //                    {
            //                        ReplaceWordStubs("{й}", "Х", document);
            //                        ReplaceWordStubs("{ц}", "", document);
            //                    }
            //                    DataColumn column11 = data.Columns[12];
            //                    DataRow row11 = data.Rows[0];
            //                    if (row11[column11].ToString() == "Да")
            //                    {
            //                        ReplaceWordStubs("{у}", "Х", document);
            //                        ReplaceWordStubs("{к}", "", document);
            //                    }
            //                    else
            //                    {
            //                        ReplaceWordStubs("{у}", "", document);
            //                        ReplaceWordStubs("{к}", "X", document);
            //                    }

            //                    DataColumn column13 = data.Columns[14];
            //                    DataRow row13 = data.Rows[0];

            //                    ReplaceWordStubs("{е}", row13[column13].ToString(), document);
            //                    List<string> numbers = new List<string>() { "Никитин Николай Михайлович", "Иванова Кристина Тимофеевна", "Зайцев Владимир Даниилович" };

            //                    Random rnd = new Random();
            //                    int randIndex = rnd.Next(numbers.Count);
            //                    string random = numbers[randIndex];
            //                    ReplaceWordStubs("{выг}", random, document);
            //                    ReplaceWordStubs("{выг}", random, document);
            //                    ReplaceWordStubs("{датан}", dataGridView1.CurrentRow.Cells[9].Value.ToString().Substring(0, 10), document);
            //                    ReplaceWordStubs("{датак}", dataGridView1.CurrentRow.Cells[10].Value.ToString().Substring(0, 10), document);
            //                    ReplaceWordStubs("{страхс}", dataGridView1.CurrentRow.Cells[11].Value.ToString(), document);
            //                    ReplaceWordStubs("{бст}", dataGridView1.CurrentRow.Cells[13].Value.ToString(), document);
            //                    ReplaceWordStubs("{кк}", dataGridView1.CurrentRow.Cells[12].Value.ToString(), document);
            //                    ReplaceWordStubs("{страхт}", dataGridView1.CurrentRow.Cells[14].Value.ToString(), document);
            //                    ReplaceWordStubs("{страхп}", dataGridView1.CurrentRow.Cells[15].Value.ToString(), document);

            //                    string query2 = $@"Select  pay.vidpay,datepay,summa
            //from pay inner join  treaty on treaty.idtreaty=pay.idtreaty
            // where treaty.idtreaty={dataGridView1.CurrentRow.Cells[0].Value.ToString()}";
            //                    System.Data.DataTable data2 = new System.Data.DataTable();
            //                    SqlDataAdapter command2 = new SqlDataAdapter(query2, sqlConnection);
            //                    command2.Fill(data2);

            //                    DataColumn column1 = data2.Columns[0];
            //                    DataRow row1 = data2.Rows[0];
            //                    if (row1[column1].ToString() == "Наличные")
            //                    {
            //                        ReplaceWordStubs("{н}", "X", document);
            //                        ReplaceWordStubs("{г}", "", document);
            //                        ReplaceWordStubs("{оплата}", dataGridView1.CurrentRow.Cells[16].Value.ToString(), document);
            //                        ReplaceWordStubs("{оплата}", "", document);
            //                    }
            //                    else
            //                    {
            //                        ReplaceWordStubs("{н}", "", document);
            //                        ReplaceWordStubs("{г}", "Х", document);
            //                        ReplaceWordStubs("{оплата}", "", document);
            //                        Random random22 = new Random();
            //                        ReplaceWordStubs("{оплата}", dataGridView1.CurrentRow.Cells[16].Value.ToString().Substring(0, 10) + " №" + random22.Next(0, 100), document);
            //                    }
            //                    if (dataGridView1.CurrentRow.Cells[8].Value.ToString() == "Единовременно")
            //                    {
            //                        ReplaceWordStubs("{ш}", "X", document);
            //                        ReplaceWordStubs("{щ}", "", document);
            //                        ReplaceWordStubs("{з}", "", document);
            //                    }
            //                    else if (dataGridView1.CurrentRow.Cells[8].Value.ToString() == "В два срока")
            //                    {
            //                        ReplaceWordStubs("{ш}", "", document);
            //                        ReplaceWordStubs("{щ}", "X", document);
            //                        ReplaceWordStubs("{з}", "", document);
            //                    }
            //                    else
            //                    {
            //                        ReplaceWordStubs("{ш}", "", document);
            //                        ReplaceWordStubs("{щ}", "", document);
            //                        ReplaceWordStubs("{з}", "X", document);
            //                    }

            //                    ReplaceWordStubs("{взнос}", dataGridView1.CurrentRow.Cells[17].Value.ToString(), document);

            //                    if (dataGridView1.CurrentRow.Cells[18].Value.ToString() != "")
            //                    {
            //                        ReplaceWordStubs("{датав}", dataGridView1.CurrentRow.Cells[18].Value.ToString().Substring(0, 10), document);
            //                        ReplaceWordStubs("{взнос}", dataGridView1.CurrentRow.Cells[17].Value.ToString(), document);
            //                        ReplaceWordStubs("{датат}", "", document);
            //                        ReplaceWordStubs("{датач}", "", document);
            //                        ReplaceWordStubs("{взнос}", "", document);
            //                        ReplaceWordStubs("{взнос}", "", document);
            //                    }
            //                    else if (dataGridView1.CurrentRow.Cells[20].Value.ToString() != "")
            //                    {
            //                        ReplaceWordStubs("{датав}", dataGridView1.CurrentRow.Cells[18].Value.ToString().Substring(0, 10), document);
            //                        ReplaceWordStubs("{взнос}", dataGridView1.CurrentRow.Cells[17].Value.ToString(), document);
            //                        ReplaceWordStubs("{датат}", dataGridView1.CurrentRow.Cells[20].Value.ToString().Substring(0, 10), document);
            //                        ReplaceWordStubs("{датач}", dataGridView1.CurrentRow.Cells[22].Value.ToString().Substring(0, 10), document);
            //                        ReplaceWordStubs("{взнос}", dataGridView1.CurrentRow.Cells[17].Value.ToString(), document);
            //                        ReplaceWordStubs("{взнос}", dataGridView1.CurrentRow.Cells[17].Value.ToString(), document);
            //                    }
            //                    else
            //                    {
            //                        ReplaceWordStubs("{датав}", "", document);
            //                        ReplaceWordStubs("{взнос}", "", document);
            //                        ReplaceWordStubs("{датат}", "", document);
            //                        ReplaceWordStubs("{датач}", "", document);
            //                        ReplaceWordStubs("{взнос}", "", document);
            //                        ReplaceWordStubs("{взнос}", "", document);
            //                    }
            //                    ReplaceWordStubs("{страховщик}", dataGridView1.CurrentRow.Cells[4].Value.ToString(), document);
            //                    ReplaceWordStubs("{время}", DateTime.Now.ToString().Substring(10, 6), document);
            //                    ReplaceWordStubs("{число}", dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(0, 2), document);
            //                    ReplaceWordStubs("{месяц}", dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(3, 2), document);
            //                    ReplaceWordStubs("{год}", dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(6, 4), document);

            //                    DataColumn column12 = data.Columns[8];
            //                    DataRow row12 = data.Rows[0];

            //                    ReplaceWordStubs("{тел}", row12[column12].ToString(), document);
            //                    application.Visible = true;
            //                }
            //                else if (dataGridView1.CurrentRow.Cells[6].Value.ToString() == "Обязательное страхование гражданской ответственности перевозчика перед пассажирами")
            //                {
            //                    strax = "";
            //                    strax2 = "";
            //                    pas = "";
            //                    vznos = "";
            //                    Word.Application wordApplication = new Word.Application();
            //                    string PathToNote = @"/Diplom/proga/Agent/treaty/pere/pere.docx";
            //                    Word.Application application = new Word.Application();
            //                    application.Visible = false;
            //                    Word.Document document = application.Documents.Open(PathToNote);
            //                    ReplaceWordStubs("{дата}", dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(0, 10), document);
            //                    ReplaceWordStubs("{страхователь}", dataGridView1.CurrentRow.Cells[4].Value.ToString(), document);
            //                    ReplaceWordStubs("{датан}", dataGridView1.CurrentRow.Cells[9].Value.ToString().Substring(0, 10), document);
            //                    ReplaceWordStubs("{датан}", dataGridView1.CurrentRow.Cells[9].Value.ToString().Substring(0, 10), document);
            //                    ReplaceWordStubs("{датак}", dataGridView1.CurrentRow.Cells[10].Value.ToString().Substring(0, 10), document);
            //                    ReplaceWordStubs("{страхс}", dataGridView1.CurrentRow.Cells[11].Value.ToString(), document);
            //                    string query2 = $@"Select  pay.vidpay,datepay,summa
            //from pay inner join  treaty on treaty.idtreaty=pay.idtreaty
            // where treaty.idtreaty={dataGridView1.CurrentRow.Cells[0].Value.ToString()}";
            //                    System.Data.DataTable data2 = new System.Data.DataTable();
            //                    SqlDataAdapter command2 = new SqlDataAdapter(query2, sqlConnection);
            //                    command2.Fill(data2);
            //                    DataColumn column1 = data2.Columns[0];
            //                    DataRow row1 = data2.Rows[0];
            //                    ReplaceWordStubs("{оплата}", row1[column1].ToString().ToLower(), document);
            //                    ReplaceWordStubs("{квзнос}", dataGridView1.CurrentRow.Cells[8].Value.ToString().ToLower(), document);
            //                    ReplaceWordStubs("{взнос}", dataGridView1.CurrentRow.Cells[17].Value.ToString(), document);
            //                    ReplaceWordStubs("{квзнос}", dataGridView1.CurrentRow.Cells[8].Value.ToString().ToLower(), document);
            //                    application.Visible = true;
            //                }
            //                else if (dataGridView1.CurrentRow.Cells[6].Value.ToString() == "Страхование от несчастных случаев пассажиров железного транспорта")
            //                {
            //                    strax = "";
            //                    strax2 = "";
            //                    pas = "";
            //                    vznos = "";
            //                    Word.Application wordApplication = new Word.Application();
            //                    string PathToNote = @"/Diplom/proga/Agent/treaty/pas/pas.docx";
            //                    Word.Application application = new Word.Application();
            //                    application.Visible = false;
            //                    Word.Document document = application.Documents.Open(PathToNote);
            //                    ReplaceWordStubs("{номер}", dataGridView1.CurrentRow.Cells[0].Value.ToString(), document);
            //                    ReplaceWordStubs("{дата}", dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(0, 10), document);
            //                    ReplaceWordStubs("{страхователь}", dataGridView1.CurrentRow.Cells[4].Value.ToString(), document);
            //                    ReplaceWordStubs("{страховщик}", dataGridView1.CurrentRow.Cells[2].Value.ToString(), document);
            //                    ReplaceWordStubs("{оплата}", dataGridView1.CurrentRow.Cells[8].Value.ToString().ToLower(), document);
            //                    ReplaceWordStubs("{взнос}", dataGridView1.CurrentRow.Cells[17].Value.ToString(), document);
            //                    ReplaceWordStubs("{число}", dataGridView1.CurrentRow.Cells[16].Value.ToString().Substring(0, 2), document);
            //                    ReplaceWordStubs("{датан}", dataGridView1.CurrentRow.Cells[9].Value.ToString().Substring(0, 10), document);
            //                    ReplaceWordStubs("{датак}", dataGridView1.CurrentRow.Cells[10].Value.ToString().Substring(0, 10), document);
            //                    application.Visible = true;
            //                }
            //                else if (dataGridView1.CurrentRow.Cells[6].Value.ToString() == "Страхование жизни")
            //                {
            //                    strax = "";
            //                    strax2 = "";
            //                    pas = "";
            //                    vznos = "";
            //                    Word.Application wordApplication = new Word.Application();

            //                    string PathToNote = @"/Diplom/proga/Agent/treaty/live/live.docx";
            //                    Word.Application application = new Word.Application();
            //                    application.Visible = false;
            //                    Word.Document document = application.Documents.Open(PathToNote);
            //                    ReplaceWordStubs("{номер}", dataGridView1.CurrentRow.Cells[0].Value.ToString(), document);
            //                    ReplaceWordStubs("{дата}", dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(0, 10), document);
            //                    ReplaceWordStubs("{страховщик}", dataGridView1.CurrentRow.Cells[2].Value.ToString(), document);

            //                    string query1 = $@"Select dateb as [Дата рождения],city.indexcity,
            //city.name as [Город прописки],address as [Адрес], passport as [Паспорт],
            //numar as [Идентификационный номер],organ as [Орган, выдавший паспорт],
            //datep as [Дата выдачи паспорта],phone as Телефон,work.name as [Место работы],
            //position.name as [Должность], heal as [Группа здоровья],sport as [Занятие спортом],
            //email as Почта,position.harmhul
            //from policyholder inner join city on policyholder.idcity=city.idcity 
            //inner join position on position.idposition=policyholder.idwork 
            //inner join work on work.idwork=position.idwork
            // where idpolicyholder={dataGridView1.CurrentRow.Cells[3].Value.ToString()}";
            //                    System.Data.DataTable data = new System.Data.DataTable();
            //                    SqlDataAdapter command1 = new SqlDataAdapter(query1, sqlConnection);
            //                    command1.Fill(data);
            //                    for (int i = 0; i < 10; i++)
            //                    {
            //                        DataColumn column = data.Columns[i];
            //                        DataRow row = data.Rows[0];
            //                        if (i == 1)
            //                        {
            //                            strax = strax + " " + row[column].ToString();

            //                            pas = pas + row[column].ToString();
            //                        }
            //                        else
            //                        if (i == 2)
            //                        {
            //                            strax = strax + ", г." + row[column].ToString();
            //                            pas = pas + ", г." + row[column].ToString();
            //                        }
            //                        else
            //                        if (i == 3)
            //                        {
            //                            strax = strax + ", " + row[column].ToString();
            //                            pas = pas + ", " + row[column].ToString();
            //                        }
            //                        else
            //                        if (i == 4)
            //                        {
            //                            strax = strax + ";";
            //                            strax2 = strax2 + " паспорт " + row[column].ToString();
            //                        }
            //                        else
            //                        if (i == 5)
            //                        {
            //                            strax2 = strax2 + ", идентификационный номер " + row[column].ToString();
            //                        }
            //                        else
            //                        if (i == 6)
            //                        {
            //                            strax2 = strax2 + ", выдан  " + row[column].ToString();
            //                        }
            //                        else
            //                        if (i == 7)
            //                        {
            //                            strax2 = strax2 + "; " + row[column].ToString().Substring(0, 10);
            //                        }
            //                        else if (i == 8)
            //                        {
            //                            pas = pas + ", " + row[column].ToString();
            //                        }
            //                    }

            //                    ReplaceWordStubs("{паспорт}", strax + strax2, document);
            //                    DataColumn column5 = data.Columns[0];
            //                    DataRow row5 = data.Rows[0];
            //                    ReplaceWordStubs("{датаро}", row5[column5].ToString().Substring(0, 10), document);
            //                    ReplaceWordStubs("{датаро}", row5[column5].ToString().Substring(0, 10), document);
            //                    ReplaceWordStubs("{паспорт}", strax + strax2, document);
            //                    List<string> numbers = new List<string>() { "Никитин Николай Михайлович", "Иванова Кристина Тимофеевна", "Зайцев Владимир Даниилович" };

            //                    Random rnd = new Random();
            //                    int randIndex = rnd.Next(numbers.Count);
            //                    string random = numbers[randIndex];

            //                    ReplaceWordStubs("{выг}", random, document);
            //                    ReplaceWordStubs("{страхс}", Convert.ToString(Math.Round((Convert.ToDecimal(dataGridView1.CurrentRow.Cells[11].Value.ToString()) / 2), 2)), document);
            //                    ReplaceWordStubs("{страхсс}", Convert.ToString(Math.Round((Convert.ToDecimal(dataGridView1.CurrentRow.Cells[11].Value.ToString()) / 3), 2)), document);
            //                    ReplaceWordStubs("{страхссс}", Convert.ToString(Math.Round((Convert.ToDecimal(dataGridView1.CurrentRow.Cells[11].Value.ToString()) - (Convert.ToDecimal(dataGridView1.CurrentRow.Cells[11].Value.ToString()) / 3) - (Convert.ToDecimal(dataGridView1.CurrentRow.Cells[11].Value.ToString()) / 2)), 2)), document);
            //                    ReplaceWordStubs("{страхп}", dataGridView1.CurrentRow.Cells[15].Value.ToString(), document);
            //                    ReplaceWordStubs("{квзнос}", dataGridView1.CurrentRow.Cells[8].Value.ToString().ToLower(), document);

            //                    if (dataGridView1.CurrentRow.Cells[20].Value.ToString() != "")
            //                    {
            //                        ReplaceWordStubs("{датав}", dataGridView1.CurrentRow.Cells[22].Value.ToString().Substring(0, 10), document);
            //                    }
            //                    else if (dataGridView1.CurrentRow.Cells[18].Value.ToString() != "")
            //                    {
            //                        ReplaceWordStubs("{датав}", dataGridView1.CurrentRow.Cells[18].Value.ToString().Substring(0, 10), document);
            //                    }
            //                    else
            //                    {
            //                        ReplaceWordStubs("{датав}", dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(0, 10), document);
            //                    }
            //                    ReplaceWordStubs("{датан}", dataGridView1.CurrentRow.Cells[9].Value.ToString().Substring(0, 10), document);
            //                    ReplaceWordStubs("{датак}", dataGridView1.CurrentRow.Cells[10].Value.ToString().Substring(0, 10), document);
            //                    ReplaceWordStubs("{страховщик}", dataGridView1.CurrentRow.Cells[2].Value.ToString(), document);
            //                    ReplaceWordStubs("{страхователь}", dataGridView1.CurrentRow.Cells[4].Value.ToString(), document);
            //                    application.Visible = true;
            //                }
            //                else if (dataGridView1.CurrentRow.Cells[6].Value.ToString() == "Страхование гражданской ответственности за причинение вреда в связи с осуществлением профессиональной деятельности")
            //                {
            //                    strax = "";
            //                    strax2 = "";
            //                    pas = "";
            //                    vznos = "";
            //                    Word.Application wordApplication = new Word.Application();

            //                    string PathToNote = @"/Diplom/proga/Agent/treaty/prof/prof.doc";
            //                    Word.Application application = new Word.Application();
            //                    application.Visible = false;
            //                    Word.Document document = application.Documents.Open(PathToNote);
            //                    ReplaceWordStubs("{номер}", dataGridView1.CurrentRow.Cells[0].Value.ToString(), document);
            //                    ReplaceWordStubs("{дата}", dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(0, 10), document);

            //                    string query1 = $@"Select 
            //position.name as [Должность]
            //from policyholder inner join city on policyholder.idcity=city.idcity 
            //inner join position on position.idposition=policyholder.idwork 
            //inner join work on work.idwork=position.idwork
            // where idpolicyholder={dataGridView1.CurrentRow.Cells[3].Value.ToString()}";
            //                    System.Data.DataTable data = new System.Data.DataTable();
            //                    SqlDataAdapter command1 = new SqlDataAdapter(query1, sqlConnection);
            //                    command1.Fill(data);
            //                    DataColumn column = data.Columns[0];
            //                    DataRow row = data.Rows[0];
            //                    ReplaceWordStubs("{должность}", row[column].ToString().ToLower(), document);
            //                    ReplaceWordStubs("{страхователь}", dataGridView1.CurrentRow.Cells[4].Value.ToString(), document);
            //                    ReplaceWordStubs("{датан}", dataGridView1.CurrentRow.Cells[9].Value.ToString().Substring(0, 10), document);
            //                    ReplaceWordStubs("{датан}", dataGridView1.CurrentRow.Cells[9].Value.ToString().Substring(0, 10), document);
            //                    ReplaceWordStubs("{датак}", dataGridView1.CurrentRow.Cells[10].Value.ToString().Substring(0, 10), document);
            //                    ReplaceWordStubs("{взнос}", dataGridView1.CurrentRow.Cells[17].Value.ToString(), document);
            //                    ReplaceWordStubs("{квзнос}", dataGridView1.CurrentRow.Cells[8].Value.ToString().ToLower(), document);

            //                    if (dataGridView1.CurrentRow.Cells[20].Value.ToString() != "")
            //                    {
            //                        ReplaceWordStubs("{датав}", dataGridView1.CurrentRow.Cells[18].Value.ToString().Substring(0, 10) + "; " + dataGridView1.CurrentRow.Cells[20].Value.ToString().Substring(0, 10) + "; " + dataGridView1.CurrentRow.Cells[22].Value.ToString().Substring(0, 10), document);

            //                    }
            //                    else if (dataGridView1.CurrentRow.Cells[18].Value.ToString() != "")
            //                    {
            //                        ReplaceWordStubs("{датав}", dataGridView1.CurrentRow.Cells[18].Value.ToString().Substring(0, 10), document);
            //                    }
            //                    else
            //                    {
            //                        ReplaceWordStubs("{датав}", dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(0, 10), document);
            //                    }
            //                    ReplaceWordStubs("{страхователь}", dataGridView1.CurrentRow.Cells[4].Value.ToString(), document);
            //                    application.Visible = true;
            //                }
            //                else if (dataGridView1.CurrentRow.Cells[6].Value.ToString() == "Страхование средств железнодорожного транспорта")
            //                {
            //                    sqlConnection.Open();
            //                    DataSet dataSet = new DataSet();
            //                    SqlDataAdapter command = new SqlDataAdapter($@"
            //Select  train.name +'; '+cast(train.nomer as  varchar(50)),train.type+'; '+cast(train.certificate as varchar(7))+'; '+
            //cast(train.year as varchar(4))+' г.',
            //price as [Страховая стоимость], sostav.summa as [Страховая сумма],
            //cast( case when pay.vidpay='Наличные' then coefficient*0.24*1 else coefficient*0.24*0.9 end  as decimal(18,2)) 
            //as [Страховой тариф],cast( case when pay.vidpay='Наличные' then sostav.summa*(coefficient*0.24)*1 else sostav.summa*(coefficient*0.24)*0.9 end as decimal(18,2))  as [Страховая премия]  
            //from train,sostav,correctionfactor,pay
            //where  train.idtrain=sostav.idtrain and  correctionfactor.name like '%средств%' and sostav.idtreaty={dataGridView1.CurrentRow.Cells[0].Value.ToString()}  and sostav.idtreaty=pay.idtreaty
            //and coefficient in (
            //select TOP(Select count(sostav.idsostav) from sostav inner join treaty on sostav.idtreaty=treaty.idtreaty where sostav.idtreaty={dataGridView1.CurrentRow.Cells[0].Value.ToString()} ) 
            //(
            //(case when(
            //(case when (SUBSTRING(note,0,3)='>=' or SUBSTRING(note,0,3)='<=')then SUBSTRING(note,0,3) 
            //else SUBSTRING(note,0,2) end)='>=') then 
            //(case when(
            //(case when (SUBSTRING(note,0,3)='>=' or SUBSTRING(note,0,3)='<=')then SUBSTRING(note,3,len(note)) 
            //else SUBSTRING(note,2,len(note)) end)<= (Select count(sostav.idsostav) from sostav inner join treaty on sostav.idtreaty=treaty.idtreaty where sostav.idtreaty={dataGridView1.CurrentRow.Cells[0].Value.ToString()} ))
            //then coefficient else 
            //0
            //end
            //)
            //else (case when(
            //(case when (SUBSTRING(note,0,3)='>=' or SUBSTRING(note,0,3)='<=')then SUBSTRING(note,0,3) 
            //else SUBSTRING(note,0,2) end)='<=') then 
            //(case when(
            //(case when (SUBSTRING(note,0,3)='>=' or SUBSTRING(note,0,3)='<=')then SUBSTRING(note,3,len(note)) 
            //else SUBSTRING(note,2,len(note)) end)>= (Select count(sostav.idsostav) from sostav inner join treaty on sostav.idtreaty=treaty.idtreaty where sostav.idtreaty={dataGridView1.CurrentRow.Cells[0].Value.ToString()} ))
            //then coefficient else 0 end
            //)
            //else 0 end) end)
            //)
            //from train,sostav,correctionfactor
            //where  train.idtrain=sostav.idtrain and  correctionfactor.name like '%средств%' and sostav.idtreaty={dataGridView1.CurrentRow.Cells[0].Value.ToString()} 
            //group by idsostav,sostav.idtreaty, sostav.idtrain, train.name ,train.nomer, 
            //price, sostav.summa ,sostav.idsostav,note,coefficient
            //order by correctionfactor.note Desc
            //)
            //group by idsostav,sostav.idtreaty, sostav.idtrain, train.name ,train.nomer, 
            //price, sostav.summa ,sostav.idsostav,note,coefficient,vidpay,type,certificate,year
            //having (
            //(case when(
            //(case when (SUBSTRING(note,0,3)='>=' or SUBSTRING(note,0,3)='<=')then SUBSTRING(note,0,3) 
            //else SUBSTRING(note,0,2) end)='>=') then 
            //(case when(
            //(case when (SUBSTRING(note,0,3)='>=' or SUBSTRING(note,0,3)='<=')then SUBSTRING(note,3,len(note)) 
            //else SUBSTRING(note,2,len(note)) end)<= (Select count(sostav.idsostav) from sostav inner join treaty on sostav.idtreaty=treaty.idtreaty where sostav.idtreaty={dataGridView1.CurrentRow.Cells[0].Value.ToString()} ))
            //then coefficient else 
            //0
            //end
            //)
            //else (case when(
            //(case when (SUBSTRING(note,0,3)='>=' or SUBSTRING(note,0,3)='<=')then SUBSTRING(note,0,3) 
            //else SUBSTRING(note,0,2) end)='<=') then 
            //(case when(
            //(case when (SUBSTRING(note,0,3)='>=' or SUBSTRING(note,0,3)='<=')then SUBSTRING(note,3,len(note)) 
            //else SUBSTRING(note,2,len(note)) end)>= (Select count(sostav.idsostav) from sostav inner join treaty on sostav.idtreaty=treaty.idtreaty where sostav.idtreaty={dataGridView1.CurrentRow.Cells[0].Value.ToString()} ))
            //then coefficient else 0 end
            //)
            //else 0 end) end)
            //) !=0

            //union
            //Select  train.name +'; '+cast(train.nomer as  varchar(50)),train.type+'; '+cast(train.certificate as varchar(7))+'; '+
            //cast(train.year as varchar(4))+' г.',
            //price as [Страховая стоимость], sostav.summa as [Страховая сумма],
            //cast( case when pay.vidpay='Наличные' then coefficient*0.24*1 else coefficient*0.24*0.9 end  as decimal(18,2)) 
            //as [Страховой тариф],cast( case when pay.vidpay='Наличные' then sostav.summa*(coefficient*0.24)*1 else sostav.summa*(coefficient*0.24)*0.9 end as decimal(18,2))  as [Страховая премия]  
            //from train,sostav,correctionfactor,pay
            //where  train.idtrain=sostav.idtrain and  correctionfactor.name like '%средств%' and sostav.idtreaty={dataGridView1.CurrentRow.Cells[0].Value.ToString()}  and sostav.idtreaty=pay.idtreaty
            //and coefficient in (
            //select TOP(Select count(sostav.idsostav) from sostav inner join treaty on sostav.idtreaty=treaty.idtreaty where sostav.idtreaty={dataGridView1.CurrentRow.Cells[0].Value.ToString()} ) 
            //(
            //(case when(
            //(case when (SUBSTRING(note,0,3)='>=' or SUBSTRING(note,0,3)='<=')then SUBSTRING(note,0,3) 
            //else SUBSTRING(note,0,2) end)='=') then 
            //(case when(
            //(case when (SUBSTRING(note,0,3)='>=' or SUBSTRING(note,0,3)='<=')then SUBSTRING(note,3,len(note)) 
            //else SUBSTRING(note,2,len(note)) end)= (Select count(sostav.idsostav) from sostav inner join treaty on sostav.idtreaty=treaty.idtreaty where sostav.idtreaty={dataGridView1.CurrentRow.Cells[0].Value.ToString()} ))
            //then coefficient else 
            //0
            //end
            //)
            //else (case when(
            //(case when (SUBSTRING(note,0,3)='>=' or SUBSTRING(note,0,3)='<=')then SUBSTRING(note,0,3) 
            //else SUBSTRING(note,0,2) end)='>') then 
            //(case when(
            //(case when (SUBSTRING(note,0,3)='>=' or SUBSTRING(note,0,3)='<=')then SUBSTRING(note,3,len(note)) 
            //else SUBSTRING(note,2,len(note)) end)< (Select count(sostav.idsostav) from sostav inner join treaty on sostav.idtreaty=treaty.idtreaty where sostav.idtreaty={dataGridView1.CurrentRow.Cells[0].Value.ToString()} ))
            //then coefficient else 0 end
            //)
            //else 0 end) end)
            //)
            //from train,sostav,correctionfactor
            //where  train.idtrain=sostav.idtrain and  correctionfactor.name like '%средств%'
            //group by idsostav,sostav.idtreaty, sostav.idtrain, train.name ,train.nomer, 
            //price, sostav.summa ,sostav.idsostav,note,coefficient
            //order by correctionfactor.note Desc
            //)
            //group by idsostav,sostav.idtreaty, sostav.idtrain, train.name ,train.nomer, 
            //price, sostav.summa ,sostav.idsostav,note,coefficient,vidpay,type,certificate,year
            //having(
            //(case when(
            //(case when (SUBSTRING(note,0,3)='>=' or SUBSTRING(note,0,3)='<=')then SUBSTRING(note,0,3) 
            //else SUBSTRING(note,0,2) end)='=') then 
            //(case when(
            //(case when (SUBSTRING(note,0,3)='>=' or SUBSTRING(note,0,3)='<=')then SUBSTRING(note,3,len(note)) 
            //else SUBSTRING(note,2,len(note)) end)= (Select count(sostav.idsostav) from sostav inner join treaty on sostav.idtreaty=treaty.idtreaty where sostav.idtreaty={dataGridView1.CurrentRow.Cells[0].Value.ToString()} ))
            //then coefficient else 
            //0
            //end
            //)
            //else (case when(
            //(case when (SUBSTRING(note,0,3)='>=' or SUBSTRING(note,0,3)='<=')then SUBSTRING(note,0,3) 
            //else SUBSTRING(note,0,2) end)='>') then 
            //(case when(
            //(case when (SUBSTRING(note,0,3)='>=' or SUBSTRING(note,0,3)='<=')then SUBSTRING(note,3,len(note)) 
            //else SUBSTRING(note,2,len(note)) end)< (Select count(sostav.idsostav) from sostav inner join treaty on sostav.idtreaty=treaty.idtreaty where sostav.idtreaty={dataGridView1.CurrentRow.Cells[0].Value.ToString()} ))
            //then coefficient else 0 end
            //)
            //else 0 end) end)
            //)!=0
            //union
            //Select  train.name +'; '+cast(train.nomer as  varchar(50)),train.type+'; '+cast(train.certificate as varchar(7))+'; '+
            //cast(train.year as varchar(4))+' г.',
            //price as [Страховая стоимость], sostav.summa as [Страховая сумма],
            //cast( case when pay.vidpay='Наличные' then coefficient*0.24*1 else coefficient*0.24*0.9 end  as decimal(18,2)) 
            //as [Страховой тариф],cast( case when pay.vidpay='Наличные' then sostav.summa*(coefficient*0.24)*1 else sostav.summa*(coefficient*0.24)*0.9 end as decimal(18,2))  as [Страховая премия]  
            //from train,sostav,correctionfactor,pay
            //where train.idtrain=sostav.idtrain and  correctionfactor.name like '%средств%' and sostav.idtreaty=pay.idtreaty
            //and sostav.idtreaty={dataGridView1.CurrentRow.Cells[0].Value.ToString()} 
            //and coefficient in (
            //select TOP(Select count(sostav.idsostav) from sostav inner join treaty on sostav.idtreaty=treaty.idtreaty where sostav.idtreaty={dataGridView1.CurrentRow.Cells[0].Value.ToString()} ) 
            //(
            //(case when(
            //(case when (SUBSTRING(note,0,3)='>=' or SUBSTRING(note,0,3)='<=')then SUBSTRING(note,0,3) 
            //else SUBSTRING(note,0,2) end)='<') then 
            //(case when(
            //(case when (SUBSTRING(note,0,3)='>=' or SUBSTRING(note,0,3)='<=')then SUBSTRING(note,3,len(note)) 
            //else SUBSTRING(note,2,len(note)) end)> (Select count(sostav.idsostav) from sostav inner join treaty on sostav.idtreaty=treaty.idtreaty where sostav.idtreaty={dataGridView1.CurrentRow.Cells[0].Value.ToString()} ))
            //then coefficient else 
            //0
            //end
            //)
            //else (case when(
            //(case when (SUBSTRING(note,0,3)='>=' or SUBSTRING(note,0,3)='<=')then SUBSTRING(note,0,3) 
            //else SUBSTRING(note,0,2) end)='<=') then 
            //(case when(
            //(case when (SUBSTRING(note,0,3)='>=' or SUBSTRING(note,0,3)='<=')then SUBSTRING(note,3,len(note)) 
            //else SUBSTRING(note,2,len(note)) end)>= (Select count(sostav.idsostav) from sostav inner join treaty on sostav.idtreaty=treaty.idtreaty where sostav.idtreaty={dataGridView1.CurrentRow.Cells[0].Value.ToString()} ))
            //then coefficient else 0 end
            //)
            //else 0 end) end)
            //)
            //from train,sostav,correctionfactor
            //where  train.idtrain=sostav.idtrain and  correctionfactor.name like '%средств%'
            //and sostav.idtreaty=75
            //group by idsostav,sostav.idtreaty, sostav.idtrain, train.name ,train.nomer, 
            //price, sostav.summa ,sostav.idsostav,note,coefficient
            //order by correctionfactor.note Desc
            //)
            //group by idsostav,sostav.idtreaty, sostav.idtrain, train.name ,train.nomer, 
            //price, sostav.summa ,sostav.idsostav,note,coefficient,vidpay,certificate,type,year
            //having (
            //(case when(
            //(case when (SUBSTRING(note,0,3)='>=' or SUBSTRING(note,0,3)='<=')then SUBSTRING(note,0,3) 
            //else SUBSTRING(note,0,2) end)='<') then 
            //(case when(
            //(case when (SUBSTRING(note,0,3)='>=' or SUBSTRING(note,0,3)='<=')then SUBSTRING(note,3,len(note)) 
            //else SUBSTRING(note,2,len(note)) end)>(Select count(sostav.idsostav) from sostav inner join treaty on sostav.idtreaty=treaty.idtreaty where sostav.idtreaty={dataGridView1.CurrentRow.Cells[0].Value.ToString()} ))
            //then coefficient else 
            //0
            //end
            //)
            //else (case when(
            //(case when (SUBSTRING(note,0,3)='>=' or SUBSTRING(note,0,3)='<=')then SUBSTRING(note,0,3) 
            //else SUBSTRING(note,0,2) end)='<=') then 
            //(case when(
            //(case when (SUBSTRING(note,0,3)='>=' or SUBSTRING(note,0,3)='<=')then SUBSTRING(note,3,len(note)) 
            //else SUBSTRING(note,2,len(note)) end)>= (Select count(sostav.idsostav) from sostav inner join treaty on sostav.idtreaty=treaty.idtreaty where sostav.idtreaty={dataGridView1.CurrentRow.Cells[0].Value.ToString()} ))
            //then coefficient else 0 end
            //)
            //else 0 end) end)
            //) !=0

            //union
            //Select  train.name +'; '+cast(train.nomer as  varchar(50)),train.type+'; '+cast(train.certificate as varchar(7))+'; '+
            //cast(train.year as varchar(4))+' г.',
            //price as [Страховая стоимость], sostav.summa as [Страховая сумма],
            //cast( case when pay.vidpay='Наличные' then coefficient*0.24*1 else coefficient*0.24*0.9 end  as decimal(18,2)) 
            //as [Страховой тариф],cast( case when pay.vidpay='Наличные' then sostav.summa*(coefficient*0.24)*1 else sostav.summa*(coefficient*0.24)*0.9 end as decimal(18,2))  as [Страховая премия]  
            //from train,sostav,correctionfactor,pay
            //where  train.idtrain=sostav.idtrain and  correctionfactor.name like '%средств%' and sostav.idtreaty=pay.idtreaty
            //and sostav.idtreaty={dataGridView1.CurrentRow.Cells[0].Value.ToString()} 
            //group by idsostav,sostav.idtreaty, sostav.idtrain, train.name ,train.nomer, 
            //price, sostav.summa ,sostav.idsostav,note,coefficient,vidpay,certificate,year,type
            //having (
            //select TOP(1) 
            //(
            //(case when(
            //(case when (SUBSTRING(note,0,3)='>=' or SUBSTRING(note,0,3)='<=')then SUBSTRING(note,0,3) 
            //else SUBSTRING(note,0,2) end)='<') then 
            //(case when(
            //(case when (SUBSTRING(note,0,3)='>=' or SUBSTRING(note,0,3)='<=')then SUBSTRING(note,3,len(note)) 
            //else SUBSTRING(note,2,len(note)) end)> (Select count(sostav.idsostav) from sostav inner join treaty on sostav.idtreaty=treaty.idtreaty where sostav.idtreaty={dataGridView1.CurrentRow.Cells[0].Value.ToString()} ))
            //then coefficient else 
            //1
            //end
            //)
            //else (case when(
            //(case when (SUBSTRING(note,0,3)='>=' or SUBSTRING(note,0,3)='<=')then SUBSTRING(note,0,3) 
            //else SUBSTRING(note,0,2) end)='<=') then 
            //(case when(
            //(case when (SUBSTRING(note,0,3)='>=' or SUBSTRING(note,0,3)='<=')then SUBSTRING(note,3,len(note)) 
            //else SUBSTRING(note,2,len(note)) end)>= (Select count(sostav.idsostav) from sostav inner join treaty on sostav.idtreaty=treaty.idtreaty where sostav.idtreaty={dataGridView1.CurrentRow.Cells[0].Value.ToString()} ))
            //then coefficient else 1 end
            //)
            //else 1 end) end)
            //)
            //from train,sostav,correctionfactor
            //where train.idtrain=sostav.idtrain and  correctionfactor.name like '%средств%'
            //and sostav.idtreaty={dataGridView1.CurrentRow.Cells[0].Value.ToString()} 
            //group by idsostav,sostav.idtreaty, sostav.idtrain, train.name ,train.nomer, 
            //price, sostav.summa ,sostav.idsostav,note,coefficient
            //order by correctionfactor.note Desc)=1
            //", sqlConnection);
            //                    command.Fill(dataSet);
            //                    dataGridView2.DataSource = dataSet.Tables[0].DefaultView;
            //                    sqlConnection.Close();
            //                    dataGridView2.AllowUserToAddRows = false;
            //                    sqlConnection.Close();
            //                    dataGridView2.RowsDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            //                    dataGridView2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            //                    dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            //                    strax = "";
            //                    strax2 = "";
            //                    pas = "";
            //                    vznos = "";
            //                    Word.Application wordApplication = new Word.Application();

            //                    string PathToNote = @"/Diplom/proga/Agent/treaty/train/train.docx";
            //                    Word.Application application = new Word.Application();
            //                    application.Visible = false;
            //                    Word.Document document = application.Documents.Open(PathToNote);
            //                    object oMissing = System.Reflection.Missing.Value;
            //                    application.Selection.Find.Execute("%метка%");
            //                    Word.Range wordRange = application.Selection.Range;

            //                    int RowCount2 = dataGridView2.RowCount + 2;
            //                    int ColumnCount2 = 7;
            //                    System.Object defaultTableBehavior2 =
            //                           Word.WdDefaultTableBehavior.wdWord9TableBehavior;
            //                    System.Object autoFitBehavior2 = Word.WdAutoFitBehavior.wdAutoFitWindow;
            //                    Word.Table wordtable2 = document.Tables.Add(application.Selection.Range, RowCount2, ColumnCount2,
            //                                      ref defaultTableBehavior2, ref autoFitBehavior2);
            //                    System.Object style2 = "Сетка таблицы";
            //                    wordtable2.set_Style(ref style2);
            //                    wordtable2.ApplyStyleFirstColumn = true;
            //                    wordtable2.ApplyStyleHeadingRows = true;
            //                    Word.Range wordcellrange10 = document.Tables[1].Cell(1, 1).Range;
            //                    wordcellrange10.Text = "№ п/п";
            //                    wordcellrange10.Font.Size = 10;

            //                    Word.Range wordcellrange11 = document.Tables[1].Cell(1, 2).Range;
            //                    wordcellrange11.Text = "Наименование средств транспорта, номера";
            //                    wordcellrange11.Font.Size = 10;

            //                    Word.Range wordcellrange12 = document.Tables[1].Cell(1, 3).Range;
            //                    wordcellrange12.Text = "Тип транспорта, сертификат, год постройки";
            //                    wordcellrange12.Font.Size = 10;

            //                    Word.Range wordcellrange13 = document.Tables[1].Cell(1, 4).Range;
            //                    wordcellrange13.Text = "Страховая стоимость (руб.)";
            //                    wordcellrange13.Font.Size = 10;

            //                    Word.Range wordcellrange14 = document.Tables[1].Cell(1, 5).Range;
            //                    wordcellrange14.Text = "Страховая сумма (руб.)";
            //                    wordcellrange14.Font.Size = 10;

            //                    Word.Range wordcellrange15 = document.Tables[1].Cell(1, 6).Range;
            //                    wordcellrange15.Text = "Тариф (%)";
            //                    wordcellrange15.Font.Size = 10;

            //                    Word.Range wordcellrange16 = document.Tables[1].Cell(1, 7).Range;
            //                    wordcellrange16.Text = "Страховая премия (руб.)";
            //                    wordcellrange16.Font.Size = 10;

            //                    document.Tables[1].Rows[RowCount2].Cells[1].Merge(document.Tables[1].Rows[RowCount2].Cells[3]);
            //                    document.Tables[1].Cell(RowCount2, 1).Range.Text = $@"Итого:";

            //                    decimal kprice = 0;
            //                    decimal ksumm = 0;

            //                    decimal kpre = 0;
            //                    for (int m2 = 2; m2 < RowCount2; m2++)
            //                    {
            //                        wordcellrange10 = wordtable2.Cell(m2, 1).Range;
            //                        wordcellrange10.Text = Convert.ToString(m2 - 1);
            //                        wordcellrange10.Font.Size = 10;
            //                        wordcellrange10 = wordtable2.Cell(m2, 2).Range;
            //                        wordcellrange10.Text = dataGridView2[0, m2 - 2].Value.ToString();
            //                        wordcellrange10.Font.Size = 10;
            //                        wordcellrange10 = wordtable2.Cell(m2, 3).Range;
            //                        wordcellrange10.Text = dataGridView2[1, m2 - 2].Value.ToString();
            //                        wordcellrange10.Font.Size = 10;
            //                        wordcellrange10 = wordtable2.Cell(m2, 4).Range;
            //                        wordcellrange10.Text = dataGridView2[2, m2 - 2].Value.ToString();
            //                        wordcellrange10.Font.Size = 10;
            //                        kprice = kprice + Convert.ToDecimal(dataGridView2[2, m2 - 2].Value.ToString());
            //                        wordcellrange10 = wordtable2.Cell(m2, 5).Range;
            //                        wordcellrange10.Text = dataGridView2[3, m2 - 2].Value.ToString();
            //                        wordcellrange10.Font.Size = 10;
            //                        ksumm = ksumm + Convert.ToDecimal(dataGridView2[3, m2 - 2].Value.ToString());
            //                        wordcellrange10 = wordtable2.Cell(m2, 6).Range;
            //                        wordcellrange10.Text = dataGridView2[4, m2 - 2].Value.ToString();
            //                        wordcellrange10.Font.Size = 10;
            //                        document.Tables[1].Cell(RowCount2, 4).Range.Text = Convert.ToString(dataGridView2[4, m2 - 2].Value.ToString());
            //                        wordcellrange10 = wordtable2.Cell(m2, 7).Range;
            //                        wordcellrange10.Text = dataGridView2[5, m2 - 2].Value.ToString();
            //                        wordcellrange10.Font.Size = 10;
            //                        kpre = kpre + Convert.ToDecimal(dataGridView2[5, m2 - 2].Value.ToString());
            //                    }
            //                    document.Tables[1].Cell(RowCount2, 2).Range.Text = Convert.ToString(kprice);
            //                    document.Tables[1].Cell(RowCount2, 3).Range.Text = Convert.ToString(ksumm);
            //                    document.Tables[1].Cell(RowCount2, 5).Range.Text = Convert.ToString(kpre);
            //                    ReplaceWordStubs("{номер}", dataGridView1.CurrentRow.Cells[0].Value.ToString(), document);
            //                    ReplaceWordStubs("{дата}", dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(0, 10), document);

            //                    string query1 = $@"Select 
            //position.name as [Должность]
            //from policyholder inner join city on policyholder.idcity=city.idcity 
            //inner join position on position.idposition=policyholder.idwork 
            //inner join work on work.idwork=position.idwork
            // where idpolicyholder={dataGridView1.CurrentRow.Cells[3].Value.ToString()}";
            //                    System.Data.DataTable data = new System.Data.DataTable();
            //                    SqlDataAdapter command1 = new SqlDataAdapter(query1, sqlConnection);
            //                    command1.Fill(data);
            //                    DataColumn column = data.Columns[0];
            //                    DataRow row = data.Rows[0];
            //                    ReplaceWordStubs("{страхователь}", dataGridView1.CurrentRow.Cells[4].Value.ToString(), document);
            //                    ReplaceWordStubs("{страхс}", dataGridView1.CurrentRow.Cells[11].Value.ToString(), document);
            //                    ReplaceWordStubs("{страхп}", dataGridView1.CurrentRow.Cells[15].Value.ToString(), document);

            //                    if (dataGridView1.CurrentRow.Cells[20].Value.ToString() != "")
            //                    {
            //                        ReplaceWordStubs("{датав}", dataGridView1.CurrentRow.Cells[17].Value.ToString() + " BYN " + dataGridView1.CurrentRow.Cells[18].Value.ToString().Substring(0, 10) + " г. ; " + dataGridView1.CurrentRow.Cells[17].Value.ToString() + " BYN " + dataGridView1.CurrentRow.Cells[20].Value.ToString().Substring(0, 10) + " г. ; " + dataGridView1.CurrentRow.Cells[17].Value.ToString() + " BYN " + dataGridView1.CurrentRow.Cells[22].Value.ToString().Substring(0, 10), document);

            //                    }
            //                    else if (dataGridView1.CurrentRow.Cells[18].Value.ToString() != "")
            //                    {
            //                        ReplaceWordStubs("{датав}", dataGridView1.CurrentRow.Cells[17].Value.ToString() + " BYN " + dataGridView1.CurrentRow.Cells[18].Value.ToString().Substring(0, 10), document);
            //                    }
            //                    else
            //                    {
            //                        ReplaceWordStubs("{датав}", dataGridView1.CurrentRow.Cells[17].Value.ToString() + " BYN " + dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(0, 10), document);
            //                    }
            //                    string query45 = $@"select date from treaty 
            //inner join bid on bid.idbid=treaty.idbid where treaty.idbid={dataGridView1.CurrentRow.Cells[24].Value.ToString()}";
            //                    System.Data.DataTable data45 = new System.Data.DataTable();
            //                    SqlDataAdapter command45 = new SqlDataAdapter(query45, sqlConnection);
            //                    command45.Fill(data45);
            //                    DataColumn column45 = data45.Columns[0];
            //                    DataRow row45 = data45.Rows[0];
            //                    ReplaceWordStubs("{датаз}", row45[column45].ToString().Substring(0, 10), document);
            //                    ReplaceWordStubs("{номерз}", dataGridView1.CurrentRow.Cells[24].Value.ToString(), document);
            //                    ReplaceWordStubs("{дата}", dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(0, 10), document);
            //                    ReplaceWordStubs("{дата}", dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(0, 10), document);
            //                    ReplaceWordStubs("{дата}", dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(0, 10), document);
            //                    ReplaceWordStubs("{страхователь}", dataGridView1.CurrentRow.Cells[4].Value.ToString(), document);
            //                    ReplaceWordStubs("{страховщик}", dataGridView1.CurrentRow.Cells[2].Value.ToString(), document);
            //                    application.Visible = true;
            //                }
            //                else
            //                {
            //                    MessageBox.Show("Нет шаблона договора!", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //                }
            //            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            word();
        }
        int visible = 0;
        public void four()
        {

            button11.Visible = false;

            panel3.Visible = false;

            panel6.Visible = false;
            dataGridView1.Enabled = true;
            clear();
        

        }
        private void button5_Click(object sender, EventArgs e)
        {
            //try
            //{
            //    sqlConnection.Open();
            //    string query = $@"DELETE FROM [sostav] WHERE [idtreaty] is null";
            //    SqlCommand command = new SqlCommand(query, sqlConnection);
            //    command.ExecuteNonQuery();
            //    sqlConnection.Close();

            //}
            //catch { }

            button11.Visible = false;

            panel6.Visible = false;
            dataGridView1.Enabled = true;
      

            panel4.Visible = false;
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
            ExcelApp.Rows[1].Columns[7] = "Договора";
            ExcelApp.Rows[visible + 3].Columns[7] = "Ридецкая Анна Михайловна";
            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                ExcelApp.Cells[2, i + 1] = dataGridView1.Columns[i].HeaderText;

            }
            int y = 0;
            for (int j = 0; j < dataGridView1.ColumnCount; j++)
            {
                y = 0;
                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    if (dataGridView1.Rows[i].Cells[j].Visible == true)
                    {
                        if (j == 7 || j == 9 || j == 10 || j == 16)
                        {
                            ExcelApp.Cells[y + 3, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString().Substring(0, 10);
                            y++;
                        }
                        else if (j == 18 && dataGridView1.Rows[i].Cells[j].Value.ToString() != "")
                        {
                            ExcelApp.Cells[y + 3, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString().Substring(0, 10);
                            y++;
                        }
                        else if ((j == 20 || j == 22) && dataGridView1.Rows[i].Cells[j].Value.ToString() != "")
                        {
                            ExcelApp.Cells[y + 3, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString().Substring(0, 10);
                            y++;
                        }
                        else
                        {
                            ExcelApp.Cells[y + 3, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                            y++;
                        }
                    }
                }
            }
            for (int i = 0; i < visible; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    Microsoft.Office.Interop.Excel.Range range = ExcelWorkSheet.Range[$"A1:X{visible + 3}"];
                    ExcelWorkSheet.Range[$"A1:X{visible + 3}"].Cells.WrapText = true;
                    range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders.Weight = Excel.XlBorderWeight.xlMedium;
                }
            }
            ExcelApp.Columns["B"].Delete();
            ExcelApp.Columns["C"].Delete();
            ExcelApp.Columns["D"].Delete();
            ExcelApp.Columns["V"].Delete();
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }
        int id = 0;
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
           
            id = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
            //if (policyholder != null)
            //{
            //    policyholder.Height = 638;
            //    policyholder.panel6.Height = 595;
            //    clear();
            //}
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && e.KeyChar != 8)
                e.Handled = true;
        }

        string kofi = "";

        decimal sumid = 0;
        private void button12_Click(object sender, EventArgs e)
        {
            //if (panel7.Visible == true || panel11.Visible == true)
            //{
            //    panel6.Visible = true;
            //    panel7.Visible = false;
            //    panel2.Visible = false;
            //    button12.Visible = false;
            //    panel11.Visible = false;
            //    button10.Top = 771;
            //    button10.Left = 1241;
            //    dataGridView1.Enabled = true;
            //    dataGridView1.Height = 407;
            //    button11.Visible = false;
            //    if (id2 != 0)
            //    {
            //        button10.Visible = true;
            //    }
            //    else
            //    {
            //        button10.Visible = false;
            //    }
            //    six();
            //    for (int i = 0; i < dataGridView3.RowCount; i++)
            //    {
            //        dataGridView3.Rows[i].Selected = false;
            //    }
            //    for (int i = 0; i < dataGridView3.RowCount; i++)
            //    {
            //        if (id2 == Convert.ToInt32(dataGridView3[0, i].Value.ToString()))
            //        {
            //            dataGridView3.Rows[i].Selected = true;
            //            dataGridView3.Rows[i].DefaultCellStyle.SelectionForeColor = Color.Black;
            //            dataGridView3.Rows[i].DefaultCellStyle.SelectionBackColor = Color.FromArgb(212, 236, 252);
            //            break;
            //        }
            //    }
            //}
            //else if (panel2.Visible == true)
            //{
            //    if (x == 0 && dataGridView3[2, 0].Value.ToString() != "Страхование средств железнодорожного транспорта")
            //    {
            //        panel2.Visible = false;
            //        panel6.Visible = false;
            //        panel7.Visible = true;
            //        button12.Visible = true;
            //        button12.Top = 740;
            //        button12.Left = 403;
            //        button10.Visible = true;
            //        button10.Top = 740;
            //        button10.Left = 773;
            //        button11.Visible = false;
            //        dataGridView2.Top = 69;
            //        dataGridView1.Height = 407;
            //        dataGridView2.Width = 950;
            //        dataGridView2.Height = 297;
            //    }
            //    else
            //    {
            //        panel2.Visible = false;
            //        panel6.Visible = false;
            //        panel7.Visible = false;
            //        panel11.Visible = true;
            //        button12.Visible = true;
            //        button12.Top = 771;
            //        button12.Left = 0;
            //        button10.Visible = true;
            //        button10.Top = 771;
            //        button10.Left = 1241;
            //        button11.Visible = false;
            //        if (comboBox2.Enabled == true)
            //        {
            //            one();
            //            two();
            //            dataGridView1.BringToFront();
            //            dataGridView2.Top = 69;
            //            dataGridView1.Height = 407;
            //            dataGridView2.Width = 950;
            //            dataGridView2.Height = 297;
            //        }
            //        else
            //        {
            //            dataGridView2.Top = 5;
            //            dataGridView1.Height = 345;
            //            dataGridView2.Width = 1402;
            //            dataGridView2.Height = 362;
            //        }
            //    }
            //}
        }

        private void button10_Click_1(object sender, EventArgs e)
        {
            //    if (panel6.Visible == true)
            //    {
            //        if (x == 0 && dataGridView3[2, 0].Value.ToString() != "Страхование средств железнодорожного транспорта")
            //        {
            //            panel2.Visible = false;
            //            panel6.Visible = false;
            //            panel11.Visible = false;
            //            panel7.Visible = true;
            //            button12.Visible = true;
            //            button12.Top = 740;
            //            button12.Left = 403;
            //            button10.Visible = true;
            //            button10.Top = 740;
            //            button10.Left = 773;
            //            textBox5.Visible = true;
            //            label10.Visible = true;
            //            dataGridView2.Top = 5;
            //            dataGridView1.Height = 407;
            //            six();
            //        }
            //        else
            //        {
            //            panel2.Visible = false;
            //            panel6.Visible = false;
            //            panel7.Visible = false;
            //            panel11.Visible = true;
            //            button12.Visible = true;
            //            button12.Top = 771;
            //            button12.Left = 0;
            //            button10.Visible = true;
            //            button10.Top = 771;
            //            button10.Left = 1241;
            //            textBox5.Visible = true;
            //            label10.Visible = true;
            //            if (comboBox5.Enabled == true)
            //            {
            //                dataGridView1.Enabled = true;
            //                dataGridView1.Height = 407;
            //                count = 0;
            //                id2 = Convert.ToInt32(dataGridView3.CurrentRow.Cells[0].Value.ToString());
            //                id3 = Convert.ToInt32(dataGridView3.CurrentRow.Cells[1].Value.ToString());
            //                for (int i = 0; i < dataGridView1.RowCount; i++)
            //                {
            //                    for (int j = 0; j < dataGridView3.RowCount; j++)
            //                    {
            //                        try
            //                        {
            //                            if (Convert.ToDateTime(dataGridView1[10, i].Value.ToString()) <= DateTime.Today && Convert.ToInt32(comboBox2.SelectedValue) == Convert.ToInt32(dataGridView1[3, j].Value.ToString()))
            //                            { count++; }
            //                        }
            //                        catch { }
            //                    }
            //                }
            //                one();
            //                two();
            //                dataGridView1.BringToFront();
            //                dataGridView2.Top = 69;
            //                dataGridView2.Width = 950;
            //                dataGridView2.Height = 297;
            //            }
            //            else
            //            {
            //                count = 0;
            //                id2 = Convert.ToInt32(dataGridView3[0, 0].Value.ToString());
            //                id3 = Convert.ToInt32(dataGridView3[1, 0].Value.ToString());
            //                dataGridView1.BringToFront();
            //                dataGridView1.Height = 345;
            //                dataGridView2.Top = 5;
            //                dataGridView2.Width = 1402;
            //                dataGridView2.Height = 362;
            //            }
            //        }
            //    }
            //    else if (panel7.Visible == true)
            //    {
            //        ten();  
            //        dataGridView1.Enabled = true;
            //    }
            //    else if (panel11.Visible == true)
            //    {
            //        dataGridView1.Height = 407;
            //    }
            //}
            //public void ten()
            //{
            //        panel7.Visible = false;
            //        panel11.Visible = false;
            //        panel6.Visible = false;
            //        panel2.Visible = true;
            //        button12.Visible = true;
            //        button12.Top = 744;
            //        button12.Left = 120;
            //        if (comboBox2.Enabled == true)
            //        {
            //            button11.Visible = true;
            //            button11.Top = 744;
            //            button11.Left = 1107;
            //        }
            //        else
            //        {
            //            button11.Visible = false;
            //        }
            //        button10.Visible = false;
            //        six();
        }
        int count = 0;
        int x = 0;
        private void dataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            count = 0;
            try
            {
                id2 = Convert.ToInt32(dataGridView3.CurrentRow.Cells[0].Value.ToString());


                
            }
            catch { }
        }

        public void comboBoxvid()
        {
            try
            {
                sqlConnection.Close();
                sqlConnection.Open();
                DataSet dataSet = new DataSet();
                SqlDataAdapter command = new SqlDataAdapter($@"select bid.idbid,vid.idvida, vid.name as Название, bid.note as Пожелания from vid, policyholder,bid " +
                    $@"where bid.idvida=vid.idvida and bid.idpolicyholder=policyholder.idpolicyholder and policyholder.idpolicyholder={comboBox2.SelectedValue} and bid.status='Ожидание' ", sqlConnection);
                command.Fill(dataSet);
                dataGridView3.DataSource = dataSet.Tables[0].DefaultView;
                sqlConnection.Close();
                dataGridView3.Columns[0].Visible = false;
                dataGridView3.Columns[1].Visible = false;
                dataGridView3.AllowUserToAddRows = false;
                sqlConnection.Close();
            }
            catch
            {
                sqlConnection.Close();
                sqlConnection.Open();
                DataSet dataSet = new DataSet();
                SqlDataAdapter command = new SqlDataAdapter($@"select bid.idbid,vid.idvida, vid.name as Название, bid.note as Пожелания from vid, policyholder,bid " +
                    $@"where bid.idvida=vid.idvida and bid.idpolicyholder=policyholder.idpolicyholder and policyholder.idpolicyholder=0 and bid.status='Ожидание' ", sqlConnection);
                command.Fill(dataSet);
                dataGridView3.DataSource = dataSet.Tables[0].DefaultView;
                sqlConnection.Close();
                dataGridView3.Columns[0].Visible = false;
                dataGridView3.Columns[1].Visible = false;
                dataGridView3.AllowUserToAddRows = false;
                sqlConnection.Close();
            }

        }
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                comboBoxvid();
            }
            catch { }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            podsvetka();
           
            four();
            visible = 0;
            if (insurerak != null)
            {
                word();
            }
            else
            {
                panel4.BringToFront();
                panel4.Visible = true;
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            panel4.Visible = false;
        }

        private Word.Paragraph wordparagraph;
        string str = "";
        int c = 0;
        int p = 0;
        public void countroe()
        {
            try
            {
                for (int m = 0; m < dataGridView1.RowCount; m++)
                {

                    if (dataGridView1.Rows[m].DefaultCellStyle.SelectionBackColor == Color.FromArgb(255, 250, 170, 209))
                    {
                        c++; p++;
                    }
                }
            }
            catch { }

        }
        int c2 = 0;
        int p2 = 0;
        public void countroe2()
        {
            try
            {
                for (int m = 0; m < dataGridView1.RowCount; m++)
                {
                    string query3 = $@"Select treaty.idtreaty as [Номер договора],(insurer.firstname+' '+insurer.name+' '+insurer.lastname) as [Страховщик],
            (policyholder.firdtname+' '+policyholder.name+' '+policyholder.lastname) as Страхователь,
            DATEDIFF(Day,GETDATE(),datefinish) as [Дата окончания действия договора] 
            from treaty 
inner join insurer on treaty.idinsurer=insurer.idinsurer inner join bid on treaty.idbid=bid.idbid
inner join policyholder on bid.idpolicyholder=policyholder.idpolicyholder  inner join vid on bid.idvida=vid.idvida
inner join correctionfactor on correctionfactor.idvida=vid.idvida inner join pay on pay.idtreaty=treaty.idtreaty
            group by treaty.idtreaty ,(insurer.firstname+' '+insurer.name+' '+insurer.lastname) ,
            (policyholder.firdtname+' '+policyholder.name+' '+policyholder.lastname),datefinish
            having DATEDIFF(Day,GETDATE(),datefinish)<=31";
                    System.Data.DataTable data3 = new System.Data.DataTable();
                    SqlDataAdapter command3 = new SqlDataAdapter(query3, sqlConnection);
                    command3.Fill(data3);
                    DataColumn column2 = data3.Columns[0];
                    DataRow row2 = data3.Rows[p2];
                    if (dataGridView1[0, m].Value.ToString() == row2[column2].ToString())
                    {
                        c2++; p2++;
                    }
                }
            }
            catch { }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            //Вывод в "Word" список договоров с оканчивающимся сроком
            countroe2();
            string PathToNote = System.IO.Path.GetFullPath(@"/Diplom/proga/Agent/Docdocx1.docx");
            Word.Application application = new Word.Application();
            application.Visible = false;
            Word.Document oDoc = application.Documents.Open(PathToNote);
            object oMissing = System.Reflection.Missing.Value;
            oDoc.Paragraphs.Add(ref oMissing);
            wordparagraph = oDoc.Paragraphs[1];
            Word.Range wordrange = wordparagraph.Range;
            wordrange.Text = "Список заканчивающих договоров";
            wordparagraph.Range.Font.Size = 24;
            oDoc.Paragraphs.Add(ref oMissing);
            wordparagraph = oDoc.Paragraphs[2];
            Word.Range wordrange1 = wordparagraph.Range;
            wordrange1.Text = "Месяц: " + DateTime.Today.AddMonths(1).ToString("MMMM", CultureInfo.GetCultureInfo("ru-RU"));
            wordparagraph.Range.Font.Size = 18;
            oDoc.Paragraphs.Add(ref oMissing);
            wordparagraph = oDoc.Paragraphs[3];
            Word.Range wordrange3 = wordparagraph.Range;
            wordparagraph.Range.Font.Size = 12;
            int RowCount2 = c2 + 1;
            int ColumnCount2 = 6;
            object unit2;
            object extend2;
            unit2 = Word.WdUnits.wdStory;
            extend2 = Word.WdMovementType.wdMove;
            application.Selection.EndKey(ref unit2, ref extend2);
            System.Object defaultTableBehavior2 =
                   Word.WdDefaultTableBehavior.wdWord9TableBehavior;
            System.Object autoFitBehavior2 = Word.WdAutoFitBehavior.wdAutoFitWindow;
            Word.Table wordtable2 = oDoc.Tables.Add(application.Selection.Range, RowCount2, ColumnCount2,
                              ref defaultTableBehavior2, ref autoFitBehavior2);
            System.Object style2 = "Сетка таблицы";
            wordtable2.set_Style(ref style2);
            wordtable2.ApplyStyleFirstColumn = true;
            wordtable2.ApplyStyleHeadingRows = true;
            Word.Range wordcellrange10 = oDoc.Tables[1].Cell(1, 1).Range;
            wordcellrange10.Text = "Номер договора";
            wordcellrange10.Font.Size = 12;

            Word.Range wordcellrange11 = oDoc.Tables[1].Cell(1, 2).Range;
            wordcellrange11.Text = "Страховщик";
            wordcellrange11.Font.Size = 12;

            Word.Range wordcellrange12 = oDoc.Tables[1].Cell(1, 3).Range;
            wordcellrange12.Text = "Страхователь";
            wordcellrange12.Font.Size = 12;

            Word.Range wordcellrange13 = oDoc.Tables[1].Cell(1, 4).Range;
            wordcellrange13.Text = "Вид страхования";
            wordcellrange13.Font.Size = 12;

            Word.Range wordcellrange14 = oDoc.Tables[1].Cell(1, 5).Range;
            wordcellrange14.Text = "Количество дней до конца";
            wordcellrange14.Font.Size = 12;

            Word.Range wordcellrange15 = oDoc.Tables[1].Cell(1, 6).Range;
            wordcellrange15.Text = "Дата окончания";
            wordcellrange15.Font.Size = 12;

            int n2 = 0;
            int i2 = 0;
            for (int m2 = 2; m2 < RowCount2 + 1; m2++)
            {
                if (n2 < (RowCount2))
                {

                    string query3 = $@"Select Distinct treaty.idtreaty as [Номер договора],(insurer.firstname+' '+insurer.name+' '+insurer.lastname) as [Страховщик],
(policyholder.firdtname+' '+policyholder.name+' '+policyholder.lastname) as Страхователь,
vid.name as [Вид страхования],
datefinish as [Дата окончания действия договора],
DATEDIFF(Day,GETDATE(),datefinish) as [Дата окончания действия договора] 
from
treaty inner join insurer on treaty.idinsurer=insurer.idinsurer
inner join bid on bid.idbid=treaty.idbid
inner join vid on vid.idvida=bid.idvida
inner join policyholder  on policyholder.idpolicyholder=bid.idpolicyholder
inner join pay on treaty.idtreaty=pay.idtreaty
left join correctionfactor on correctionfactor.idvida=vid.idvida
group by
treaty.idtreaty,(insurer.firstname+' '+insurer.name+' '+insurer.lastname),
vid.name,
(policyholder.firdtname+' '+policyholder.name+' '+policyholder.lastname),
 datefinish 
 having DATEDIFF(Day,GETDATE(),datefinish)<=31";
                    System.Data.DataTable data3 = new System.Data.DataTable();
                    SqlDataAdapter command3 = new SqlDataAdapter(query3, sqlConnection);
                    command3.Fill(data3);
                    DataColumn column2 = data3.Columns[0];
                    DataRow row2 = data3.Rows[n2];

                    if (Convert.ToInt32(dataGridView1[0, i2].Value.ToString()) == Convert.ToInt32(row2[column2].ToString()))
                    {
                        DataColumn column3 = data3.Columns[0];
                        DataRow row3 = data3.Rows[n2];
                        wordcellrange10 = wordtable2.Cell(m2, 1).Range;
                        wordcellrange10.Text = row3[column3].ToString();
                        wordcellrange10.Font.Size = 12;
                        DataColumn column4 = data3.Columns[1];
                        DataRow row4 = data3.Rows[n2];
                        wordcellrange10 = wordtable2.Cell(m2, 2).Range;
                        wordcellrange10.Text = row4[column4].ToString();
                        wordcellrange10.Font.Size = 12;
                        DataColumn column5 = data3.Columns[2];
                        DataRow row5 = data3.Rows[n2];
                        wordcellrange10 = wordtable2.Cell(m2, 3).Range;
                        wordcellrange10.Text = row5[column5].ToString();
                        wordcellrange10.Font.Size = 12;
                        DataColumn column6 = data3.Columns[3];
                        DataRow row6 = data3.Rows[n2];
                        wordcellrange10 = wordtable2.Cell(m2, 4).Range;
                        wordcellrange10.Text = row6[column6].ToString();
                        wordcellrange10.Font.Size = 12;
                        DataColumn column7 = data3.Columns[5];
                        DataRow row7 = data3.Rows[n2];
                        wordcellrange10 = wordtable2.Cell(m2, 5).Range;
                        wordcellrange10.Text = row7[column7].ToString();
                        wordcellrange10.Font.Size = 12;
                        DataColumn column8 = data3.Columns[4];
                        DataRow row8 = data3.Rows[n2];
                        wordcellrange10 = wordtable2.Cell(m2, 6).Range;
                        wordcellrange10.Text = row8[column8].ToString().Substring(0, 10);
                        wordcellrange10.Font.Size = 12;
                        n2++;
                    }
                    else { i2++; m2--; }
                }
            }
            i2 = 0; n2 = 0;
            c2 = 0; p2 = 0;
            application.Visible = true;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            countroe();
            str = "";
            try
            {
                //Вывод в "Word" список договоров с очередными взносами
                string PathToNote = @"/Diplom/proga/Agent/Docdocx.docx";
                Word.Application application = new Word.Application();
                application.Visible = false;
                Word.Document oDoc = application.Documents.Open(PathToNote);
                object oMissing = System.Reflection.Missing.Value;
                oDoc.Paragraphs.Add(ref oMissing);
                wordparagraph = oDoc.Paragraphs[1];
                Word.Range wordrange = wordparagraph.Range;
                wordrange.Text = "Список договоров с очередным взносом";
                wordrange.Font.Size = 24;
                oDoc.Paragraphs.Add(ref oMissing);
                wordparagraph = oDoc.Paragraphs[2];
                Word.Range wordrange1 = wordparagraph.Range;
                wordrange1.Text = "Месяц: " + DateTime.Today.AddMonths(1).ToString("MMMM", CultureInfo.GetCultureInfo("ru-RU"));
                wordparagraph.Range.Font.Size = 18;
                oDoc.Paragraphs.Add(ref oMissing);
                wordparagraph = oDoc.Paragraphs[3];
                Word.Range wordrange2 = wordparagraph.Range;
                wordparagraph.Range.Font.Size = 12;
                int RowCount = c + 1;
                int ColumnCount = 5;
                object unit;
                object extend;
                unit = Word.WdUnits.wdStory;
                extend = Word.WdMovementType.wdMove;
                application.Selection.EndKey(ref unit, ref extend);
                System.Object defaultTableBehavior =
                       Word.WdDefaultTableBehavior.wdWord9TableBehavior;
                System.Object autoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitWindow;
                Word.Table wordtable = oDoc.Tables.Add(application.Selection.Range, RowCount, ColumnCount,
                                  ref defaultTableBehavior, ref autoFitBehavior);
                System.Object style = "Сетка таблицы";
                wordtable.set_Style(ref style);
                wordtable.ApplyStyleFirstColumn = true;
                wordtable.ApplyStyleHeadingRows = true;
                Word.Range wordcellrange = oDoc.Tables[1].Cell(1, 1).Range;
                wordcellrange.Text = "Номер договора";
                wordcellrange.Font.Size = 12;

                Word.Range wordcellrange1 = oDoc.Tables[1].Cell(1, 2).Range;
                wordcellrange1.Text = "Страховщик";
                wordcellrange1.Font.Size = 12;

                Word.Range wordcellrange2 = oDoc.Tables[1].Cell(1, 3).Range;
                wordcellrange2.Text = "Страхователь";
                wordcellrange2.Font.Size = 12;

                Word.Range wordcellrange3 = oDoc.Tables[1].Cell(1, 4).Range;
                wordcellrange3.Text = "Сумма  взноса";
                wordcellrange3.Font.Size = 12;

                Word.Range wordcellrange4 = oDoc.Tables[1].Cell(1, 5).Range;
                wordcellrange4.Text = "Последний день оплаты";
                wordcellrange4.Font.Size = 12;

                int n = 0;
                int i = 0;
                for (int m = 2; m < RowCount + 1; m++)
                {
                    if (n < (RowCount - 1))
                    {
                        try
                        {
                            int hall = 0;
                            string hyll = "";
                            DateTime min = DateTime.MinValue;
                            DateTime max = DateTime.MaxValue;
                            string query2 = $@"select count(pay.idtreaty),treaty.idtreaty,dateconclusion,
(CASE WHEN(treaty.term='В два срока')
THEN  DATEADD(MONTH,6, datestart)
else (CASE WHEN(treaty.term='Ежеквартально')
THEN  DATEADD(MONTH,3, datestart)
else null end) end) as [Дата второго взноса],
(CASE WHEN(treaty.term='Ежеквартально')
THEN  DATEADD(MONTH,6, datestart)
else null end
)as [Дата третьего взноса],
(CASE WHEN(treaty.term='Ежеквартально')
THEN   DATEADD(MONTH,9, datestart)
else null end
)as [Дата 4 взноса],
(CASE WHEN(treaty.term='В два срока')
THEN 2-count(pay.idtreaty) else 4-count(pay.idtreaty) end) as t,
treaty.datestart
from treaty inner join pay  on treaty.idtreaty=pay.idtreaty
where treaty.term!='Единовременно' 
and treaty.idtreaty={dataGridView1[0, i].Value.ToString()}
group by treaty.idtreaty,treaty.term,treaty.datestart,datefinish,dateconclusion
having (treaty.term='В два срока' and 2-count(pay.idtreaty)!=0 and treaty.term!='Ежеквартально') 
or (4-count(pay.idtreaty)!=0 and treaty.term='Ежеквартально') ";
                            System.Data.DataTable data2 = new System.Data.DataTable();
                            SqlDataAdapter command2 = new SqlDataAdapter(query2, sqlConnection);
                            command2.Fill(data2);
                            DataColumn column2 = data2.Columns[6];
                            DataRow row2 = data2.Rows[0];
                            hall = Convert.ToInt32(row2[column2].ToString());
                            DataColumn column24 = data2.Columns[4];
                            DataRow row24 = data2.Rows[0];
                            hyll = (row24[column24].ToString());
                            if (hyll == "")
                            {
                                DataColumn column3 = data2.Columns[2];
                                DataRow row3 = data2.Rows[0];
                                min = Convert.ToDateTime(row3[column3].ToString()); ;
                                DataColumn column4 = data2.Columns[3];
                                DataRow row4 = data2.Rows[0];
                                max = Convert.ToDateTime(row4[column4].ToString());
                            }
                            else if (hall == 3)
                            {
                                DataColumn column3 = data2.Columns[2];
                                DataRow row3 = data2.Rows[0];
                                min = Convert.ToDateTime(row3[column3].ToString()); ;
                                DataColumn column4 = data2.Columns[3];
                                DataRow row4 = data2.Rows[0];
                                max = Convert.ToDateTime(row4[column4].ToString());
                            }
                            else if (hall == 2)
                            {
                                DataColumn column3 = data2.Columns[3];
                                DataRow row3 = data2.Rows[0];
                                min = Convert.ToDateTime(row3[column3].ToString());
                                DataColumn column4 = data2.Columns[4];
                                DataRow row4 = data2.Rows[0];
                                max = Convert.ToDateTime(row4[column4].ToString());
                            }
                            else if (hall == 1 && hyll != "")
                            {
                                DataColumn column3 = data2.Columns[4];
                                DataRow row3 = data2.Rows[0];
                                min = Convert.ToDateTime(row3[column3].ToString());
                                DataColumn column4 = data2.Columns[5];
                                DataRow row4 = data2.Rows[0];
                                max = Convert.ToDateTime(row4[column4].ToString());
                            }

                            if (Convert.ToInt32(DateTime.Today.ToString().Substring(3, 2)) + 1 == Convert.ToInt32(max.ToString().Substring(3, 2)))
                            {

                                dataGridView1.Rows[i].Selected = true;
                                dataGridView1.Rows[i].DefaultCellStyle.SelectionForeColor = Color.Black;
                                dataGridView1.Rows[i].DefaultCellStyle.SelectionBackColor = Color.FromArgb(250, 170, 209);
                                // break;
                                if (dataGridView1[23, i].Value.ToString() != "")
                                {
                                    if (hall == 3)
                                    {
                                        wordcellrange = wordtable.Cell(m, 4).Range;
                                        wordcellrange.Text = dataGridView1[19, i].Value.ToString();
                                        wordcellrange = wordtable.Cell(m, 5).Range;
                                        wordcellrange.Text = dataGridView1[18, i].Value.ToString().Substring(0, 10);

                                    }
                                    else if (hall == 2)
                                    {
                                        wordcellrange = wordtable.Cell(m, 4).Range;
                                        wordcellrange.Text = dataGridView1[21, i].Value.ToString();
                                        wordcellrange3 = wordtable.Cell(m, 5).Range;
                                        wordcellrange3.Text = dataGridView1[20, i].Value.ToString().Substring(0, 10);
                                    }
                                    else
                                    {
                                        wordcellrange = wordtable.Cell(m, 4).Range;
                                        wordcellrange.Text = dataGridView1[23, i].Value.ToString();
                                        wordcellrange = wordtable.Cell(m, 5).Range;
                                        wordcellrange.Text = dataGridView1[22, i].Value.ToString().Substring(0, 10);
                                    }
                                }
                                else
                                {
                                    wordcellrange = wordtable.Cell(m, 4).Range;
                                    wordcellrange.Text = dataGridView1[19, i].Value.ToString();
                                    wordcellrange = wordtable.Cell(m, 5).Range;
                                    wordcellrange.Text = dataGridView1[18, i].Value.ToString().Substring(0, 10);
                                }
                                wordcellrange = wordtable.Cell(m, 2).Range;
                                wordcellrange.Text = dataGridView1[2, i].Value.ToString();
                                wordcellrange = wordtable.Cell(m, 1).Range;
                                wordcellrange.Text = dataGridView1[0, i].Value.ToString();
                                wordcellrange = wordtable.Cell(m, 3).Range;
                                wordcellrange.Text = dataGridView1[4, i].Value.ToString();

                            }
                            else { m--; }
                            i++;
                        }
                        catch { i++; m--; }
                    }
                }
                application.Visible = true;
            }
            catch { }
        }

       
        int j = 0;
       
        
    
     
       
        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (MessageBox.Show($@"Задача выполнена?", "Выполнение", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                if ((dataGridView1.CurrentRow.Cells[5].Value.ToString() == "Назначена")) {
                    if (Convert.ToDateTime(dataGridView1.CurrentRow.Cells[4].Value.ToString().Substring(0, 10)) == Convert.ToDateTime(Convert.ToString(DateTime.Now).ToString().Substring(0, 10)))
                    {     
                        sqlConnection.Open();
                        SqlCommand command = new SqlCommand($@"UPDATE working SET date=@d ," +
                        $"status=@s WHERE idworking=@id", sqlConnection);
                        command.Parameters.AddWithValue("@s", ("Выполнено"));
                        command.Parameters.AddWithValue("@d", (dateTimePicker1.Value));
                        command.Parameters.AddWithValue("@id", (id));
                        command.ExecuteNonQuery();
                        sqlConnection.Close();
                        dataGridView1.Enabled = true;
                        clear();
                        if (insurerak != null)
                        {
                            Tread_load2();
                        }
                        {
                            Tread_load();
                        }
                    }
                    else if (Convert.ToDateTime(dataGridView1.CurrentRow.Cells[4].Value.ToString().Substring(0,10)) <Convert.ToDateTime( Convert.ToString(DateTime.Now).ToString().Substring(0,10)))
                    
                        {
                        sqlConnection.Open();
                        SqlCommand command = new SqlCommand($@"UPDATE working SET date=@d ," +
                        $"status=@s WHERE idworking=@id", sqlConnection);
                        command.Parameters.AddWithValue("@s", ("Позже срока"));
                        command.Parameters.AddWithValue("@d", (dateTimePicker1.Value));
                        command.Parameters.AddWithValue("@id", (id));
                        command.ExecuteNonQuery();
                        sqlConnection.Close();
                        dataGridView1.Enabled = true;
                        clear();
                        if (insurerak != null)
                        {
                            Tread_load2();
                        }
                        {
                            Tread_load();
                        }
                    
                }
                else
                {
                    MessageBox.Show("Дата реализации проекта еще не наступила!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
                else
                {
                    MessageBox.Show("Работа уже выполнена!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

        }

        private void Treaty_Layout(object sender, LayoutEventArgs e)
        {
            podsvetka();
        }

        private void Treaty_Validated(object sender, EventArgs e)
        {
            podsvetka();
        }

        private void Treaty_Paint(object sender, PaintEventArgs e)
        {
            podsvetka();
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            sqlConnection.Close();
            sqlConnection.Open();
            string query = $@"select post.name from post, worker where post.idpost=worker.idpost and (worker.firstname +' '+worker.name +' '+worker.lastname) = '{comboBox1.Text}'";
            SqlCommand command = new SqlCommand(query, sqlConnection);
            string result = (string)command.ExecuteScalar();
            sqlConnection.Close();
            if (result == "Уборщик")
            {
                comboBox2.Visible = true;
                comboBox2.Text = "";
                comboBox2.Items.Clear();
                comboBox2.Items.Add("Уборка территории (двор, тротуары, проезды)");
                comboBox2.Items.Add("Мойка полов в помещение");
                comboBox2.Items.Add("Уборка мусора");
                comboBox2.Items.Add("Обслуживать контейнер для мусора");
                ///  one();
                comboBox2.SelectedIndex = -1;

            }
            else if (result == "Электрик")
            {
                comboBox2.Visible = true;
                comboBox2.Text = "";
                comboBox2.Items.Clear();
                comboBox2.Items.Add("Плановые проверки электрических систем");
                comboBox2.Items.Add("Регулярное обслуживание освещения в подъездах и дворах");
                comboBox2.Items.Add("Контроль за состоянием электробезопасности на объектах");
                //  one();
                comboBox2.SelectedIndex = -1;
            }
            else if (result == "Слесарь-сантехник")
            {
                comboBox2.Visible = true;
                comboBox2.Text = "";
                comboBox2.Items.Clear();
                comboBox2.Items.Add("Профилактическое обслуживание сантехнических систем");
                comboBox2.Items.Add("Регулярный осмотр и чистка канализации");
                comboBox2.Items.Add("Проверка состояния и исправности водопроводных труб");

                //   one();
                comboBox2.SelectedIndex = -1;
            }
            else if (result == "Кровельщик")
            {
                comboBox2.Visible = true;
                comboBox2.Text = "";
                comboBox2.Items.Clear();
                comboBox2.Items.Add("Регулярный осмотр кровли для выявления повреждений");
                comboBox2.Items.Add("Устранение потенциальных источников протечек");
                comboBox2.Items.Add("Обслуживание водосточных систем");
                // one();
                comboBox2.SelectedIndex = -1;
            }
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
           
        }

        private void dataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (id != 0)
            {
                try
                {
                    if (MessageBox.Show($@"Вы уверены что хотите удалить задачу?", "Удаление", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        if (dataGridView1.CurrentRow.Cells[5].Value.ToString() != "Выполнено"|| dataGridView1.CurrentRow.Cells[5].Value.ToString() != "Позже срока")
                        {
                            sqlConnection.Open();
                            string query = $@"DELETE FROM [working] WHERE [idworking] = 
                             {dataGridView1.CurrentRow.Cells[0].Value.ToString()}";
                            SqlCommand command = new SqlCommand(query, sqlConnection);
                            command.ExecuteNonQuery();
                            sqlConnection.Close();

                            if (insurerak != null)
                            {
                                Tread_load2();
                            }
                            {
                                Tread_load();
                            }
                        }
                        else
                        {
                            MessageBox.Show("Удаление невозможно, данные все еще нужны!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                catch { MessageBox.Show("Удаление невозможно, данные все еще нужны!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            }
            else { MessageBox.Show("Выберите строку для удаления!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }
    }
}
