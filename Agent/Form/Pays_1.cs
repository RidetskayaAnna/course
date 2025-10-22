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
    public partial class Pays : UserControl
    {
        datebase db = new datebase();
        public SqlConnection sqlConnection = null;
        Form1 form1;
        public Pays(Form1 form)
        {
            InitializeComponent();
            sqlConnection = new SqlConnection(db.connection);
            form1 = form;
        }
        Policyholderak policyholder;
        public Pays(Policyholderak policyholder1)
        {
            InitializeComponent();
            sqlConnection = new SqlConnection(db.connection);
            policyholder = policyholder1;
        }

        public void Pays_load()
        {
            sqlConnection.Open();
            DataSet dataSet = new DataSet();
            SqlDataAdapter command = new SqlDataAdapter($@"Select idpay,pay.idtreaty as [Номер договора] ,datepay as [Дата оплаты],summa as [Сумма взноса],vidpay as [Вид оплаты] from pay inner join treaty on pay.idtreaty=treaty.idtreaty", sqlConnection);
            command.Fill(dataSet);
            dataGridView1.DataSource = dataSet.Tables[0].DefaultView;
            sqlConnection.Close();
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.AllowUserToAddRows = false;
            sqlConnection.Close();
        }
        public void Pays_load2()
        {
            sqlConnection.Open();
            DataSet dataSet = new DataSet();
            SqlDataAdapter command = new SqlDataAdapter($@"Select idpay,pay.idtreaty as [Номер договора] ,datepay as [Дата оплаты],summa as [Сумма взноса],vidpay as [Вид оплаты] from pay inner join treaty on pay.idtreaty=treaty.idtreaty where treaty.idpolicyholder={policyholder.idakk}", sqlConnection);
            command.Fill(dataSet);
            dataGridView1.DataSource = dataSet.Tables[0].DefaultView;
            sqlConnection.Close();
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.AllowUserToAddRows = false;
            sqlConnection.Close();
        }
        public void clear()
        {
            comboBox1.SelectedIndex = -1;
            comboBox5.SelectedIndex = 1;
            dateTimePicker1.MinDate = Convert.ToDateTime("01.01.1753");
            dateTimePicker1.MaxDate = Convert.ToDateTime("31.12.9998");
            dateTimePicker1.Value = DateTime.Today;
            textBox2.Text = "";
        }
        public void comboBoxtreaty2()
        {
            sqlConnection.Open();
            string query = $@"select idtreaty from treaty";
            SqlDataAdapter sqlDbDataAdapter = new SqlDataAdapter(query, sqlConnection);
            DataSet dataSet = new DataSet();
            sqlDbDataAdapter.Fill(dataSet);
            comboBox2.DataSource = dataSet.Tables[0];
            comboBox2.DisplayMember = "idtreaty";
            comboBox2.ValueMember = "idtreaty";
            comboBox2.SelectedIndex = -1;
            sqlConnection.Close();

        }
        public void comboBoxtreaty()
        {
            sqlConnection.Open();
            string query = $@"select count(pay.idtreaty),treaty.term,treaty.idtreaty  as o,
(CASE WHEN(treaty.term='В два срока')
THEN 2-count(pay.idtreaty) else 4-count(pay.idtreaty) end) as t
from treaty inner join pay on treaty.idtreaty=pay.idtreaty
where treaty.term!='Единовременно'
group by pay.idtreaty,treaty.term,treaty.idtreaty
having (treaty.term='В два срока' and 2-count(pay.idtreaty)!=0 and treaty.term!='Ежеквартально') 
or
(4-count(pay.idtreaty)!=0 and treaty.term='Ежеквартально')";
            SqlDataAdapter sqlDbDataAdapter = new SqlDataAdapter(query, sqlConnection);
            DataSet dataSet = new DataSet();
            sqlDbDataAdapter.Fill(dataSet);
            comboBox1.DataSource = dataSet.Tables[0];
            comboBox1.DisplayMember = "o";
            comboBox1.ValueMember = "o";
            comboBox1.SelectedIndex = -1;
            sqlConnection.Close();

        }
        public void comboBoxtreaty3()
        {
            sqlConnection.Open();
            string query = $@"select count(pay.idtreaty),treaty.term,treaty.idtreaty  as o,
(CASE WHEN(treaty.term='В два срока')
THEN 2-count(pay.idtreaty) else 4-count(pay.idtreaty) end) as t
from treaty inner join pay on treaty.idtreaty=pay.idtreaty
where treaty.term!='Единовременно' and treaty.idpolicyholder={policyholder.idakk}
group by pay.idtreaty,treaty.term,treaty.idtreaty
having (treaty.term='В два срока' and 2-count(pay.idtreaty)!=0 and treaty.term!='Ежеквартально') 
or
(4-count(pay.idtreaty)!=0 and treaty.term='Ежеквартально')";
            SqlDataAdapter sqlDbDataAdapter = new SqlDataAdapter(query, sqlConnection);
            DataSet dataSet = new DataSet();
            sqlDbDataAdapter.Fill(dataSet);
            comboBox1.DataSource = dataSet.Tables[0];
            comboBox1.DisplayMember = "o";
            comboBox1.ValueMember = "o";
            comboBox1.SelectedIndex = -1;
            sqlConnection.Close();

        }
        private void Pays_Load(object sender, EventArgs e)
        {

            if (policyholder != null)
            {
                try
                {
                    Pays_load2();
                    comboBoxtreaty3();
                    button6.Visible = false;
                    button3.Visible = false;
                    panel3.Visible = false;
                    panel2.Visible = false;
                }
                catch { }
            }
            else
            {
                Pays_load();
                comboBoxtreaty();
                comboBoxtreaty2();
                button6.Visible = true;
                button3.Visible = true;
                panel3.Visible = false;
                panel2.Visible = false;
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            panel3.Visible = false;
            button3.Image = new Bitmap(@"D:\College\4kurs\Praktica4Kurs\proga\Agent\Agent\Resources\pngwing.com (48).png");
            if (panel2.Visible == false)
            {
                if (comboBox1.Items.Count != 0)
            {

                panel2.Visible = true;
                clear();
                label6.Text = "Добавить оплату";
                button11.Text = "Добавить";
                button11.Width = 190;
                button11.Left = 155;
                    if (policyholder != null)
                    {
                        try
                        {
                            dataGridView1.Width = 803;
                            Pays_load2();
                            panel2.Left = 820;
                            panel2.Top = 104;
                        }
                        catch { }
                    }
                    else
                    {
                        dataGridView1.Width = 1311;
                        panel2.Left = 422;
                        panel2.Top = 486;
                        Pays_load();
                    }
                if (Convert.ToInt32(comboBox1.SelectedIndex) != -1)
                {
                    dateTimePicker1.Enabled = true;
                }
                else
                {
                    dateTimePicker1.Enabled = false;
                }
                dataGridView1.Enabled = true;
            }
            else
            {
                MessageBox.Show("Все договора уже оплачены!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
            else
            {
                if (policyholder != null)
                {
                    try
                    {
                        dataGridView1.Width = 1311;

                    }
                    catch { }
                }
                else
                {

                }
                clear();
        panel2.Visible = false;
            }
}

        private void button9_Click(object sender, EventArgs e)
        {
            panel3.Visible = false;
            panel2.Visible = false;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            panel3.Visible = false;
            panel2.Visible = false;
        }
        int visible = 0;
        private void button6_Click(object sender, EventArgs e)
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
            ExcelApp.Rows[1].Columns[2] = "Оплата";
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
                        if (j == 2)
                        {
                            ExcelApp.Cells[i + 3, j] = dataGridView1.Rows[i].Cells[j].Value.ToString().Substring(0, 10);
                        }
                        else
                        {
                             ExcelApp.Cells[i + 3, j] = dataGridView1.Rows[i].Cells[j].Value.ToString();

                        }
                    }
                }

            }
            for (int i = 0; i < visible; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    Microsoft.Office.Interop.Excel.Range range = ExcelWorkSheet.Range[$"A1:D{visible + 3}"];
                    ExcelWorkSheet.Range[$"A1:D{visible + 3}"].Cells.WrapText = true;
                    range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders.Weight = Excel.XlBorderWeight.xlMedium;
                }
            }
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            clear();            
            panel2.Visible = false;
            dataGridView1.Enabled = true;
            if (panel3.Visible == true)
            {
                panel3.Visible = false;
                Pays_load();
                button3.Image = new Bitmap(@"D:\College\4kurs\Praktica4Kurs\proga\Agent\Agent\Resources\pngwing.com (48).png");
            }
            else
            {
                panel3.Visible = true;
                comboBox2.SelectedItem = -1;
                button3.Image = new Bitmap(@"D:\College\4kurs\Praktica4Kurs\proga\Agent\Agent\Resources\pngwing.com (47).png");
            }

        }

        private void button11_Click(object sender, EventArgs e)
        {
            try
            {
                if (label6.Text == "Добавить оплату")
            {
                sqlConnection.Open();
                SqlCommand command = new SqlCommand($@"INSERT INTO [pay](idtreaty,datepay,summa,vidpay) VALUES (@idt,@date,@s,@v);", sqlConnection);
                command.Parameters.AddWithValue("@idt", (comboBox1.SelectedValue));
                command.Parameters.AddWithValue("@date", (dateTimePicker1.Value));
                command.Parameters.AddWithValue("@s", Convert.ToDecimal(textBox2.Text));
                command.Parameters.AddWithValue("@v", (comboBox5.SelectedItem));
                command.ExecuteNonQuery();
                sqlConnection.Close();

                    if (policyholder != null)
                    {
                        try
                        {
                            dataGridView1.Width = 1311;
                            Pays_load2();
                            comboBoxtreaty3();

                        }
                        catch { }
                    }
                    else
                    {
                       comboBoxtreaty();
                Pays_load(); 
                    }
                    
                panel2.Visible = false;
            }
        }
            catch { }
        }
        Decimal kk = 0;
        int hall = 0;
        string hyll = "";
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (Convert.ToInt32(comboBox1.SelectedIndex)!=-1)
                {
                    dateTimePicker1.Enabled = true;
                }
                else
                {
                    dateTimePicker1.Enabled = false;
                }
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
and treaty.idtreaty={comboBox1.SelectedValue}
group by treaty.idtreaty,treaty.term,treaty.datestart,datefinish,dateconclusion
having (treaty.term='В два срока' and 2-count(pay.idtreaty)!=0 and treaty.term!='Ежеквартально') 
or (4-count(pay.idtreaty)!=0 and treaty.term='Ежеквартально') ";
            
                System.Data.DataTable data2 = new System.Data.DataTable();
                SqlDataAdapter command2 = new SqlDataAdapter(query2, sqlConnection);
                command2.Fill(data2);
                DataColumn column2 = data2.Columns[6];
                DataRow row2 = data2.Rows[0];
                hall =Convert.ToInt32 (row2[column2].ToString());
                DataColumn column24 = data2.Columns[4];
                DataRow row24 = data2.Rows[0];
                hyll = (row24[column24].ToString());
                if (hyll=="")
                {
                    DataColumn column3 = data2.Columns[2];
                    DataRow row3 = data2.Rows[0];
                    dateTimePicker1.MinDate = Convert.ToDateTime(row3[column3].ToString()); ;
                    DataColumn column4 = data2.Columns[3];
                    DataRow row4 = data2.Rows[0];
                    dateTimePicker1.MaxDate = Convert.ToDateTime(row4[column4].ToString());
                }
                else if(hall==3)
                {
                    DataColumn column3 = data2.Columns[2];
                    DataRow row3 = data2.Rows[0];
                    dateTimePicker1.MinDate = Convert.ToDateTime(row3[column3].ToString()); ;
                    DataColumn column4 = data2.Columns[3];
                    DataRow row4 = data2.Rows[0];
                    dateTimePicker1.MaxDate = Convert.ToDateTime(row4[column4].ToString());
                }
                else if (hall == 2)
                {
                    DataColumn column3 = data2.Columns[3];
                    DataRow row3 = data2.Rows[0];
                    dateTimePicker1.MinDate = Convert.ToDateTime(row3[column3].ToString());
                    DataColumn column4 = data2.Columns[4];
                    DataRow row4 = data2.Rows[0];
                    dateTimePicker1.MaxDate = Convert.ToDateTime(row4[column4].ToString());
                }
                else if (hall == 1&& hyll != "")
                {
                    DataColumn column3 = data2.Columns[4];
                    DataRow row3 = data2.Rows[0];
                    dateTimePicker1.MinDate = Convert.ToDateTime(row3[column3].ToString());
                    DataColumn column4 = data2.Columns[5];
                    DataRow row4 = data2.Rows[0];
                    dateTimePicker1.MaxDate = Convert.ToDateTime(row4[column4].ToString());
                }

                string query3 = $@"Select vznos from treaty where idtreaty={comboBox1.SelectedValue}";
                System.Data.DataTable data3 = new System.Data.DataTable();
                SqlDataAdapter command3 = new SqlDataAdapter(query3, sqlConnection);
                command3.Fill(data3);
                DataColumn column23 = data3.Columns[0];
                DataRow row23 = data3.Rows[0];
                kk = Convert.ToDecimal(row23[column23].ToString());
                textBox2.Text = Convert.ToString(kk);
        }
            catch { }
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

        private void button12_Click(object sender, EventArgs e)
        {

                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    dataGridView1.CurrentCell = null;
                    dataGridView1.Rows[i].Visible = false;

                    if (Convert.ToInt32( dataGridView1[1, i].Value.ToString())== Convert.ToInt32(comboBox2.SelectedValue))
                    {
                        dataGridView1.Rows[i].Visible = true;
                    }
                    else
                    {
                        dataGridView1.Rows[i].Visible = false;
                    }
                }

            
        }

        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            try
            {
                string query2 = $@" Select pay.idtreaty as [Номер договора] ,
vidpay as [Вид оплаты] from pay inner join treaty on pay.idtreaty=treaty.idtreaty  where pay.idtreaty={comboBox1.SelectedValue}
 ";
                System.Data.DataTable data2 = new System.Data.DataTable();
                SqlDataAdapter command2 = new SqlDataAdapter(query2, sqlConnection);
                command2.Fill(data2);
                DataColumn column2 = data2.Columns[1];
                DataRow row2 = data2.Rows[0];
                comboBox5.SelectedItem = (row2[column2].ToString());
            }
            catch { }
            
        }
    }
    
}
