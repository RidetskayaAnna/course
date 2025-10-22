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
        Tenantak policyholder;
        public Pays(Tenantak policyholder1)
        {
            InitializeComponent();
            sqlConnection = new SqlConnection(db.connection);
            policyholder = policyholder1;
        }
        Workerak insurerak;
        public Pays(Workerak insurerak1)
        {
            InitializeComponent();
            sqlConnection = new SqlConnection(db.connection);
            insurerak = insurerak1;
        }
        public void Pays_load()
        { if (policyholder != null)
            {
                sqlConnection.Open();
                DataSet dataSet = new DataSet();
                SqlDataAdapter command = new SqlDataAdapter($@"Select idpay,pay.idbid as [Номер заявки] ,datepay as [Дата оплаты],summa as [Сумма] from pay inner join bid on pay.idbid=bid.idbid where bid.idtenant={policyholder.idakk}", sqlConnection);
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
                SqlDataAdapter command = new SqlDataAdapter($@"Select idpay,pay.idbid as [Номер заявки] ,datepay as [Дата оплаты],summa as [Сумма] from pay inner join bid on pay.idbid=bid.idbid where bid.idworker={insurerak.idakk}", sqlConnection);
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
                SqlDataAdapter command = new SqlDataAdapter($@"Select idpay,pay.idbid as [Номер заявки] ,datepay as [Дата оплаты],summa as [Сумма] from pay inner join bid on pay.idbid=bid.idbid", sqlConnection);
                command.Fill(dataSet);
                dataGridView1.DataSource = dataSet.Tables[0].DefaultView;
                sqlConnection.Close();
                dataGridView1.Columns[0].Visible = false;
                dataGridView1.AllowUserToAddRows = false;
                sqlConnection.Close();
            }
        }
       
        public void clear()
        {
            comboBox1.SelectedIndex = -1;
            dateTimePicker1.Value = DateTime.Today;
            textBox2.Text = "";
        }
      
        public void comboBoxtreaty()
        {
            if (policyholder != null)
            {
                sqlConnection.Open();
                string query = $@"select bid.idbid as o from bid left join pay on bid.idbid=pay.idbid where status='Выполнено' and pay.idbid IS NULL and bid.idtenant={policyholder.idakk}";
                SqlDataAdapter sqlDbDataAdapter = new SqlDataAdapter(query, sqlConnection);
                DataSet dataSet = new DataSet();
                sqlDbDataAdapter.Fill(dataSet);
                comboBox1.DataSource = dataSet.Tables[0];
                comboBox1.DisplayMember = "o";
                comboBox1.ValueMember = "o";
                comboBox1.SelectedIndex = -1;
                sqlConnection.Close();
            }
            else
            {
                sqlConnection.Open();
                string query = $@"select bid.idbid as o from bid left join pay on bid.idbid=pay.idbid where status='Выполнено' and pay.idbid IS NULL";
                SqlDataAdapter sqlDbDataAdapter = new SqlDataAdapter(query, sqlConnection);
                DataSet dataSet = new DataSet();
                sqlDbDataAdapter.Fill(dataSet);
                comboBox1.DataSource = dataSet.Tables[0];
                comboBox1.DisplayMember = "o";
                comboBox1.ValueMember = "o";
                comboBox1.SelectedIndex = -1;
                sqlConnection.Close();
            }

        }

        private void Pays_Load(object sender, EventArgs e)
        {

            if (insurerak != null) {  panel2.Visible = false; button6.Visible = false; button10.Visible = false; }
            else
            {
               
                
                //comboBoxtreaty2();
                button6.Visible = true;
                button10.Visible = true;
                panel2.Visible = false;
            }
            Pays_load();comboBoxtreaty();
        }

        private void button10_Click(object sender, EventArgs e)
        {
           
            if (panel2.Visible == false)
            {
                comboBoxtreaty();
                if (comboBox1.Items.Count != 0)
                {

                    panel2.Visible = true;
                    clear();
                    label6.Text = "Добавить оплату";
                    button11.Text = "Добавить";

                    
                }
                else
                {
                    MessageBox.Show("Все заявки уже оплачены!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else
            {
                   

                
                clear();
                panel2.Visible = false;
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
           
            panel2.Visible = false;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            
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
            int y = 0;
            for (int j = 1; j < dataGridView1.ColumnCount; j++)
            {
                y = 0;
                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    if (dataGridView1.Rows[i].Cells[j].Visible == true)
                    {
                        if (j == 2)
                        {
                            ExcelApp.Cells[y + 3, j] = dataGridView1.Rows[i].Cells[j].Value.ToString().Substring(0, 10);
                            y++;
                        }
                        else
                        {
                            ExcelApp.Cells[y + 3, j] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                            y++;
                        }
                    }
                }

            }
            for (int i = 0; i < visible; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    Microsoft.Office.Interop.Excel.Range range = ExcelWorkSheet.Range[$"A1:С{visible + 3}"];
                    ExcelWorkSheet.Range[$"A1:С{visible + 3}"].Cells.WrapText = true;
                    range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders.Weight = Excel.XlBorderWeight.xlMedium;
                }
            }
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }

  

        private void button11_Click(object sender, EventArgs e)
        {
            try
            {
                if (label6.Text == "Добавить оплату")
                {

                    //Фрагмент кода добавления оплаты по договору
                    sqlConnection.Open();
                    SqlCommand command = new SqlCommand($@"INSERT INTO [pay](idbid,datepay,summa)
                    VALUES (@idt,@date,@s);", sqlConnection);
                    command.Parameters.AddWithValue("@idt", (comboBox1.SelectedValue));
                    command.Parameters.AddWithValue("@date", (dateTimePicker1.Value));
                    command.Parameters.AddWithValue("@s", Convert.ToDecimal(textBox2.Text));
                   
                    command.ExecuteNonQuery();
                    sqlConnection.Close();
                        comboBoxtreaty();
                        Pays_load();
                    

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
                
                string query2 = $@"select  cost from bid left join pay on bid.idbid=pay.idbid inner join service on
service.idservice=bid.idservice
where status='Выполнено' and bid.idbid='{comboBox1.Text}' and pay.idbid IS NULL";

                System.Data.DataTable data2 = new System.Data.DataTable();
                SqlDataAdapter command2 = new SqlDataAdapter(query2, sqlConnection);
                command2.Fill(data2);
                DataColumn column2 = data2.Columns[0];
                DataRow row2 = data2.Rows[0];
                textBox2.Text = (row2[column2].ToString());
                
               
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

       

        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {


        }
    }

}
