using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;
using Excel = Microsoft.Office.Interop.Excel;

namespace Agent.Form
{
    public partial class Calendar : UserControl
    {
        datebase db = new datebase();
        public SqlConnection sqlConnection = null;
        public Calendar()
        {
            InitializeComponent();
            sqlConnection = new SqlConnection(db.connection);
        }
        int id = 0;
       
       public int month, year;
        private void City_Load(object sender, EventArgs e)
        {
            Display();
           
        }
        string stroca = "";
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (textBox1.Text!="")
            {
                richTextBox1.BringToFront();
                richTextBox1.Visible = true;
                richTextBox1.Clear();
                int s = 0; int f = 0;
                string p = $@"SELECT COUNT(*) AS TotalRows
            FROM(
                SELECT(working.name + '- ' + object.type + ', ' + object.address+','+ CONVERT(NVARCHAR(10), date, 105)) AS FullInfo
                FROM worker
                INNER JOIN working ON worker.idworker = working.idworker        
                INNER JOIN object ON object.idobject = working.idobject        
               where  (working.status = 'Назначена' OR working.status = 'Выполнено' OR working.status = 'Позже срока') 
    AND (CONVERT(NVARCHAR(10), date, 105) + ', ' + working.name + '- ' + type + ', ' + address) LIKE '%{textBox1.Text}%'

                UNION
                SELECT(service.name + '- ' + object.type + ', ' + object.address+','+ CONVERT(NVARCHAR(10), datec, 105)) AS FullInfo
                FROM bid
                INNER JOIN object ON bid.idobject = object.idobject
                INNER JOIN tenant ON bid.idtenant = tenant.idtenant
                INNER JOIN worker ON bid.idworker = worker.idworker
                INNER JOIN service ON bid.idservice = service.idservice
               where  (bid.status = 'Принят' OR bid.status = 'Выполнено' OR bid.status = 'Позже срока') 
    AND (CONVERT(NVARCHAR(10), datec, 105) + ', ' + service.name + '- ' + type + ', ' + address) LIKE '%{textBox1.Text}%'
            ) AS CombinedResults; ";
                System.Data.DataTable data2 = new System.Data.DataTable();
                SqlDataAdapter sqlData2 = new SqlDataAdapter(p, sqlConnection);
                sqlData2.Fill(data2);
                DataColumn column2 = data2.Columns[0];
                DataRow row2 = data2.Rows[0];
                s = Convert.ToInt32(row2[column2].ToString());

                string p1 = $@"SELECT 
    (CONVERT(NVARCHAR(10), date, 105) + ', ' + working.name + '- ' + type + ', ' + address) AS a, 
    status  
FROM 
    worker
INNER JOIN 
    working ON worker.idworker = working.idworker
INNER JOIN 
    object ON object.idobject = working.idobject 
WHERE 
    (working.status = 'Назначена' OR working.status = 'Выполнено' OR working.status = 'Позже срока') 
    AND (CONVERT(NVARCHAR(10), date, 105) + ', ' + working.name + '- ' + type + ', ' + address) LIKE '%{textBox1.Text}%'

UNION

SELECT 
    (CONVERT(NVARCHAR(10), datec, 105) + ', ' + service.name + '- ' + type + ', ' + address) AS a, 
    status  
FROM 
    bid 
INNER JOIN 
    object ON bid.idobject = object.idobject 
INNER JOIN 
    tenant ON bid.idtenant = tenant.idtenant 
INNER JOIN 
    worker ON bid.idworker = worker.idworker 
INNER JOIN 
    service ON bid.idservice = service.idservice
WHERE 
    (bid.status = 'Принят' OR bid.status = 'Выполнено' OR bid.status = 'Позже срока') 
    AND (CONVERT(NVARCHAR(10), datec, 105) + ', ' + service.name + '- ' + type + ', ' + address) LIKE '%{textBox1.Text}%';";
                System.Data.DataTable data1 = new System.Data.DataTable();
                SqlDataAdapter sqlData1 = new SqlDataAdapter(p1, sqlConnection);
                sqlData1.Fill(data1);
                DataColumn column1 = data1.Columns[0];
                for (int i = 0; i < s; i++)
                {
                    DataColumn column7 = data1.Columns[1];
                    DataRow row1 = data1.Rows[i];
                    stroca = (row1[column1].ToString());
                    // Добавление элементов
                    // Добавление многострочного текста с изменением цвета
                    //  richTextBox1.SelectionColor = Color.DarkRed; // Установка цвета текста

                    if ((row1[column7].ToString()) == "Принят" || (row1[column7].ToString()) == "Назначена")
                    { richTextBox1.SelectionColor = Color.DarkRed;
                        richTextBox1.AppendText("\u2716" + stroca + "\n");
                       
                    }
                    else if ((row1[column7].ToString()) == "Позже срока") { 
                        richTextBox1.SelectionColor = Color.DarkOrange;
                        richTextBox1.AppendText("\u2714" + stroca + "\n");
                      
                    }
                    else
                    {
                        richTextBox1.SelectionColor = Color.DarkGreen;
                        richTextBox1.AppendText("\u2714" + stroca + "\n");
                    }
                }
               
                
            }
            else if (textBox1.Text == "")
            {
                richTextBox1.Visible = false;
            }
           
        }

        private void button11_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }
        private void Display()
        {
            DateTime now = DateTime.Now;
            month=now.Month;
            year=now.Year;

            String monthame = DateTimeFormatInfo.CurrentInfo.GetMonthName(month);
            label9.Text = monthame+" "+ year;
            DateTime startofthemonth = new DateTime(year, month, 1);
            int days = DateTime.DaysInMonth(year, month);
            int dayoftheweek = Convert.ToInt32(startofthemonth.DayOfWeek.ToString("d"));

            for(int i=1; i < dayoftheweek; i++)
            {
                Blank blank = new Blank();
                flowLayoutPanel1.Controls.Add(blank);
            }

            for(int i=1; i<= days; i++)
            {
                Days ucdays = new Days();
                ucdays.days(i);
                ucdays.task(i,year,month);
                flowLayoutPanel1.Controls.Add(ucdays);
                
            }
           
        }

        private void button11_Click_1(object sender, EventArgs e)
        {

            textBox1.Text = "";
            richTextBox1.Visible = false;
            flowLayoutPanel1.Controls.Clear();

            if (month == 1)
            {
                month=12;year--;
            }
            else
            {
                month--;
            }
            String monthame = DateTimeFormatInfo.CurrentInfo.GetMonthName(month);
            label9.Text = monthame + " " + year;
            DateTime startofthemonth = new DateTime(year, month, 1);
            int days = DateTime.DaysInMonth(year, month);
            int dayoftheweek = Convert.ToInt32(startofthemonth.DayOfWeek.ToString("d"));

            for (int i = 1; i < dayoftheweek; i++)
            {
                Blank blank = new Blank();
                flowLayoutPanel1.Controls.Add(blank);
            }

            for (int i = 1; i <= days; i++)
            {

                
               Days ucdays = new Days();
                ucdays.days(i);
                ucdays.task(i, year, month);
                flowLayoutPanel1.Controls.Add(ucdays);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
                richTextBox1.Visible = false;
            
            flowLayoutPanel1.Controls.Clear();

            if (month == 12)
            {
                month = 1;year++;
            }
            else
            {
                month++;
            }
        
            String monthame = DateTimeFormatInfo.CurrentInfo.GetMonthName(month);
            label9.Text = monthame + " " + year;
            DateTime startofthemonth = new DateTime(year, month, 1);
            int days = DateTime.DaysInMonth(year, month);
            int dayoftheweek = Convert.ToInt32(startofthemonth.DayOfWeek.ToString("d"));

            for (int i = 1; i < dayoftheweek; i++)
            {
                Blank blank = new Blank();
                flowLayoutPanel1.Controls.Add(blank);
            }

            for (int i = 1; i <= days; i++)
            {
                Days ucdays = new Days();
                ucdays.days(i);
                ucdays.task(i, year, month);
                flowLayoutPanel1.Controls.Add(ucdays);
            }
        }
    }
}
