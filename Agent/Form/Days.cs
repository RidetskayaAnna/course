using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Agent.Form
{
    public partial class Days : UserControl
    {
        datebase db = new datebase();
        public SqlConnection sqlConnection = null;
        
        public Days()
        {
            InitializeComponent();
            sqlConnection = new SqlConnection(db.connection);
            ;
        }

        private void Days_Load(object sender, EventArgs e)
        {
           
        }
        public void days(int numday)
        {
            label6.Text = numday+"";
            
        }

        public void task(int numday, int year, int month)
        {
            Calendar ucdays = new Calendar();
            DateTime date = new DateTime(year, month, numday);
            string stroca = "";
            int s = 0; int f = 0;
            string p = $@"SELECT COUNT(*) AS TotalRows
            FROM(
                SELECT(working.name + '- ' + object.type + ', ' + object.address) AS FullInfo
                FROM worker
                INNER JOIN working ON worker.idworker = working.idworker        
                INNER JOIN object ON object.idobject = working.idobject        
                WHERE working.status = 'Назначена' and working.date='{date.ToString("yyyy-MM-dd")}'
                UNION
                SELECT(service.name + '- ' + object.type + ', ' + object.address) AS FullInfo
                FROM bid
                INNER JOIN object ON bid.idobject = object.idobject
                INNER JOIN tenant ON bid.idtenant = tenant.idtenant
                INNER JOIN worker ON bid.idworker = worker.idworker
                INNER JOIN service ON bid.idservice = service.idservice
                WHERE bid.status = 'Принят' and bid.datec='{date.ToString("yyyy-MM-dd")}'
            ) AS CombinedResults; ";
            DataTable data2 = new DataTable();
            SqlDataAdapter sqlData2 = new SqlDataAdapter(p, sqlConnection);
            sqlData2.Fill(data2);
            DataColumn column2 = data2.Columns[0];
            DataRow row2 = data2.Rows[0];
            s = Convert.ToInt32(row2[column2].ToString());



            richTextBox1.Clear(); // Очистка перед добавлением текста

            string p1 = $@"Select (working.name +'- '+type+', '+address) from worker, 
working, object where object.idobject=working.idobject and worker.idworker=working.idworker and
working.status='Назначена' and working.date='{date.ToString("yyyy-MM-dd")}'
union
Select (service.name+'- '+type+', '+address) 
from bid inner join object on bid.idobject=object.idobject inner join tenant on bid.idtenant=tenant.idtenant 
inner join worker on bid.idworker=worker.idworker inner join service on bid.idservice=service.idservice
where bid.status='Принят' and bid.datec='{date.ToString("yyyy-MM-dd")}'";
            DataTable data1 = new DataTable();
            SqlDataAdapter sqlData1 = new SqlDataAdapter(p1, sqlConnection);
            sqlData1.Fill(data1);
            DataColumn column1 = data1.Columns[0];
            for (int i = 0; i < s; i++)
            {
                DataRow row1 = data1.Rows[i];
                stroca = (row1[column1].ToString());
                // Добавление элементов
                // Добавление многострочного текста с изменением цвета
                richTextBox1.SelectionColor = Color.DarkRed; // Установка цвета текста
                richTextBox1.AppendText("\u2716" + stroca + "\n");
            }


            string p4 = $@"SELECT COUNT(*) AS TotalRows
            FROM(
                SELECT(working.name + '- ' + object.type + ', ' + object.address) AS FullInfo
                FROM worker
                INNER JOIN working ON worker.idworker = working.idworker
            
                INNER JOIN object ON object.idobject = working.idobject
            
                WHERE working.status = 'Выполнено'    and working.date='{date.ToString("yyyy-MM-dd")}'        
                UNION
           
                SELECT(service.name + '- ' + object.type + ', ' + object.address) AS FullInfo


                FROM bid
            
                INNER JOIN object ON bid.idobject = object.idobject
            
                INNER JOIN tenant ON bid.idtenant = tenant.idtenant
            
                INNER JOIN worker ON bid.idworker = worker.idworker
            
                INNER JOIN service ON bid.idservice = service.idservice
            
                WHERE bid.status = 'Выполнено' and bid.datec='{date.ToString("yyyy-MM-dd")}'
            ) AS CombinedResults; ";
            DataTable data4 = new DataTable();
            SqlDataAdapter sqlData4 = new SqlDataAdapter(p4, sqlConnection);
            sqlData4.Fill(data4);
            DataColumn column4 = data4.Columns[0];
            DataRow row4 = data4.Rows[0];
            f = Convert.ToInt32(row4[column4].ToString());

            string p3 = $@"Select (working.name +'- '+type+', '+address) from worker, 
working, object where object.idobject=working.idobject and worker.idworker=working.idworker and
working.status='Выполнено'   and working.date='{date.ToString("yyyy-MM-dd")}' 
union
Select (service.name+'- '+type+', '+address) 
from bid inner join object on bid.idobject=object.idobject inner join tenant on bid.idtenant=tenant.idtenant 
inner join worker on bid.idworker=worker.idworker inner join service on bid.idservice=service.idservice
where bid.status='Выполнено' and bid.datec='{date.ToString("yyyy-MM-dd")}'";
            DataTable data3 = new DataTable();
            SqlDataAdapter sqlData3 = new SqlDataAdapter(p3, sqlConnection);
            sqlData3.Fill(data3);
            DataColumn column3 = data3.Columns[0];
            for (int i = 0; i < f; i++)
            {
                DataRow row3 = data3.Rows[i];
                stroca = (row3[column3].ToString());
                // Добавление элементов
                // Добавление многострочного текста с изменением цвета
                richTextBox1.SelectionColor = Color.DarkGreen; // Установка цвета текста
                richTextBox1.AppendText("\u2714" + stroca + "\n");

            }
        }
    }
}
