using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Agent.Form
{
    public partial class Static : UserControl
    {
        datebase db = new datebase();
        public SqlConnection sqlConnection = null;
        public Static()
        {
            InitializeComponent();
            sqlConnection = new SqlConnection(db.connection);
        }string so = "";
        string st = "";
        public void Static_loads()
        {
            so = dateTimePicker1.Value.ToString("yyyy-MM-dd");
            st = dateTimePicker2.Value.ToString("yyyy-MM-dd");
            chart1.Series[0].Points.Clear();
            string query2 = $@"Select count( bid.idworker)
from bid
inner join worker on bid.idworker = worker.idworker
where datec>='{so}' and datec<='{st}' and (status='Выполнено' or status='Позже срока')
group by (worker.firstname + ' ' + worker.name + ' ' + worker.lastname)";
            DataTable data2 = new DataTable();
            SqlDataAdapter command2 = new SqlDataAdapter(query2, sqlConnection);
            command2.Fill(data2);
            DataColumn column2 = data2.Columns[0];
            DataRow row2 = data2.Rows[0];
            int y =Convert.ToInt32( row2[column2].ToString());

            
            sqlConnection.Open();
            string query = $@"Select count( bid.idworker) as d,(worker.firstname + ' ' + worker.name + ' ' + worker.lastname) as y
from bid
inner join worker on bid.idworker = worker.idworker
where datec>='{so}' and datec<='{st}' and (status='Выполнено' or status='Позже срока')
group by (worker.firstname + ' ' + worker.name + ' ' + worker.lastname)";
            SqlDataAdapter command = new SqlDataAdapter(query, sqlConnection);
            DataTable dataSet = new DataTable();
            command.Fill(dataSet);
            chart1.DataSource = dataSet;
            chart1.Series[0].XValueMember = "y";
            chart1.Legends[0].Name = "d";
            chart1.Series[0].YValueMembers = "d";
            chart1.DataBind();
            sqlConnection.Close();

        }
        
        public void Static_loadsс()
        { so = dateTimePicker1.Value.ToString("yyyy-MM-dd");
            st = dateTimePicker2.Value.ToString("yyyy-MM-dd");
            chart2.Series[0].Points.Clear();
            string query2 = $@"Select count( working.idworker)
from working inner join worker on working.idworker=worker.idworker 
where date>='{so}' and date<='{st}'  and (status='Выполнено' or status='Позже срока')
group by (worker.firstname + ' ' + worker.name + ' ' + worker.lastname)
";
            DataTable data2 = new DataTable();
            SqlDataAdapter command2 = new SqlDataAdapter(query2, sqlConnection);
            command2.Fill(data2);
            DataColumn column2 = data2.Columns[0];
            DataRow row2 = data2.Rows[0];
            int y = Convert.ToInt32(row2[column2].ToString());

            sqlConnection.Open();
            string query = $@"Select count( working.idworker),(worker.firstname + ' ' + worker.name + ' ' + worker.lastname)  from working inner join worker on working.idworker=worker.idworker
where date>='{so}' and date<='{st}'  and (status='Выполнено' or status='Позже срока')
group by (worker.firstname + ' ' + worker.name + ' ' + worker.lastname)";
            SqlDataAdapter command = new SqlDataAdapter(query, sqlConnection);
            DataTable dataSet = new DataTable();
            command.Fill(dataSet);
            chart2.DataSource = dataSet;
            sqlConnection.Close();
            try
            {
                for (int j = 0; j < y; j++)
                {
                    DataColumn column = dataSet.Columns[0];
                    DataRow row = dataSet.Rows[j];
                    DataColumn column3 = dataSet.Columns[1];
                    DataRow row3 = dataSet.Rows[j];
                    chart2.Series[0].Points.AddXY(row3[column3].ToString(), row[column].ToString());
                }
            }
            catch
            {

            }
        }

        public void Static_loadscс()
        {
            so = dateTimePicker1.Value.ToString("yyyy-MM-dd");
            st = dateTimePicker2.Value.ToString("yyyy-MM-dd");
            chart3.Series[0].Points.Clear();
            string query2 = $@"select count( bid.idbid)
from bid
inner join worker on bid.idworker = worker.idworker
where datec>='{so}' and datec<='{st}' and (status='Выполнено' or status='Позже срока')
";
            DataTable data2 = new DataTable();
            SqlDataAdapter command2 = new SqlDataAdapter(query2, sqlConnection);
            command2.Fill(data2);
            DataColumn column2 = data2.Columns[0];
            DataRow row2 = data2.Rows[0];
            int y = Convert.ToInt32(row2[column2].ToString());

            sqlConnection.Open();
            string query = $@" SELECT 
    
    COUNT(bid.idbid) AS [Количество заявок],CONCAT(DATEPART(MONTH, datec), '-',DATEPART(YEAR, datec) ) AS [Год-Месяц]
FROM 
    bid
INNER JOIN 
    worker ON bid.idworker = worker.idworker
WHERE 
    (status = 'Выполнено' OR status = 'Позже срока') and datec>='{so}' and datec<='{st}'
GROUP BY 
    DATEPART(YEAR, datec), 
    DATEPART(MONTH, datec)
ORDER BY 
    DATEPART(YEAR, datec), 
    DATEPART(MONTH, datec)";
            SqlDataAdapter command = new SqlDataAdapter(query, sqlConnection);
            DataTable dataSet = new DataTable();
            command.Fill(dataSet);

            string query12 = $@"select Distinct  CONCAT(DATEPART(MONTH, datec), '-',DATEPART(YEAR, datec) )
from bid
inner join worker on bid.idworker = worker.idworker
where datec>='{so}' and datec<='{st}' and (status='Выполнено' or status='Позже срока')";
            SqlDataAdapter command12 = new SqlDataAdapter(query12, sqlConnection);
            DataTable dataSet12 = new DataTable();
            command12.Fill(dataSet12);
            chart3.DataSource = dataSet12;
            sqlConnection.Close();
            try
            {
                for (int j = 0; j < y; j++)
                {
                    DataColumn column = dataSet.Columns[0];
                    DataRow row = dataSet.Rows[j];
                    DataColumn column3 = dataSet12.Columns[0];
                    DataRow row3 = dataSet12.Rows[j];
                    chart3.Series[0].Points.AddXY(row3[column3].ToString(), row[column].ToString());
                }
            }
            catch
            {

            }

        }
        public void Static_loadscсс()
        {
            so = dateTimePicker1.Value.ToString("yyyy-MM-dd");
            st = dateTimePicker2.Value.ToString("yyyy-MM-dd");
            chart4.Series[0].Points.Clear();
            string query2 = $@"select count(bid.idbid)
from bid
inner join object on bid.idobject=object.idobject 
where datec>='{so}' and datec<='{st}' and (status='Выполнено' or status='Позже срока')";
            DataTable data2 = new DataTable();
            SqlDataAdapter command2 = new SqlDataAdapter(query2, sqlConnection);
            command2.Fill(data2);
            DataColumn column2 = data2.Columns[0];
            DataRow row2 = data2.Rows[0];
            int y = Convert.ToInt32(row2[column2].ToString());

            sqlConnection.Open();
            string query = $@"select count(bid.idbid),type
from bid
inner join object on bid.idobject=object.idobject 
where datec>='{so}' and datec<='{st}' and (status='Выполнено' or status='Позже срока')
group by type";
            SqlDataAdapter command = new SqlDataAdapter(query, sqlConnection);
            DataTable dataSet = new DataTable();
            command.Fill(dataSet);
            chart3.DataSource = dataSet;
            sqlConnection.Close();
            try
            {
                for (int j = 0; j < y; j++)
                {
                    DataColumn column = dataSet.Columns[0];
                    DataRow row = dataSet.Rows[j];
                    DataColumn column3 = dataSet.Columns[1];
                    DataRow row3 = dataSet.Rows[j];
                    chart4.Series[0].Points.AddXY(row3[column3].ToString(), row[column].ToString());
                }
            }
            catch
            {

            }
        }

        //
        public void Static_loads2()
        {
            chart1.Series[0].Points.Clear();
            string query2 = $@"Select count( bid.idworker)
from bid
inner join worker on bid.idworker = worker.idworker
where  (status='Выполнено' or status='Позже срока')
group by (worker.firstname + ' ' + worker.name + ' ' + worker.lastname) ";
            DataTable data2 = new DataTable();
            SqlDataAdapter command2 = new SqlDataAdapter(query2, sqlConnection);
            command2.Fill(data2);
            DataColumn column2 = data2.Columns[0];
            DataRow row2 = data2.Rows[0];
            int y = Convert.ToInt32(row2[column2].ToString());

            sqlConnection.Open();
            string query = $@"Select count( bid.idworker) as d,(worker.firstname + ' ' + worker.name + ' ' + worker.lastname) as y
from bid
inner join worker on bid.idworker = worker.idworker
where (status='Выполнено' or status='Позже срока')
group by (worker.firstname + ' ' + worker.name + ' ' + worker.lastname)";
            SqlDataAdapter command = new SqlDataAdapter(query, sqlConnection);
            DataTable dataSet = new DataTable();
            command.Fill(dataSet);
            chart1.DataSource = dataSet;
            chart1.Series[0].XValueMember = "y";
            chart1.Legends[0].Name = "d";
            chart1.Series[0].YValueMembers = "d";
                    chart1.DataBind();
           
            sqlConnection.Close();
        }
        public void Static_loadsс2()
        {
            chart2.Series[0].Points.Clear();
            string query2 = $@"Select count( working.idworker)
from working inner join worker on working.idworker=worker.idworker 
where    (status='Выполнено' or status='Позже срока')
group by (worker.firstname + ' ' + worker.name + ' ' + worker.lastname)";
            DataTable data2 = new DataTable();
            SqlDataAdapter command2 = new SqlDataAdapter(query2, sqlConnection);
            command2.Fill(data2);
            DataColumn column2 = data2.Columns[0];
            DataRow row2 = data2.Rows[0];
            int y = Convert.ToInt32(row2[column2].ToString());

            sqlConnection.Open();
            string query = $@"Select count( working.idworker),(worker.firstname + ' ' + worker.name + ' ' + worker.lastname)  from working inner join worker on working.idworker=worker.idworker
where    (status='Выполнено' or status='Позже срока')
group by (worker.firstname + ' ' + worker.name + ' ' + worker.lastname) ";
            SqlDataAdapter command = new SqlDataAdapter(query, sqlConnection);
            DataTable dataSet = new DataTable();
            command.Fill(dataSet);
            chart2.DataSource = dataSet;
            sqlConnection.Close();
            try
            {
                for (int j = 0; j < y; j++)
                {
                    DataColumn column = dataSet.Columns[0];
                    DataRow row = dataSet.Rows[j];
                    DataColumn column3 = dataSet.Columns[1];
                    DataRow row3 = dataSet.Rows[j];
                    chart2.Series[0].Points.AddXY(row3[column3].ToString(), row[column].ToString());
                }
            }
            catch
            {

            }
        }

        public void Static_loadscс2()
        {
            chart3.Series[0].Points.Clear();
            string query2 = $@"select count( bid.idbid)
from bid
inner join worker on bid.idworker = worker.idworker
where  (status='Выполнено' or status='Позже срока')";
            DataTable data2 = new DataTable();
            SqlDataAdapter command2 = new SqlDataAdapter(query2, sqlConnection);
            command2.Fill(data2);
            DataColumn column2 = data2.Columns[0];
            DataRow row2 = data2.Rows[0];
            int y = Convert.ToInt32(row2[column2].ToString());

            sqlConnection.Open();
            string query = $@"SELECT 
 
    COUNT(bid.idbid) AS [Количество заявок] ,  CONCAT(DATEPART(MONTH, datec), '-',DATEPART(YEAR, datec) ) AS [Год-Месяц]
FROM 
    bid
INNER JOIN 
    worker ON bid.idworker = worker.idworker
WHERE 
    (status = 'Выполнено' OR status = 'Позже срока')
GROUP BY 
    DATEPART(YEAR, datec), 
    DATEPART(MONTH, datec)
ORDER BY 
    DATEPART(YEAR, datec), 
    DATEPART(MONTH, datec)";
            SqlDataAdapter command = new SqlDataAdapter(query, sqlConnection);
            DataTable dataSet = new DataTable();
            command.Fill(dataSet);

            string query12 = $@"select Distinct  CONCAT(DATEPART(MONTH, datec), '-',DATEPART(YEAR, datec) )
from bid
inner join worker on bid.idworker = worker.idworker
where  (status='Выполнено' or status='Позже срока')";
            SqlDataAdapter command12 = new SqlDataAdapter(query12, sqlConnection);
            DataTable dataSet12 = new DataTable();
            command12.Fill(dataSet12);
            chart3.DataSource = dataSet12;
            sqlConnection.Close();
            try
            {
                for (int j = 0; j < y; j++)
                {
                    DataColumn column = dataSet.Columns[0];
                    DataRow row = dataSet.Rows[j];
                    DataColumn column3 = dataSet12.Columns[0];
                    DataRow row3 = dataSet12.Rows[j];
                    chart3.Series[0].Points.AddXY(row3[column3].ToString(), row[column].ToString());
                }
            }
            catch
            {

            }
        }
        public void Static_loadscсс2()
        {
            //Фрагмент кода для статистики 
            chart4.Series[0].Points.Clear();
            string query2 = $@"select count(bid.idbid)
from bid
inner join object on bid.idobject=object.idobject 
where  (status='Выполнено' or status='Позже срока')";
            DataTable data2 = new DataTable();
            SqlDataAdapter command2 = new SqlDataAdapter(query2, sqlConnection);
            command2.Fill(data2);
            DataColumn column2 = data2.Columns[0];
            DataRow row2 = data2.Rows[0];
            int y = Convert.ToInt32(row2[column2].ToString());
            sqlConnection.Open();
            string query = $@"select count(bid.idbid),type
from bid
inner join object on bid.idobject=object.idobject 
where  (status='Выполнено' or status='Позже срока')
group by type";
            SqlDataAdapter command = new SqlDataAdapter(query, sqlConnection);
            DataTable dataSet = new DataTable();
            command.Fill(dataSet);
            sqlConnection.Close();
            try
            {
                for (int j = 0; j < y; j++)
                {
                    DataColumn column = dataSet.Columns[0];
                    DataRow row = dataSet.Rows[j];
                    DataColumn column3 = dataSet.Columns[1];
                    DataRow row3 = dataSet.Rows[j];
                    chart4.Series[0].Points.AddXY(row3[column3].ToString(), row[column].ToString());
                }
            }
            catch
            {

            }
        }
        //

        private void Static_Load(object sender, EventArgs e)
        {
            string query2 = $@"Select Min(datec) from bid";
            System.Data.DataTable data2 = new System.Data.DataTable();
            SqlDataAdapter command2 = new SqlDataAdapter(query2, sqlConnection);
            command2.Fill(data2);
            DataColumn column2 = data2.Columns[0];
            DataRow row2 = data2.Rows[0];
            dateTimePicker1.MinDate = Convert.ToDateTime(row2[column2].ToString());
            dateTimePicker2.MaxDate = DateTime.Today;
            dateTimePicker1.MaxDate = DateTime.Today;
            Static_loads2();
            Static_loadsс2();
            Static_loadscс2();
            Static_loadscсс2();
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            dateTimePicker2.MinDate = dateTimePicker1.Value;
           
        }

        private void dateTimePicker1_DropDown(object sender, EventArgs e)
        {
            string query2 = $@"Select Min(datec) from bid";
            System.Data.DataTable data2 = new System.Data.DataTable();
            SqlDataAdapter command2 = new SqlDataAdapter(query2, sqlConnection);
            command2.Fill(data2);
            DataColumn column2 = data2.Columns[0];
            DataRow row2 = data2.Rows[0];
            dateTimePicker1.MinDate = Convert.ToDateTime(row2[column2].ToString());
            dateTimePicker1.MaxDate = DateTime.Today;
            dateTimePicker2.MaxDate = DateTime.Today;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Static_loads2();
            Static_loadsс2();
            Static_loadscс2();
            Static_loadscсс2();

        }
       
        private void dateTimePicker1_CloseUp(object sender, EventArgs e)
        {
            Static_loads();
            Static_loadsс();
            Static_loadscс();
            Static_loadscсс();

        }

        private void dateTimePicker2_CloseUp(object sender, EventArgs e)
        {
            Static_loads();
            Static_loadsс();
            Static_loadscс();
            Static_loadscсс();
        }

        private void dateTimePicker1_ValueChanged_1(object sender, EventArgs e)
        {
            dateTimePicker2.MinDate = dateTimePicker1.Value;
        }
    }
}
