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
            string query2 = $@"Select count( treaty.idinsurer)
from treaty 
inner join insurer on treaty.idinsurer=insurer.idinsurer ";
            DataTable data2 = new DataTable();
            SqlDataAdapter command2 = new SqlDataAdapter(query2, sqlConnection);
            command2.Fill(data2);
            DataColumn column2 = data2.Columns[0];
            DataRow row2 = data2.Rows[0];
            int y =Convert.ToInt32( row2[column2].ToString());

            sqlConnection.Open();
            string query = $@"Select count( treaty.idinsurer),(insurer.firstname+' '+insurer.name+' '+insurer.lastname) as [Страховщик]
from treaty
inner join insurer on treaty.idinsurer = insurer.idinsurer
where dateconclusion>='{so}' and dateconclusion<='{st}'
group by (insurer.firstname + ' ' + insurer.name + ' ' + insurer.lastname)";
            SqlDataAdapter command = new SqlDataAdapter(query, sqlConnection);
            DataTable dataSet = new DataTable();
            command.Fill(dataSet);
            chart1.DataSource = dataSet;
            sqlConnection.Close();
            try
            {
                for (int j = 0; j < y; j++)
                {
                    DataColumn column = dataSet.Columns[0];
                    DataRow row = dataSet.Rows[j];
                    DataColumn column3 = dataSet.Columns[1];
                    DataRow row3 = dataSet.Rows[j];
                    chart1.Series[0].Points.AddXY(row3[column3].ToString(), row[column].ToString());
                }
            }
            catch
            {

            }
        }
        
        public void Static_loadsс()
        { so = dateTimePicker1.Value.ToString("yyyy-MM-dd");
            st = dateTimePicker2.Value.ToString("yyyy-MM-dd");
            chart2.Series[0].Points.Clear();
            string query2 = $@"Select count( treaty.idpolicyholder)
from treaty 
inner join policyholder on treaty.idpolicyholder=policyholder.idpolicyholder  ";
            DataTable data2 = new DataTable();
            SqlDataAdapter command2 = new SqlDataAdapter(query2, sqlConnection);
            command2.Fill(data2);
            DataColumn column2 = data2.Columns[0];
            DataRow row2 = data2.Rows[0];
            int y = Convert.ToInt32(row2[column2].ToString());

            sqlConnection.Open();
            string query = $@"Select count( treaty.idpolicyholder),(policyholder.firdtname+' '+policyholder.name+' '+policyholder.lastname) as Страхователь from treaty inner join policyholder on treaty.idpolicyholder=policyholder.idpolicyholder where dateconclusion>='{so}' and dateconclusion<='{st}' group by (policyholder.firdtname+' '+policyholder.name+' '+policyholder.lastname) ";
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
            string query2 = $@"select count( treaty.idtreaty)
from treaty 
inner join policyholder on treaty.idpolicyholder=policyholder.idpolicyholder ";
            DataTable data2 = new DataTable();
            SqlDataAdapter command2 = new SqlDataAdapter(query2, sqlConnection);
            command2.Fill(data2);
            DataColumn column2 = data2.Columns[0];
            DataRow row2 = data2.Rows[0];
            int y = Convert.ToInt32(row2[column2].ToString());

            sqlConnection.Open();
            string query = $@"select count( treaty.idtreaty),treaty.dateconclusion
from treaty 
inner join policyholder on treaty.idpolicyholder=policyholder.idpolicyholder 
where dateconclusion>='{so}' and dateconclusion<='{st}'
group by treaty.dateconclusion";
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
                    chart3.Series[0].Points.AddXY(row3[column3].ToString().Substring(0,10),row[column].ToString());
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
            string query2 = $@"select count(treaty.idtreaty)
from treaty 
inner join vid on vid.idvida=treaty.idvida  ";
            DataTable data2 = new DataTable();
            SqlDataAdapter command2 = new SqlDataAdapter(query2, sqlConnection);
            command2.Fill(data2);
            DataColumn column2 = data2.Columns[0];
            DataRow row2 = data2.Rows[0];
            int y = Convert.ToInt32(row2[column2].ToString());

            sqlConnection.Open();
            string query = $@"select count(treaty.idtreaty),name
from treaty 
inner join vid on vid.idvida=treaty.idvida 
where dateconclusion>='{so}' and dateconclusion<='{st}'
group by name";
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
            //dateTimePicker2.MaxDate = DateTime.Today;
            //dateTimePicker1.MaxDate = DateTime.Today;
            chart1.Series[0].Points.Clear();
            string query2 = $@"Select count( treaty.idinsurer)
from treaty 
inner join insurer on treaty.idinsurer=insurer.idinsurer ";
            DataTable data2 = new DataTable();
            SqlDataAdapter command2 = new SqlDataAdapter(query2, sqlConnection);
            command2.Fill(data2);
            DataColumn column2 = data2.Columns[0];
            DataRow row2 = data2.Rows[0];
            int y = Convert.ToInt32(row2[column2].ToString());

            sqlConnection.Open();
            string query = $@"Select count( treaty.idinsurer),(insurer.firstname+' '+insurer.name+' '+insurer.lastname) as [Страховщик]
from treaty
inner join insurer on treaty.idinsurer = insurer.idinsurer
group by (insurer.firstname + ' ' + insurer.name + ' ' + insurer.lastname)";
            SqlDataAdapter command = new SqlDataAdapter(query, sqlConnection);
            DataTable dataSet = new DataTable();
            command.Fill(dataSet);
            chart1.DataSource = dataSet;
            sqlConnection.Close();
            try
            {
                for (int j = 0; j < y; j++)
                {
                    DataColumn column = dataSet.Columns[0];
                    DataRow row = dataSet.Rows[j];
                    DataColumn column3 = dataSet.Columns[1];
                    DataRow row3 = dataSet.Rows[j];
                    chart1.Series[0].Points.AddXY(row3[column3].ToString(), row[column].ToString());
                }
            }
            catch
            {

            }
        }
        public void Static_loadsс2()
        {
            //dateTimePicker2.MaxDate = DateTime.Today;
            //dateTimePicker1.MaxDate = DateTime.Today;
            chart2.Series[0].Points.Clear();
            string query2 = $@"Select count( treaty.idpolicyholder)
from treaty 
inner join policyholder on treaty.idpolicyholder=policyholder.idpolicyholder  ";
            DataTable data2 = new DataTable();
            SqlDataAdapter command2 = new SqlDataAdapter(query2, sqlConnection);
            command2.Fill(data2);
            DataColumn column2 = data2.Columns[0];
            DataRow row2 = data2.Rows[0];
            int y = Convert.ToInt32(row2[column2].ToString());

            sqlConnection.Open();
            string query = $@"Select count( treaty.idpolicyholder),(policyholder.firdtname+' '+policyholder.name+' '+policyholder.lastname) as Страхователь
from treaty
inner join policyholder on treaty.idpolicyholder=policyholder.idpolicyholder 
group by (policyholder.firdtname+' '+policyholder.name+' '+policyholder.lastname) ";
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
            //dateTimePicker2.MaxDate = DateTime.Today;
            //dateTimePicker1.MaxDate = DateTime.Today;
            chart3.Series[0].Points.Clear();
            string query2 = $@"select count( treaty.idtreaty)
from treaty 
inner join policyholder on treaty.idpolicyholder=policyholder.idpolicyholder ";
            DataTable data2 = new DataTable();
            SqlDataAdapter command2 = new SqlDataAdapter(query2, sqlConnection);
            command2.Fill(data2);
            DataColumn column2 = data2.Columns[0];
            DataRow row2 = data2.Rows[0];
            int y = Convert.ToInt32(row2[column2].ToString());

            sqlConnection.Open();
            string query = $@"select count( treaty.idtreaty),treaty.dateconclusion
from treaty 
inner join policyholder on treaty.idpolicyholder=policyholder.idpolicyholder 
group by treaty.dateconclusion";
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
                    chart3.Series[0].Points.AddXY(row3[column3].ToString().Substring(0, 10), row[column].ToString());
                }
            }
            catch
            {

            }
        }
        public void Static_loadscсс2()
        {
            //dateTimePicker2.MaxDate = DateTime.Today;
            //dateTimePicker1.MaxDate = DateTime.Today;
            chart4.Series[0].Points.Clear();
            string query2 = $@"select count(treaty.idtreaty)
from treaty 
inner join vid on vid.idvida=treaty.idvida  ";
            DataTable data2 = new DataTable();
            SqlDataAdapter command2 = new SqlDataAdapter(query2, sqlConnection);
            command2.Fill(data2);
            DataColumn column2 = data2.Columns[0];
            DataRow row2 = data2.Rows[0];
            int y = Convert.ToInt32(row2[column2].ToString());

            sqlConnection.Open();
            string query = $@"select count(treaty.idtreaty),name
from treaty 
inner join vid on vid.idvida=treaty.idvida 
group by name";
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

        private void Static_Load(object sender, EventArgs e)
        {
            string query2 = $@"Select Min(dateconclusion) from treaty";
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
            string query2 = $@"Select Min(dateconclusion) from treaty";
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
