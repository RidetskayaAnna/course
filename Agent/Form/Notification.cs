using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Word = Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;

namespace Agent
{
    public partial class Notification : UserControl
    {
        datebase db = new datebase();
        public SqlConnection sqlConnection = null;
        Workerak insurerak;
        public Notification(Workerak insurerak1)
        {
            InitializeComponent();
            sqlConnection = new SqlConnection(db.connection);
            insurerak = insurerak1;
        }
        Form1 form1;
        public Notification(Form1 form)
        {
            InitializeComponent();
            sqlConnection = new SqlConnection(db.connection);
            form1 = form;
        }
        Startcs startcs;
        public Notification(Startcs startcs1)
        {
            InitializeComponent();
            sqlConnection = new SqlConnection(db.connection);
            startcs=startcs1;
        }

        int id = 0;
        private void Notification_Load(object sender, EventArgs e)
        {
            if (insurerak!=null)
            {
                panel1.Visible = false;
                panel2.Visible = true;
                comboBoxpolicyholder2();
            }

            dataGridView1.RowsDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }
        public void comboBoxpolicyholder2()
        {

            sqlConnection.Open();
            string query = "select idtenant,(firstname+' '+name+' '+lastname) as p from tenant";
            SqlDataAdapter sqlDbDataAdapter = new SqlDataAdapter(query, sqlConnection);
            DataSet dataSet = new DataSet();
            sqlDbDataAdapter.Fill(dataSet);
            comboBox2.DataSource = dataSet.Tables[0];
            comboBox2.DisplayMember = "p";
            comboBox2.ValueMember = "idtenant";
            comboBox2.SelectedIndex = -1;
            sqlConnection.Close();

        }
        public void policyholder()
        {
            try
            {
                sqlConnection.Close();
                sqlConnection.Open();
                DataSet dataSet = new DataSet();
                SqlDataAdapter command = new SqlDataAdapter($@"Select 
 (service.name) as Услуга, 
 (type+', '+address) as Объект,
 datec as [Дата выполнения] 
from bid inner join object on bid.idobject=object.idobject inner join tenant on bid.idtenant=tenant.idtenant 
inner join worker on bid.idworker=worker.idworker inner join service on bid.idservice=service.idservice
where  bid.idtenant={comboBox2.SelectedValue}  and (status='Выполнено' or status='Позже срока')", sqlConnection);
                command.Fill(dataSet);
                dataGridView1.DataSource = dataSet.Tables[0].DefaultView;
                sqlConnection.Close();
                dataGridView1.Columns[1].Visible = false;
                
                dataGridView1.AllowUserToAddRows = false;
                sqlConnection.Close();
            }
            catch { }

        }
        private void button1_Click(object sender, EventArgs e)
        {
           panel2.Visible = true;
            label3.Text = "Пользователь";
                comboBoxpolicyholder2();
                textBox10.Text = "";
                textBox2.Text = "";
                checkBox1.Checked = false;
            if (dataGridView1 != null)
            {
                dataGridView1.DataSource = null; // Сбрасывает источник данных
                dataGridView1.Rows.Clear();

            }


           
          

        }
        public void comboBoxinsurer()
        {
            sqlConnection.Close();
            sqlConnection.Open();
            string query = "select idworker,(firstname+' '+name+' '+lastname) as i from worker where datelayoffs is null";
            SqlDataAdapter sqlDbDataAdapter = new SqlDataAdapter(query, sqlConnection);
            DataSet dataSet = new DataSet();
            sqlDbDataAdapter.Fill(dataSet);
            comboBox2.DataSource = dataSet.Tables[0];
            comboBox2.DisplayMember = "i";
            comboBox2.ValueMember = "idworker";
            comboBox2.SelectedIndex = -1;
            sqlConnection.Close();
        }
        public void insurer()
        {
            try
            {
                sqlConnection.Close();
                sqlConnection.Open();
                DataSet dataSet = new DataSet();
                SqlDataAdapter command = new SqlDataAdapter($@"Select idbid as [Номер заявки],  (service.name) as Услуга, 
(type+', '+address) as Объект,datec
from bid inner join object on bid.idobject=object.idobject inner join tenant on bid.idtenant=tenant.idtenant 
inner join worker on bid.idworker=worker.idworker inner join service on bid.idservice=service.idservice
where  bid.idworker={comboBox2.SelectedValue} and (status='Выполнено' or status='Позже срока')
union
Select idworking as [Номер заявки],  (name) as Услуга, 
 (type+', '+address) as Объект, date
from working inner join object on working.idobject=object.idobject 
 where  working.idworker={comboBox2.SelectedValue} and (status='Выполнено' or status='Позже срока') ", sqlConnection);
                command.Fill(dataSet);
                dataGridView1.DataSource = dataSet.Tables[0].DefaultView;
                sqlConnection.Close();
                dataGridView1.Columns[0].Visible = false;
                
                dataGridView1.AllowUserToAddRows = false;
                sqlConnection.Close();
            }
            catch { }

        }
        private void button2_Click(object sender, EventArgs e)
        {
                panel2.Visible = true;
                label3.Text = "Работник";
                comboBoxinsurer();
                textBox10.Text = "";
                textBox2.Text = "";
                checkBox1.Checked = false;
            if (dataGridView1 != null)
            {
                dataGridView1.DataSource = null; // Сбрасывает источник данных
                dataGridView1.Rows.Clear();

            }
        }
        private void ReplaceWordStubs(string stubToReplace, string text, Word.Document WordDoc)
        {
            var range = WordDoc.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: stubToReplace, ReplaceWith: text);
        }
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {   if (label3.Text == "Работник")
            {
                try
                {
                    string query1 = $@"Select 
email as Почта
from worker 
 where idworker={comboBox2.SelectedValue}";
                    DataTable data = new DataTable();
                    SqlDataAdapter command1 = new SqlDataAdapter(query1, sqlConnection);
                    command1.Fill(data);
                    DataColumn column = data.Columns[0];
                    DataRow row = data.Rows[0];
                    textBox10.Text = row[column].ToString();
                }
                catch { }
                insurer();
            }
            else
            {
                try
                {
                    string query1 = $@"Select 
email as Почта
from tenant
 where idtenant={comboBox2.SelectedValue}";
                    DataTable data = new DataTable();
                    SqlDataAdapter command1 = new SqlDataAdapter(query1, sqlConnection);
                    command1.Fill(data);
                    DataColumn column = data.Columns[0];
                    DataRow row = data.Rows[0];
                    textBox10.Text = row[column].ToString();
                }
                catch { }
                    policyholder();
            }
            
             
        }
        string strax = "";
        string strax2 = "";
        string pas = "";
        string vznos = "";
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (checkBox1.Checked == true)
                { if (id != 0)
                        {
                    if (dataGridView1.CurrentRow == null)
                    {
                        throw new Exception();
                    }
                    else
                    {
//                            if (dataGridView1.CurrentRow.Cells[6].Value.ToString() == "Медицинское страхование")
//                            {
//                                strax = "";
//                                strax2 = "";
//                                pas = "";
//                                vznos = "";
//                                Word.Application wordApplication = new Word.Application();

//                                string PathToNote = @"/Diplom/proga/Agent/treaty/med/med.docx";
//                                Word.Application application = new Word.Application();
//                                application.Visible = false;
//                                Word.Document document = application.Documents.Open(PathToNote);
//                                ReplaceWordStubs("{номер}", dataGridView1.CurrentRow.Cells[0].Value.ToString(), document);
//                                ReplaceWordStubs("{число}", dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(0, 2), document);
//                                ReplaceWordStubs("{месяц}", dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(3, 2), document);
//                                ReplaceWordStubs("{год}", dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(6, 4), document);
//                                ReplaceWordStubs("{страховщик}", dataGridView1.CurrentRow.Cells[2].Value.ToString(), document);
//                                strax = dataGridView1.CurrentRow.Cells[4].Value.ToString();
//                                string query45 = $@"select date from treaty 
//inner join bid on bid.idbid=treaty.idbid where treaty.idbid={dataGridView1.CurrentRow.Cells[24].Value.ToString()}";
//                                DataTable data45 = new DataTable();
//                                SqlDataAdapter command45 = new SqlDataAdapter(query45, sqlConnection);
//                                command45.Fill(data45);
//                                DataColumn column45 = data45.Columns[0];
//                                DataRow row45 = data45.Rows[0];

//                                ReplaceWordStubs("{датазаявки}", row45[column45].ToString().Substring(0, 10), document);

//                                string query1 = $@"Select dateb as [Дата рождения],city.indexcity,
//city.name as [Город прописки],address as [Адрес], passport as [Паспорт],
//numar as [Идентификационный номер],organ as [Орган, выдавший паспорт],
//datep as [Дата выдачи паспорта],phone as Телефон,work.name as [Место работы],
//position.name as [Должность], heal as [Группа здоровья],sport as [Занятие спортом],
//email as Почта
//from policyholder inner join city on policyholder.idcity=city.idcity 
//inner join position on position.idposition=policyholder.idwork 
//inner join work on work.idwork=position.idwork
// where idpolicyholder={dataGridView1.CurrentRow.Cells[3].Value.ToString()}";
//                                DataTable data = new DataTable();
//                                SqlDataAdapter command1 = new SqlDataAdapter(query1, sqlConnection);
//                                command1.Fill(data);
//                                for (int i = 0; i < 10; i++)
//                                {
//                                    DataColumn column = data.Columns[i];
//                                    DataRow row = data.Rows[0];


//                                    if (i == 1)
//                                    {
//                                        strax = strax + "; " + row[column].ToString();

//                                        pas = pas + row[column].ToString();
//                                    }
//                                    else
//                                    if (i == 2)
//                                    {
//                                        strax = strax + ", г." + row[column].ToString();
//                                        pas = pas + ", г." + row[column].ToString();
//                                    }
//                                    else
//                                    if (i == 3)
//                                    {
//                                        strax = strax + ", " + row[column].ToString();
//                                        pas = pas + ", " + row[column].ToString();
//                                    }
//                                    else
//                                    if (i == 4)
//                                    {
//                                        strax = strax + ";";
//                                        strax2 = strax2 + " паспорт " + row[column].ToString();
//                                    }
//                                    else
//                                    if (i == 5)
//                                    {
//                                        strax2 = strax2 + ", идентификационный номер " + row[column].ToString();
//                                    }
//                                    else
//                                    if (i == 6)
//                                    {
//                                        strax2 = strax2 + ", выдан  " + row[column].ToString();
//                                    }
//                                    else
//                                    if (i == 7)
//                                    {
//                                        strax2 = strax2 + "; " + row[column].ToString().Substring(0, 10);
//                                    }
//                                    else if (i == 8)
//                                    {
//                                        pas = pas + ", " + row[column].ToString();
//                                    }

//                                }

//                                ReplaceWordStubs("{страхователь}", strax, document);
//                                ReplaceWordStubs("{паспорт}", strax2, document);
//                                ReplaceWordStubs("{фио}", dataGridView1.CurrentRow.Cells[4].Value.ToString(), document);
//                                DataColumn column5 = data.Columns[0];
//                                DataRow row5 = data.Rows[0];
//                                ReplaceWordStubs("{датаро}", row5[column5].ToString().Substring(0, 10), document);
//                                ReplaceWordStubs("{паспорт}", strax2 + pas, document);

//                                ReplaceWordStubs("{страхс}", dataGridView1.CurrentRow.Cells[11].Value.ToString(), document);
//                                ReplaceWordStubs("{страхп}", dataGridView1.CurrentRow.Cells[15].Value.ToString(), document);

//                                DataColumn column10 = data.Columns[11];
//                                DataRow row10 = data.Rows[0];
//                                if (row10[column10].ToString() == "Д-1")
//                                {
//                                    ReplaceWordStubs("{х}", "X", document);
//                                    ReplaceWordStubs("{х}", "", document);
//                                    ReplaceWordStubs("{х}", "", document);
//                                    ReplaceWordStubs("{х}", "", document);
//                                }
//                                else if (row10[column10].ToString() == "Д-2")
//                                {
//                                    ReplaceWordStubs("{х}", "", document);
//                                    ReplaceWordStubs("{х}", "X", document);
//                                    ReplaceWordStubs("{х}", "", document);
//                                    ReplaceWordStubs("{х}", "", document);
//                                }
//                                else if (row10[column10].ToString() == "Д-3")
//                                {
//                                    ReplaceWordStubs("{х}", "", document);
//                                    ReplaceWordStubs("{х}", "", document);
//                                    ReplaceWordStubs("{х}", "X", document);
//                                    ReplaceWordStubs("{х}", "", document);
//                                }
//                                else
//                                {
//                                    ReplaceWordStubs("{х}", "", document);
//                                    ReplaceWordStubs("{х}", "", document);
//                                    ReplaceWordStubs("{х}", "", document);
//                                    ReplaceWordStubs("{х}", "X", document);
//                                }

//                                ReplaceWordStubs("{страхс}", dataGridView1.CurrentRow.Cells[11].Value.ToString(), document);

//                                ReplaceWordStubs("{датан}", dataGridView1.CurrentRow.Cells[9].Value.ToString().Substring(0, 10), document);
//                                ReplaceWordStubs("{датак}", dataGridView1.CurrentRow.Cells[10].Value.ToString().Substring(0, 10), document);

//                                ReplaceWordStubs("{коэф}", dataGridView1.CurrentRow.Cells[12].Value.ToString(), document);
//                                ReplaceWordStubs("{страхп}", dataGridView1.CurrentRow.Cells[15].Value.ToString(), document);

//                                string query2 = $@"Select  pay.vidpay,datepay,summa
//from pay inner join  treaty on treaty.idtreaty=pay.idtreaty
// where treaty.idtreaty={dataGridView1.CurrentRow.Cells[0].Value.ToString()}";
//                                DataTable data2 = new DataTable();
//                                SqlDataAdapter command2 = new SqlDataAdapter(query2, sqlConnection);
//                                command2.Fill(data2);
//                                DataColumn column1 = data2.Columns[0];
//                                DataRow row1 = data2.Rows[0];
//                                ReplaceWordStubs("{оплата}", row1[column1].ToString(), document);

//                                if (dataGridView1.CurrentRow.Cells[8].Value.ToString() == "Единовременно")
//                                {
//                                    ReplaceWordStubs("{х}", "X", document);
//                                    ReplaceWordStubs("{х}", "", document);
//                                    ReplaceWordStubs("{х}", "", document);

//                                }
//                                else if (dataGridView1.CurrentRow.Cells[8].Value.ToString() == "В два срока")
//                                {
//                                    ReplaceWordStubs("{х}", "", document);
//                                    ReplaceWordStubs("{х}", "X", document);
//                                    ReplaceWordStubs("{х}", "", document);

//                                }
//                                else
//                                {
//                                    ReplaceWordStubs("{х}", "", document);
//                                    ReplaceWordStubs("{х}", "", document);
//                                    ReplaceWordStubs("{х}", "X", document);
//                                }
//                                Random random = new Random();
//                                ReplaceWordStubs("{страх}", dataGridView1.CurrentRow.Cells[4].Value.ToString(), document);
//                                ReplaceWordStubs("{взнос}", dataGridView1.CurrentRow.Cells[17].Value.ToString(), document);
//                                ReplaceWordStubs("{дата}", dataGridView1.CurrentRow.Cells[16].Value.ToString().Substring(0, 10), document);
//                                ReplaceWordStubs("{платеж}", row1[column1].ToString() + " №" + random.Next(0, 100), document);

//                                if (dataGridView1.CurrentRow.Cells[18].Value.ToString() != "")
//                                {
//                                    vznos = vznos + " " + dataGridView1.CurrentRow.Cells[18].Value.ToString().Substring(0, 10) + "; " + dataGridView1.CurrentRow.Cells[19].Value.ToString() + ";";
//                                    if (dataGridView1.CurrentRow.Cells[20].Value.ToString() != "")
//                                    {
//                                        vznos = vznos + " " + dataGridView1.CurrentRow.Cells[20].Value.ToString().Substring(0, 10) + "; " + dataGridView1.CurrentRow.Cells[21].Value.ToString() + ";";
//                                        vznos = vznos + " " + dataGridView1.CurrentRow.Cells[22].Value.ToString().Substring(0, 10) + "; " + dataGridView1.CurrentRow.Cells[23].Value.ToString() + ";";

//                                    }
//                                }
//                                else { vznos = ""; }
//                                ReplaceWordStubs("{взносы}", vznos, document);
//                                ReplaceWordStubs("{страхов}", dataGridView1.CurrentRow.Cells[2].Value.ToString(), document);

//                                DataColumn column11 = data.Columns[8];
//                                DataRow row11 = data.Rows[0];

//                                ReplaceWordStubs("{тел}", row11[column11].ToString(), document);
//                                if (dataGridView1.Rows.Count != 0)
//                                {
//                                    if (File.Exists($@"/Diplom/proga/Agent/treaty/med/med{id}.doc") == true)
//                                    {

//                                    }
//                                    else
//                                    {

//                                        try
//                                        {
//                                            document.SaveAs2($@"/Diplom/proga/Agent/treaty/med/med{id}.doc");
//                                            document.Close();
//                                            application.Quit();
//                                        }
//                                        catch { MessageBox.Show("Ошибка!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }
//                                    }
//                                }

//                            }
//                            else if (dataGridView1.CurrentRow.Cells[6].Value.ToString() == "Страхование от несчатного случая")
//                            {
//                                strax = "";
//                                strax2 = "";
//                                pas = "";
//                                vznos = "";
//                                Word.Application wordApplication = new Word.Application();

//                                string PathToNote = @"/Diplom/proga/Agent/treaty/nes/nes.docx";
//                                Word.Application application = new Word.Application();
//                                application.Visible = false;
//                                Word.Document document = application.Documents.Open(PathToNote);
//                                ReplaceWordStubs("{номер}", dataGridView1.CurrentRow.Cells[0].Value.ToString(), document);

//                                ReplaceWordStubs("{страховщик}", dataGridView1.CurrentRow.Cells[2].Value.ToString(), document);
//                                strax = dataGridView1.CurrentRow.Cells[4].Value.ToString();

//                                string query1 = $@"Select dateb as [Дата рождения],city.indexcity,
//city.name as [Город прописки],address as [Адрес], passport as [Паспорт],
//numar as [Идентификационный номер],organ as [Орган, выдавший паспорт],
//datep as [Дата выдачи паспорта],phone as Телефон,work.name as [Место работы],
//position.name as [Должность], heal as [Группа здоровья],sport as [Занятие спортом],
//email as Почта,position.harmhul
//from policyholder inner join city on policyholder.idcity=city.idcity 
//inner join position on position.idposition=policyholder.idwork 
//inner join work on work.idwork=position.idwork
// where idpolicyholder={dataGridView1.CurrentRow.Cells[3].Value.ToString()}";
//                                DataTable data = new DataTable();
//                                SqlDataAdapter command1 = new SqlDataAdapter(query1, sqlConnection);
//                                command1.Fill(data);
//                                for (int i = 0; i < 10; i++)
//                                {
//                                    DataColumn column = data.Columns[i];
//                                    DataRow row = data.Rows[0];


//                                    if (i == 1)
//                                    {
//                                        strax = strax + "; " + row[column].ToString();

//                                        pas = pas + row[column].ToString();
//                                    }
//                                    else
//                                    if (i == 2)
//                                    {
//                                        strax = strax + ", г." + row[column].ToString();
//                                        pas = pas + ", г." + row[column].ToString();
//                                    }
//                                    else
//                                    if (i == 3)
//                                    {
//                                        strax = strax + ", " + row[column].ToString();
//                                        pas = pas + ", " + row[column].ToString();
//                                    }
//                                    else
//                                    if (i == 4)
//                                    {
//                                        strax = strax + ";";
//                                        strax2 = strax2 + " паспорт " + row[column].ToString();
//                                    }
//                                    else
//                                    if (i == 5)
//                                    {
//                                        strax2 = strax2 + ", идентификационный номер " + row[column].ToString();
//                                    }
//                                    else
//                                    if (i == 6)
//                                    {
//                                        strax2 = strax2 + ", выдан  " + row[column].ToString();
//                                    }
//                                    else
//                                    if (i == 7)
//                                    {
//                                        strax2 = strax2 + "; " + row[column].ToString().Substring(0, 10);
//                                    }
//                                    else if (i == 8)
//                                    {
//                                        pas = pas + ", " + row[column].ToString();
//                                    }

//                                }

//                                ReplaceWordStubs("{страхователь}", strax, document);
//                                ReplaceWordStubs("{паспорт}", strax2, document);

//                                string query45 = $@"select date from treaty 
//inner join bid on bid.idbid=treaty.idbid where treaty.idbid={dataGridView1.CurrentRow.Cells[24].Value.ToString()}";
//                                DataTable data45 = new DataTable();
//                                SqlDataAdapter command45 = new SqlDataAdapter(query45, sqlConnection);
//                                command45.Fill(data45);
//                                DataColumn column45 = data45.Columns[0];
//                                DataRow row45 = data45.Rows[0];

//                                ReplaceWordStubs("{датез}", row45[column45].ToString().Substring(0, 10), document);

//                                ReplaceWordStubs("{страхователь}", dataGridView1.CurrentRow.Cells[4].Value.ToString(), document);

//                                DataColumn column5 = data.Columns[0];
//                                DataRow row5 = data.Rows[0];
//                                ReplaceWordStubs("{датаро}", row5[column5].ToString().Substring(0, 10), document);

//                                DataColumn column15 = data.Columns[9];
//                                DataRow row15 = data.Rows[0];
//                                DataColumn column16 = data.Columns[10];
//                                DataRow row16 = data.Rows[0];
//                                ReplaceWordStubs("{работа}", row15[column15].ToString() + "; " + row16[column16].ToString(), document);

//                                DataColumn column10 = data.Columns[11];
//                                DataRow row10 = data.Rows[0];
//                                if (row10[column10].ToString() == "Д-1")
//                                {
//                                    ReplaceWordStubs("{й}", "", document);
//                                    ReplaceWordStubs("{ц}", "Х", document);

//                                }
//                                else if (row10[column10].ToString() == "Д-2")
//                                {
//                                    ReplaceWordStubs("{й}", "", document);
//                                    ReplaceWordStubs("{ц}", "Х", document);

//                                }
//                                else if (row10[column10].ToString() == "Д-3")
//                                {
//                                    ReplaceWordStubs("{й}", "", document);
//                                    ReplaceWordStubs("{ц}", "Х", document);
//                                }
//                                else
//                                {
//                                    ReplaceWordStubs("{й}", "Х", document);
//                                    ReplaceWordStubs("{ц}", "", document);

//                                }

//                                DataColumn column11 = data.Columns[12];
//                                DataRow row11 = data.Rows[0];
//                                if (row11[column11].ToString() == "Да")
//                                {
//                                    ReplaceWordStubs("{у}", "Х", document);
//                                    ReplaceWordStubs("{к}", "", document);

//                                }
//                                else
//                                {
//                                    ReplaceWordStubs("{у}", "", document);
//                                    ReplaceWordStubs("{к}", "X", document);

//                                }

//                                DataColumn column13 = data.Columns[14];
//                                DataRow row13 = data.Rows[0];

//                                ReplaceWordStubs("{е}", row13[column13].ToString(), document);

//                                List<string> numbers = new List<string>() { "Никитин Николай Михайлович", "Иванова Кристина Тимофеевна", "Зайцев Владимир Даниилович" };

//                                Random rnd = new Random();
//                                int randIndex = rnd.Next(numbers.Count);
//                                string random = numbers[randIndex];

//                                ReplaceWordStubs("{выг}", random, document);
//                                ReplaceWordStubs("{выг}", random, document);

//                                ReplaceWordStubs("{датан}", dataGridView1.CurrentRow.Cells[9].Value.ToString().Substring(0, 10), document);
//                                ReplaceWordStubs("{датак}", dataGridView1.CurrentRow.Cells[10].Value.ToString().Substring(0, 10), document);
//                                ReplaceWordStubs("{страхс}", dataGridView1.CurrentRow.Cells[11].Value.ToString(), document);
//                                ReplaceWordStubs("{бст}", dataGridView1.CurrentRow.Cells[13].Value.ToString(), document);
//                                ReplaceWordStubs("{кк}", dataGridView1.CurrentRow.Cells[12].Value.ToString(), document);
//                                ReplaceWordStubs("{страхт}", dataGridView1.CurrentRow.Cells[14].Value.ToString(), document);
//                                ReplaceWordStubs("{страхп}", dataGridView1.CurrentRow.Cells[15].Value.ToString(), document);

//                                string query2 = $@"Select  pay.vidpay,datepay,summa
//from pay inner join  treaty on treaty.idtreaty=pay.idtreaty
// where treaty.idtreaty={dataGridView1.CurrentRow.Cells[0].Value.ToString()}";
//                                DataTable data2 = new DataTable();
//                                SqlDataAdapter command2 = new SqlDataAdapter(query2, sqlConnection);
//                                command2.Fill(data2);

//                                DataColumn column1 = data2.Columns[0];
//                                DataRow row1 = data2.Rows[0];
//                                if (row1[column1].ToString() == "Наличные")
//                                {
//                                    ReplaceWordStubs("{н}", "X", document);
//                                    ReplaceWordStubs("{г}", "", document);
//                                    ReplaceWordStubs("{оплата}", dataGridView1.CurrentRow.Cells[16].Value.ToString(), document);
//                                    ReplaceWordStubs("{оплата}", "", document);
//                                }
//                                else
//                                {
//                                    ReplaceWordStubs("{н}", "", document);
//                                    ReplaceWordStubs("{г}", "Х", document);
//                                    ReplaceWordStubs("{оплата}", "", document);
//                                    Random random22 = new Random();
//                                    ReplaceWordStubs("{оплата}", dataGridView1.CurrentRow.Cells[16].Value.ToString().Substring(0, 10) + " №" + random22.Next(0, 100), document);
//                                }


//                                if (dataGridView1.CurrentRow.Cells[8].Value.ToString() == "Единовременно")
//                                {
//                                    ReplaceWordStubs("{ш}", "X", document);
//                                    ReplaceWordStubs("{щ}", "", document);
//                                    ReplaceWordStubs("{з}", "", document);

//                                }
//                                else if (dataGridView1.CurrentRow.Cells[8].Value.ToString() == "В два срока")
//                                {
//                                    ReplaceWordStubs("{ш}", "", document);
//                                    ReplaceWordStubs("{щ}", "X", document);
//                                    ReplaceWordStubs("{з}", "", document);

//                                }
//                                else
//                                {
//                                    ReplaceWordStubs("{ш}", "", document);
//                                    ReplaceWordStubs("{щ}", "", document);
//                                    ReplaceWordStubs("{з}", "X", document);
//                                }

//                                ReplaceWordStubs("{взнос}", dataGridView1.CurrentRow.Cells[17].Value.ToString(), document);

//                                if (dataGridView1.CurrentRow.Cells[18].Value.ToString() != "")
//                                {
//                                    ReplaceWordStubs("{датав}", dataGridView1.CurrentRow.Cells[18].Value.ToString().Substring(0, 10), document);
//                                    ReplaceWordStubs("{взнос}", dataGridView1.CurrentRow.Cells[17].Value.ToString(), document);
//                                    ReplaceWordStubs("{датат}", "", document);
//                                    ReplaceWordStubs("{датач}", "", document);
//                                    ReplaceWordStubs("{взнос}", "", document);
//                                    ReplaceWordStubs("{взнос}", "", document);
//                                }
//                                else if (dataGridView1.CurrentRow.Cells[20].Value.ToString() != "")
//                                {
//                                    ReplaceWordStubs("{датав}", dataGridView1.CurrentRow.Cells[18].Value.ToString().Substring(0, 10), document);
//                                    ReplaceWordStubs("{взнос}", dataGridView1.CurrentRow.Cells[17].Value.ToString(), document);
//                                    ReplaceWordStubs("{датат}", dataGridView1.CurrentRow.Cells[20].Value.ToString().Substring(0, 10), document);
//                                    ReplaceWordStubs("{датач}", dataGridView1.CurrentRow.Cells[22].Value.ToString().Substring(0, 10), document);
//                                    ReplaceWordStubs("{взнос}", dataGridView1.CurrentRow.Cells[17].Value.ToString(), document);
//                                    ReplaceWordStubs("{взнос}", dataGridView1.CurrentRow.Cells[17].Value.ToString(), document);
//                                }
//                                else
//                                {
//                                    ReplaceWordStubs("{датав}", "", document);
//                                    ReplaceWordStubs("{взнос}", "", document);
//                                    ReplaceWordStubs("{датат}", "", document);
//                                    ReplaceWordStubs("{датач}", "", document);
//                                    ReplaceWordStubs("{взнос}", "", document);
//                                    ReplaceWordStubs("{взнос}", "", document);
//                                }

//                                ReplaceWordStubs("{страховщик}", dataGridView1.CurrentRow.Cells[4].Value.ToString(), document);

//                                ReplaceWordStubs("{время}", DateTime.Now.ToString().Substring(10, 6), document);
//                                ReplaceWordStubs("{число}", dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(0, 2), document);
//                                ReplaceWordStubs("{месяц}", dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(3, 2), document);
//                                ReplaceWordStubs("{год}", dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(6, 4), document);

//                                DataColumn column12 = data.Columns[8];
//                                DataRow row12 = data.Rows[0];

//                                ReplaceWordStubs("{тел}", row12[column12].ToString(), document);
//                                if (dataGridView1.Rows.Count != 0)
//                                {
//                                    if (File.Exists($@"/Diplom/proga/Agent/treaty/nes/nes{id}.doc") == true)
//                                    {

//                                    }
//                                    else
//                                    {

//                                        try
//                                        {
//                                            document.SaveAs2($@"/Diplom/proga/Agent/treaty/nes/nes{id}.doc");
//                                            document.Close();
//                                            application.Quit();
//                                           }
//                                        catch { MessageBox.Show("Ошибка!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }
//                                    }
//                                }
//                            }
//                            else if (dataGridView1.CurrentRow.Cells[6].Value.ToString() == "Обязательное страхование гражданской ответственности перевозчика перед пассажирами")
//                            {
//                                strax = "";
//                                strax2 = "";
//                                pas = "";
//                                vznos = "";


//                                Word.Application wordApplication = new Word.Application();

//                                string PathToNote = @"/Diplom/proga/Agent/treaty/pere/pere.docx";
//                                Word.Application application = new Word.Application();
//                                application.Visible = false;
//                                Word.Document document = application.Documents.Open(PathToNote);
//                                ReplaceWordStubs("{дата}", dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(0, 10), document);
//                                ReplaceWordStubs("{страхователь}", dataGridView1.CurrentRow.Cells[4].Value.ToString(), document);
//                                ReplaceWordStubs("{датан}", dataGridView1.CurrentRow.Cells[9].Value.ToString().Substring(0, 10), document);
//                                ReplaceWordStubs("{датан}", dataGridView1.CurrentRow.Cells[9].Value.ToString().Substring(0, 10), document);
//                                ReplaceWordStubs("{датак}", dataGridView1.CurrentRow.Cells[10].Value.ToString().Substring(0, 10), document);
//                                ReplaceWordStubs("{страхс}", dataGridView1.CurrentRow.Cells[11].Value.ToString(), document);
//                                string query2 = $@"Select  pay.vidpay,datepay,summa
//from pay inner join  treaty on treaty.idtreaty=pay.idtreaty
// where treaty.idtreaty={dataGridView1.CurrentRow.Cells[0].Value.ToString()}";
//                                DataTable data2 = new DataTable();
//                                SqlDataAdapter command2 = new SqlDataAdapter(query2, sqlConnection);
//                                command2.Fill(data2);
//                                DataColumn column1 = data2.Columns[0];
//                                DataRow row1 = data2.Rows[0];
//                                ReplaceWordStubs("{оплата}", row1[column1].ToString().ToLower(), document);
//                                ReplaceWordStubs("{квзнос}", dataGridView1.CurrentRow.Cells[8].Value.ToString().ToLower(), document);
//                                ReplaceWordStubs("{взнос}", dataGridView1.CurrentRow.Cells[17].Value.ToString(), document);
//                                ReplaceWordStubs("{квзнос}", dataGridView1.CurrentRow.Cells[8].Value.ToString().ToLower(), document);
//                                if (dataGridView1.Rows.Count != 0)
//                                {
//                                    if (File.Exists($@"/Diplom/proga/Agent/treaty/pere/pere{id}.doc") == true)
//                                    {
//                                    }
//                                    else
//                                    {

//                                        try
//                                        {
//                                            document.SaveAs2($@"/Diplom/proga/Agent/treaty/pere/pere{id}.doc");
//                                            document.Close();
//                                            application.Quit();
//                                         }
//                                        catch { MessageBox.Show("Ошибка!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }
//                                    }
//                                }
//                            }
//                            else if (dataGridView1.CurrentRow.Cells[6].Value.ToString() == "Страхование от несчастных случаев пассажиров железного транспорта")
//                            {
//                                strax = "";
//                                strax2 = "";
//                                pas = "";
//                                vznos = "";
//                                Word.Application wordApplication = new Word.Application();

//                                string PathToNote = @"/Diplom/proga/Agent/treaty/pas/pas.docx";
//                                Word.Application application = new Word.Application();
//                                application.Visible = false;
//                                Word.Document document = application.Documents.Open(PathToNote);
//                                ReplaceWordStubs("{номер}", dataGridView1.CurrentRow.Cells[0].Value.ToString(), document);
//                                ReplaceWordStubs("{дата}", dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(0, 10), document);
//                                ReplaceWordStubs("{страхователь}", dataGridView1.CurrentRow.Cells[4].Value.ToString(), document);
//                                ReplaceWordStubs("{страховщик}", dataGridView1.CurrentRow.Cells[2].Value.ToString(), document);
//                                ReplaceWordStubs("{оплата}", dataGridView1.CurrentRow.Cells[8].Value.ToString().ToLower(), document);
//                                ReplaceWordStubs("{взнос}", dataGridView1.CurrentRow.Cells[17].Value.ToString(), document);
//                                ReplaceWordStubs("{число}", dataGridView1.CurrentRow.Cells[16].Value.ToString().Substring(0, 2), document);
//                                ReplaceWordStubs("{датан}", dataGridView1.CurrentRow.Cells[9].Value.ToString().Substring(0, 10), document);
//                                ReplaceWordStubs("{датак}", dataGridView1.CurrentRow.Cells[10].Value.ToString().Substring(0, 10), document);
//                                if (dataGridView1.Rows.Count != 0)
//                                {
//                                    if (File.Exists($@"/Diplom/proga/Agent/treaty/pas/pas{id}.doc") == true)
//                                    {
//                                    }
//                                    else
//                                    {

//                                        try
//                                        {
//                                            document.SaveAs2($@"/Diplom/proga/Agent/treaty/pas/pas{id}.doc");
//                                            document.Close();
//                                            application.Quit();
//                                        }
//                                        catch { MessageBox.Show("Ошибка!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }
//                                    }
//                                }
//                            }
//                            else if (dataGridView1.CurrentRow.Cells[6].Value.ToString() == "Страхование жизни")
//                            {
//                                strax = "";
//                                strax2 = "";
//                                pas = "";
//                                vznos = "";
//                                Word.Application wordApplication = new Word.Application();

//                                string PathToNote = @"/Diplom/proga/Agent/treaty/live/live.docx";
//                                Word.Application application = new Word.Application();
//                                application.Visible = false;
//                                Word.Document document = application.Documents.Open(PathToNote);
//                                ReplaceWordStubs("{номер}", dataGridView1.CurrentRow.Cells[0].Value.ToString(), document);
//                                ReplaceWordStubs("{дата}", dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(0, 10), document);
//                                ReplaceWordStubs("{страховщик}", dataGridView1.CurrentRow.Cells[2].Value.ToString(), document);
//                                string query1 = $@"Select dateb as [Дата рождения],city.indexcity,
//city.name as [Город прописки],address as [Адрес], passport as [Паспорт],
//numar as [Идентификационный номер],organ as [Орган, выдавший паспорт],
//datep as [Дата выдачи паспорта],phone as Телефон,work.name as [Место работы],
//position.name as [Должность], heal as [Группа здоровья],sport as [Занятие спортом],
//email as Почта,position.harmhul
//from policyholder inner join city on policyholder.idcity=city.idcity 
//inner join position on position.idposition=policyholder.idwork 
//inner join work on work.idwork=position.idwork
// where idpolicyholder={dataGridView1.CurrentRow.Cells[3].Value.ToString()}";
//                                DataTable data = new DataTable();
//                                SqlDataAdapter command1 = new SqlDataAdapter(query1, sqlConnection);
//                                command1.Fill(data);
//                                for (int i = 0; i < 10; i++)
//                                {
//                                    DataColumn column = data.Columns[i];
//                                    DataRow row = data.Rows[0];


//                                    if (i == 1)
//                                    {
//                                        strax = strax + " " + row[column].ToString();

//                                        pas = pas + row[column].ToString();
//                                    }
//                                    else
//                                    if (i == 2)
//                                    {
//                                        strax = strax + ", г." + row[column].ToString();
//                                        pas = pas + ", г." + row[column].ToString();
//                                    }
//                                    else
//                                    if (i == 3)
//                                    {
//                                        strax = strax + ", " + row[column].ToString();
//                                        pas = pas + ", " + row[column].ToString();
//                                    }
//                                    else
//                                    if (i == 4)
//                                    {
//                                        strax = strax + ";";
//                                        strax2 = strax2 + " паспорт " + row[column].ToString();
//                                    }
//                                    else
//                                    if (i == 5)
//                                    {
//                                        strax2 = strax2 + ", идентификационный номер " + row[column].ToString();
//                                    }
//                                    else
//                                    if (i == 6)
//                                    {
//                                        strax2 = strax2 + ", выдан  " + row[column].ToString();
//                                    }
//                                    else
//                                    if (i == 7)
//                                    {
//                                        strax2 = strax2 + "; " + row[column].ToString().Substring(0, 10);
//                                    }
//                                    else if (i == 8)
//                                    {
//                                        pas = pas + ", " + row[column].ToString();
//                                    }

//                                }

//                                ReplaceWordStubs("{паспорт}", strax + strax2, document);

//                                DataColumn column5 = data.Columns[0];
//                                DataRow row5 = data.Rows[0];
//                                ReplaceWordStubs("{датаро}", row5[column5].ToString().Substring(0, 10), document);
//                                ReplaceWordStubs("{датаро}", row5[column5].ToString().Substring(0, 10), document);
//                                ReplaceWordStubs("{паспорт}", strax + strax2, document);
//                                List<string> numbers = new List<string>() { "Никитин Николай Михайлович", "Иванова Кристина Тимофеевна", "Зайцев Владимир Даниилович" };

//                                Random rnd = new Random();
//                                int randIndex = rnd.Next(numbers.Count);
//                                string random = numbers[randIndex];

//                                ReplaceWordStubs("{выг}", random, document);

//                                ReplaceWordStubs("{страхс}", Convert.ToString(Math.Round((Convert.ToDecimal(dataGridView1.CurrentRow.Cells[11].Value.ToString()) / 2), 2)), document);
//                                ReplaceWordStubs("{страхсс}", Convert.ToString(Math.Round((Convert.ToDecimal(dataGridView1.CurrentRow.Cells[11].Value.ToString()) / 3), 2)), document);
//                                ReplaceWordStubs("{страхссс}", Convert.ToString(Math.Round((Convert.ToDecimal(dataGridView1.CurrentRow.Cells[11].Value.ToString()) - (Convert.ToDecimal(dataGridView1.CurrentRow.Cells[11].Value.ToString()) / 3) - (Convert.ToDecimal(dataGridView1.CurrentRow.Cells[11].Value.ToString()) / 2)), 2)), document);
//                                ReplaceWordStubs("{страхп}", dataGridView1.CurrentRow.Cells[15].Value.ToString(), document);

//                                ReplaceWordStubs("{квзнос}", dataGridView1.CurrentRow.Cells[8].Value.ToString().ToLower(), document);

//                                if (dataGridView1.CurrentRow.Cells[18].Value.ToString() != "")
//                                {
//                                    ReplaceWordStubs("{датав}", dataGridView1.CurrentRow.Cells[18].Value.ToString().Substring(0, 10), document);

//                                }
//                                else if (dataGridView1.CurrentRow.Cells[20].Value.ToString() != "")
//                                {

//                                    ReplaceWordStubs("{датав}", dataGridView1.CurrentRow.Cells[22].Value.ToString().Substring(0, 10), document);

//                                }
//                                else
//                                {
//                                    ReplaceWordStubs("{датав}", dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(0, 10), document);

//                                }


//                                ReplaceWordStubs("{датан}", dataGridView1.CurrentRow.Cells[9].Value.ToString().Substring(0, 10), document);
//                                ReplaceWordStubs("{датак}", dataGridView1.CurrentRow.Cells[10].Value.ToString().Substring(0, 10), document);
//                                ReplaceWordStubs("{страховщик}", dataGridView1.CurrentRow.Cells[2].Value.ToString(), document);

//                                ReplaceWordStubs("{страхователь}", dataGridView1.CurrentRow.Cells[4].Value.ToString(), document);
//                                if (dataGridView1.Rows.Count != 0)
//                                {
//                                    if (File.Exists($@"/Diplom/proga/Agent/treaty/live/live{id}.doc") == true)
//                                    {
                                      
//                                    }
//                                    else
//                                    {
//                                        try
//                                        {
//                                            document.SaveAs2($@"/Diplom/proga/Agent/treaty/live/live{id}.doc");
//                                            document.Close();
//                                            application.Quit();
//                                        }
//                                        catch { MessageBox.Show("Ошибка!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }
//                                    }
//                                }
//                            }
//                            else if (dataGridView1.CurrentRow.Cells[6].Value.ToString() == "Страхование гражданской ответственности за причинение вреда в связи с осуществлением профессиональной деятельности")
//                            {
//                                strax = "";
//                                strax2 = "";
//                                pas = "";
//                                vznos = "";
//                                Word.Application wordApplication = new Word.Application();

//                                string PathToNote = @"/Diplom/proga/Agent/treaty/prof/prof.doc";
//                                Word.Application application = new Word.Application();
//                                application.Visible = false;
//                                Word.Document document = application.Documents.Open(PathToNote);
//                                ReplaceWordStubs("{номер}", dataGridView1.CurrentRow.Cells[0].Value.ToString(), document);
//                                ReplaceWordStubs("{дата}", dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(0, 10), document);

//                                string query1 = $@"Select 
//position.name as [Должность]
//from policyholder inner join city on policyholder.idcity=city.idcity 
//inner join position on position.idposition=policyholder.idwork 
//inner join work on work.idwork=position.idwork
// where idpolicyholder={dataGridView1.CurrentRow.Cells[3].Value.ToString()}";
//                                DataTable data = new DataTable();
//                                SqlDataAdapter command1 = new SqlDataAdapter(query1, sqlConnection);
//                                command1.Fill(data);
//                                DataColumn column = data.Columns[0];
//                                DataRow row = data.Rows[0];
//                                ReplaceWordStubs("{должность}", row[column].ToString().ToLower(), document);
//                                ReplaceWordStubs("{страхователь}", dataGridView1.CurrentRow.Cells[4].Value.ToString(), document);
//                                ReplaceWordStubs("{датан}", dataGridView1.CurrentRow.Cells[9].Value.ToString().Substring(0, 10), document);
//                                ReplaceWordStubs("{датан}", dataGridView1.CurrentRow.Cells[9].Value.ToString().Substring(0, 10), document);
//                                ReplaceWordStubs("{датак}", dataGridView1.CurrentRow.Cells[10].Value.ToString().Substring(0, 10), document);
//                                ReplaceWordStubs("{взнос}", dataGridView1.CurrentRow.Cells[17].Value.ToString(), document);
//                                ReplaceWordStubs("{квзнос}", dataGridView1.CurrentRow.Cells[8].Value.ToString().ToLower(), document);

//                                if (dataGridView1.CurrentRow.Cells[18].Value.ToString() != "")
//                                {
//                                    ReplaceWordStubs("{датав}", dataGridView1.CurrentRow.Cells[18].Value.ToString().Substring(0, 10), document);
//                                }
//                                else if (dataGridView1.CurrentRow.Cells[20].Value.ToString() != "")
//                                {
//                                    ReplaceWordStubs("{датав}", dataGridView1.CurrentRow.Cells[18].Value.ToString().Substring(0, 10) + "; " + dataGridView1.CurrentRow.Cells[20].Value.ToString().Substring(0, 10) + "; " + dataGridView1.CurrentRow.Cells[22].Value.ToString().Substring(0, 10), document);

//                                }
//                                else
//                                {
//                                    ReplaceWordStubs("{датав}", dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(0, 10), document);
//                                }
//                                ReplaceWordStubs("{страхователь}", dataGridView1.CurrentRow.Cells[4].Value.ToString(), document);
//                                if (dataGridView1.Rows.Count != 0)
//                                {
//                                    if (File.Exists($@"/Diplom/proga/Agent/treaty/prof/prof{id}.doc") == true)
//                                    {
//                                    }
//                                    else
//                                    {
//                                        try
//                                        {
//                                            document.SaveAs2($@"/Diplom/proga/Agent/treaty/prof/prof{id}.doc");
//                                            document.Close();
//                                            application.Quit();
//                                        }
//                                        catch { MessageBox.Show("Ошибка!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }

//                                    }
//                                }
//                            }
//                            else if (dataGridView1.CurrentRow.Cells[6].Value.ToString() == "Страхование средств железнодорожного транспорта")
//                            {

//                                sqlConnection.Open();
//                                DataSet dataSet = new DataSet();
//                                SqlDataAdapter command = new SqlDataAdapter($@"
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
//                                command.Fill(dataSet);
//                                dataGridView2.DataSource = dataSet.Tables[0].DefaultView;
//                                sqlConnection.Close();
//                                dataGridView2.AllowUserToAddRows = false;
//                                sqlConnection.Close();
//                                dataGridView2.RowsDefaultCellStyle.WrapMode = DataGridViewTriState.True;
//                                dataGridView2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
//                                dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;


//                                strax = "";
//                                strax2 = "";
//                                pas = "";
//                                vznos = "";
//                                Word.Application wordApplication = new Word.Application();

//                                string PathToNote = @"/Diplom/proga/Agent/treaty/train/train.docx";
//                                Word.Application application = new Word.Application();
//                                application.Visible = false;
//                                Word.Document document = application.Documents.Open(PathToNote);
//                                object oMissing = System.Reflection.Missing.Value;
//                                application.Selection.Find.Execute("%метка%");
//                                Word.Range wordRange = application.Selection.Range;

//                                int RowCount2 = dataGridView2.RowCount + 2;
//                                int ColumnCount2 = 7;
//                                Object defaultTableBehavior2 =
//                                       Word.WdDefaultTableBehavior.wdWord9TableBehavior;
//                                Object autoFitBehavior2 = Word.WdAutoFitBehavior.wdAutoFitWindow;
//                                Word.Table wordtable2 = document.Tables.Add(application.Selection.Range, RowCount2, ColumnCount2,
//                                                  ref defaultTableBehavior2, ref autoFitBehavior2);
//                                Object style2 = "Сетка таблицы";
//                                wordtable2.set_Style(ref style2);
//                                wordtable2.ApplyStyleFirstColumn = true;
//                                wordtable2.ApplyStyleHeadingRows = true;
//                                Word.Range wordcellrange10 = document.Tables[1].Cell(1, 1).Range;
//                                wordcellrange10.Text = "№ п/п";
//                                wordcellrange10.Font.Size = 10;

//                                Word.Range wordcellrange11 = document.Tables[1].Cell(1, 2).Range;
//                                wordcellrange11.Text = "Наименование средств транспорта, номера";
//                                wordcellrange11.Font.Size = 10;

//                                Word.Range wordcellrange12 = document.Tables[1].Cell(1, 3).Range;
//                                wordcellrange12.Text = "Тип транспорта, сертификат, год постройки";
//                                wordcellrange12.Font.Size = 10;

//                                Word.Range wordcellrange13 = document.Tables[1].Cell(1, 4).Range;
//                                wordcellrange13.Text = "Страховая стоимость (руб.)";
//                                wordcellrange13.Font.Size = 10;

//                                Word.Range wordcellrange14 = document.Tables[1].Cell(1, 5).Range;
//                                wordcellrange14.Text = "Страховая сумма (руб.)";
//                                wordcellrange14.Font.Size = 10;

//                                Word.Range wordcellrange15 = document.Tables[1].Cell(1, 6).Range;
//                                wordcellrange15.Text = "Тариф (%)";
//                                wordcellrange15.Font.Size = 10;

//                                Word.Range wordcellrange16 = document.Tables[1].Cell(1, 7).Range;
//                                wordcellrange16.Text = "Страховая премия (руб.)";
//                                wordcellrange16.Font.Size = 10;

//                                document.Tables[1].Rows[RowCount2].Cells[1].Merge(document.Tables[1].Rows[RowCount2].Cells[3]);
//                                document.Tables[1].Cell(RowCount2, 1).Range.Text = $@"Итого:";


                                
//                                decimal kprice = 0;
//                                decimal ksumm = 0;

//                                decimal kpre = 0;
//                                for (int m2 = 2; m2 < RowCount2; m2++)
//                                {
//                                    wordcellrange10 = wordtable2.Cell(m2, 1).Range;
//                                    wordcellrange10.Text = Convert.ToString(m2 - 1);
//                                    wordcellrange10.Font.Size = 10;
//                                    wordcellrange10 = wordtable2.Cell(m2, 2).Range;
//                                    wordcellrange10.Text = dataGridView2[0, m2 - 2].Value.ToString();
//                                    wordcellrange10.Font.Size = 10;
//                                    wordcellrange10 = wordtable2.Cell(m2, 3).Range;
//                                    wordcellrange10.Text = dataGridView2[1, m2 - 2].Value.ToString();
//                                    wordcellrange10.Font.Size = 10;
//                                    wordcellrange10 = wordtable2.Cell(m2, 4).Range;
//                                    wordcellrange10.Text = dataGridView2[2, m2 - 2].Value.ToString();
//                                    wordcellrange10.Font.Size = 10;
//                                    kprice = kprice + Convert.ToDecimal(dataGridView2[2, m2 - 2].Value.ToString());
//                                    wordcellrange10 = wordtable2.Cell(m2, 5).Range;
//                                    wordcellrange10.Text = dataGridView2[3, m2 - 2].Value.ToString();
//                                    wordcellrange10.Font.Size = 10;
//                                    ksumm = ksumm + Convert.ToDecimal(dataGridView2[3, m2 - 2].Value.ToString());
//                                    wordcellrange10 = wordtable2.Cell(m2, 6).Range;
//                                    wordcellrange10.Text = dataGridView2[4, m2 - 2].Value.ToString();
//                                    wordcellrange10.Font.Size = 10;
//                                    document.Tables[1].Cell(RowCount2, 4).Range.Text = Convert.ToString(dataGridView2[4, m2 - 2].Value.ToString());
//                                    wordcellrange10 = wordtable2.Cell(m2, 7).Range;
//                                    wordcellrange10.Text = dataGridView2[5, m2 - 2].Value.ToString();
//                                    wordcellrange10.Font.Size = 10;
//                                    kpre = kpre + Convert.ToDecimal(dataGridView2[5, m2 - 2].Value.ToString());

//                                }
//                                document.Tables[1].Cell(RowCount2, 2).Range.Text = Convert.ToString(kprice);
//                                document.Tables[1].Cell(RowCount2, 3).Range.Text = Convert.ToString(ksumm);
//                                document.Tables[1].Cell(RowCount2, 5).Range.Text = Convert.ToString(kpre);
//                                ReplaceWordStubs("{номер}", dataGridView1.CurrentRow.Cells[0].Value.ToString(), document);
//                                ReplaceWordStubs("{дата}", dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(0, 10), document);

//                                string query1 = $@"Select 
//position.name as [Должность]
//from policyholder inner join city on policyholder.idcity=city.idcity 
//inner join position on position.idposition=policyholder.idwork 
//inner join work on work.idwork=position.idwork
// where idpolicyholder={dataGridView1.CurrentRow.Cells[3].Value.ToString()}";
//                                DataTable data = new DataTable();
//                                SqlDataAdapter command1 = new SqlDataAdapter(query1, sqlConnection);
//                                command1.Fill(data);
//                                DataColumn column = data.Columns[0];
//                                DataRow row = data.Rows[0];
//                                ReplaceWordStubs("{страхователь}", dataGridView1.CurrentRow.Cells[4].Value.ToString(), document);
//                                ReplaceWordStubs("{страхс}", dataGridView1.CurrentRow.Cells[11].Value.ToString(), document);
//                                ReplaceWordStubs("{страхп}", dataGridView1.CurrentRow.Cells[15].Value.ToString(), document);

//                                if (dataGridView1.CurrentRow.Cells[20].Value.ToString() != "")
//                                {
//                                    ReplaceWordStubs("{датав}", dataGridView1.CurrentRow.Cells[17].Value.ToString() + " BYN " + dataGridView1.CurrentRow.Cells[18].Value.ToString().Substring(0, 10) + " г. ; " + dataGridView1.CurrentRow.Cells[17].Value.ToString() + " BYN " + dataGridView1.CurrentRow.Cells[20].Value.ToString().Substring(0, 10) + " г. ; " + dataGridView1.CurrentRow.Cells[17].Value.ToString() + " BYN " + dataGridView1.CurrentRow.Cells[22].Value.ToString().Substring(0, 10), document);

//                                }
//                                else if (dataGridView1.CurrentRow.Cells[18].Value.ToString() != "")
//                                {
//                                    ReplaceWordStubs("{датав}", dataGridView1.CurrentRow.Cells[17].Value.ToString() + " BYN " + dataGridView1.CurrentRow.Cells[18].Value.ToString().Substring(0, 10), document);
//                                }
//                                else
//                                {
//                                    ReplaceWordStubs("{датав}", dataGridView1.CurrentRow.Cells[17].Value.ToString() + " BYN " + dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(0, 10), document);
//                                }
//                                string query45 = $@"select date from treaty 
//inner join bid on bid.idbid=treaty.idbid where treaty.idbid={dataGridView1.CurrentRow.Cells[24].Value.ToString()}";
//                                DataTable data45 = new DataTable();
//                                SqlDataAdapter command45 = new SqlDataAdapter(query45, sqlConnection);
//                                command45.Fill(data45);
//                                DataColumn column45 = data45.Columns[0];
//                                DataRow row45 = data45.Rows[0];

//                                ReplaceWordStubs("{датаз}", row45[column45].ToString().Substring(0, 10), document);
//                                ReplaceWordStubs("{номерз}", dataGridView1.CurrentRow.Cells[24].Value.ToString(), document);
//                                ReplaceWordStubs("{дата}", dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(0, 10), document);
//                                ReplaceWordStubs("{дата}", dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(0, 10), document);
//                                ReplaceWordStubs("{дата}", dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(0, 10), document);
//                                ReplaceWordStubs("{страхователь}", dataGridView1.CurrentRow.Cells[4].Value.ToString(), document);
//                                ReplaceWordStubs("{страховщик}", dataGridView1.CurrentRow.Cells[2].Value.ToString(), document);
//                                if (dataGridView1.Rows.Count != 0)
//                                {
//                                    if (File.Exists($@"/Diplom/proga/Agent/treaty/train/train{id}.doc") == true)
//                                    {
//                                    }
//                                    else
//                                    {
//                                        try
//                                        {
//                                            document.SaveAs2($@"/Diplom/proga/Agent/treaty/train/train{id}.doc");
//                                            document.Close();
//                                            application.Quit();
//                                        }
//                                        catch { MessageBox.Show("Ошибка!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }

//                                    }
//                                }
//                            }
//                            else
//                            {
//                                MessageBox.Show("Нет шаблона договора!", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);

//                            }

                        
                    }}
                        else
                        {
                            checkBox1.Checked = false;
                            MessageBox.Show("Выберите договор!", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }
                }
                else
                {

                }
            }catch
            {
                MessageBox.Show("Что-то не так!", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            id = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
        }
        string dos = "";
        private void button5_Click(object sender, EventArgs e)
        {
            //Фрагмент кода отправления письма 
            Regex r2 = new Regex(@"(\S*(@mail\.ru|@gmail\.com|@yandex\.ru))$");
            if (r2.IsMatch(textBox10.Text) && textBox10.TextLength>8)
            {
                try
                {
                    if (insurerak!=null)
                    {
                        string query2 = $@"select (firstname+' '+name+' '+lastname) from worker where idworker={insurerak.idakk}";
                        System.Data.DataTable data2 = new System.Data.DataTable();
                        SqlDataAdapter command2 = new SqlDataAdapter(query2, sqlConnection);
                        command2.Fill(data2);
                            DataColumn column2 = data2.Columns[0];
                            DataRow row2 = data2.Rows[0];
                            dos =(row2[column2].ToString());   
                    }
                    else
                    {
                        dos = "Admin";
                    }
                    MailAddress fromadress = new MailAddress("mih2023@mail.ru",dos );
                    MailAddress toadress = new MailAddress(textBox10.Text, comboBox2.Text);
                    MailMessage Message = new MailMessage(fromadress, toadress);
                    if (checkBox1.Checked == true)
                    {
                        //if (dataGridView1.CurrentRow.Cells[6].Value.ToString() == "Медицинское страхование")
                        //{
                        //    Message.Attachments.Add(new Attachment($@"/Diplom/proga/Agent/treaty/med/med{id}.doc"));
                        //}
                        //else if (dataGridView1.CurrentRow.Cells[6].Value.ToString() == "Страхование от несчатного случая")
                        //{
                        //    Message.Attachments.Add(new Attachment($@"/Diplom/proga/Agent/treaty/nes/nes{id}.doc"));
                        //}
                        //else if (dataGridView1.CurrentRow.Cells[6].Value.ToString() == "Обязательное страхование гражданской ответственности перевозчика перед пассажирами")
                        //{
                        //    Message.Attachments.Add(new Attachment($@"/Diplom/proga/Agent/treaty/pere/pere{id}.doc"));
                        //}
                        //else if (dataGridView1.CurrentRow.Cells[6].Value.ToString() == "Страхование от несчастных случаев пассажиров железного транспорта")
                        //{
                        //    Message.Attachments.Add(new Attachment($@"/Diplom/proga/Agent/treaty/pas/pas{id}.doc"));
                        //}
                        //else if (dataGridView1.CurrentRow.Cells[6].Value.ToString() == "Страхование жизни")
                        //{
                        //    Message.Attachments.Add(new Attachment($@"/Diplom/proga/Agent/treaty/live/live{id}.doc"));
                        //}
                        //else if (dataGridView1.CurrentRow.Cells[6].Value.ToString() == "Страхование гражданской ответственности за причинение вреда в связи с осуществлением профессиональной деятельности")
                        //{
                        //    Message.Attachments.Add(new Attachment($@"/Diplom/proga/Agent/treaty/prof/prof{id}.doc"));
                        //}
                        //else if (dataGridView1.CurrentRow.Cells[6].Value.ToString() == "Страхование средств железнодорожного транспорта")
                        //{
                        //    Message.Attachments.Add(new Attachment($@"/Diplom/proga/Agent/treaty/train/train{id}.doc"));
                        //}
                    }
                    else { }
                    Message.Subject = "Isurance";
                    Message.Body = textBox2.Text;
                    SmtpClient smtpClient = new SmtpClient();
                    smtpClient.Host = "smtp.mail.ru";
                    smtpClient.Port = 587;
                    smtpClient.EnableSsl = true;
                    smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
                    smtpClient.UseDefaultCredentials = false;
                    smtpClient.Credentials = new NetworkCredential("mih2023@mail.ru", "P9faG8acG4uxy89LnRDK");
                    smtpClient.Send(Message);
                    MessageBox.Show("Письмо отправлено", "Отправлено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    checkBox1.Checked = false;
                    textBox2.Text = "";

                }
                catch { MessageBox.Show("Ошибка!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            }
            else
            {
                MessageBox.Show("Некорректная почта!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
    }
    }

