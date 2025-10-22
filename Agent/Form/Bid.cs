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
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Excel = Microsoft.Office.Interop.Excel;

namespace Agent.Form
{
    public partial class Bid : UserControl
    {    datebase db = new datebase();
        public SqlConnection sqlConnection = null;
        Tenantak policyholder;
        public Bid(Tenantak policyholder1)
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
        Workerak insurerak;
        public Bid(Workerak insurerak1)
        {
            InitializeComponent();
            sqlConnection = new SqlConnection(db.connection);
            insurerak = insurerak1;
        }
        int id = 0;int id2 = 0;int idpost = 0;
        public void Bid_load()
        {
            if (policyholder != null)
            {
                sqlConnection.Open();
                DataSet dataSet = new DataSet();
                SqlDataAdapter command = new SqlDataAdapter($@"Select idbid as [Номер заявки], bid.idtenant,(tenant.firstname +' '+tenant.name+' '+tenant.lastname) as Житель,
bid.idworker,(worker.firstname +' '+worker.name+' '+worker.lastname) as Ответственный, bid.idservice, (service.name+'- '+description+', '+ CONVERT(NVARCHAR(10),cost)) as Услуга, 
bid.idobject, (type+', '+address) as Объект, poz as Комментарий, status as Статус, 
date as [Дата заявки], datec as [Дата выполнения] 
from bid inner join object on bid.idobject=object.idobject inner join tenant on bid.idtenant=tenant.idtenant 
inner join worker on bid.idworker=worker.idworker inner join service on bid.idservice=service.idservice
where  bid.idtenant={policyholder.idakk}
union
Select idbid as [Номер заявки], bid.idtenant,(tenant.firstname +' '+tenant.name+' '+tenant.lastname) as Житель,
bid.idworker,NULL as Ответственный, bid.idservice, (service.name+'- '+description+', '+ CONVERT(NVARCHAR(10),cost)) as Услуга, 
bid.idobject, (type+', '+address) as Объект, poz as Комментарий, status as Статус, 
date as [Дата заявки], datec as [Дата выполнения] 
from bid inner join object on bid.idobject=object.idobject inner join tenant on bid.idtenant=tenant.idtenant 
 inner join service on bid.idservice=service.idservice
 where bid.idworker is NULL and bid.idtenant={policyholder.idakk}", sqlConnection);
                command.Fill(dataSet);
                dataGridView1.DataSource = dataSet.Tables[0].DefaultView;
                sqlConnection.Close();
                dataGridView1.Columns[1].Visible = false;
                dataGridView1.Columns[3].Visible = false;
                dataGridView1.Columns[5].Visible = false;
                dataGridView1.Columns[7].Visible = false;
                dataGridView1.Columns[0].Width = 100;
                dataGridView1.AllowUserToAddRows = false;
                sqlConnection.Close();
                button6.Visible = false;
                button7.Visible = false;
                comboBoxcity3();
                comboBoxcity2();
            }
            else if (insurerak != null)
            {
                string query1 = $@"Select idpost from worker where idworker='{insurerak.idakk}'";
                DataTable data = new DataTable();
                SqlDataAdapter command1 = new SqlDataAdapter(query1, sqlConnection);
                command1.Fill(data);
                DataColumn column = data.Columns[0];
                DataRow row = data.Rows[0];
                idpost = Convert.ToInt32(row[column].ToString());
                sqlConnection.Open();
                DataSet dataSet = new DataSet();
                SqlDataAdapter command = new SqlDataAdapter($@"Select idbid as [Номер заявки], bid.idtenant,(tenant.firstname +' '+tenant.name+' '+tenant.lastname) as Житель,
bid.idworker,(worker.firstname +' '+worker.name+' '+worker.lastname) as Ответственный, bid.idservice, (service.name+'- '+description+', '+ CONVERT(NVARCHAR(10),cost)) as Услуга, 
bid.idobject, (type+', '+address) as Объект, poz as Комментарий, status as Статус, 
date as [Дата заявки], datec as [Дата выполнения] 
from bid inner join object on bid.idobject=object.idobject inner join tenant on bid.idtenant=tenant.idtenant 
inner join worker on bid.idworker=worker.idworker inner join service on bid.idservice=service.idservice
where  bid.idworker={insurerak.idakk}
union
SELECT 
    bid.idbid AS [Номер заявки], 
    bid.idtenant, 
    (tenant.firstname + ' ' + tenant.name + ' ' + tenant.lastname) AS Житель,
    bid.idworker, 
    NULL AS Ответственный, 
    bid.idservice, 
    (service.name + '- ' + description + ', ' + CONVERT(NVARCHAR(10), cost)) AS Услуга, 
    bid.idobject, 
    (type + ', ' + address) AS Объект, 
    poz AS Комментарий, 
    status AS Статус, 
    date AS [Дата заявки], 
    datec AS [Дата выполнения] 
FROM 
    bid 
INNER JOIN 
    object ON bid.idobject = object.idobject 
INNER JOIN 
    tenant ON bid.idtenant = tenant.idtenant 
INNER JOIN 
    service ON bid.idservice = service.idservice 
INNER JOIN 
    worker ON worker.idpost = service.idpost 
WHERE 
    bid.idworker IS NULL
       AND worker.idpost = {idpost};", sqlConnection);
                command.Fill(dataSet);
                dataGridView1.DataSource = dataSet.Tables[0].DefaultView;
                sqlConnection.Close();
                dataGridView1.Columns[1].Visible = false;
                dataGridView1.Columns[3].Visible = false;
                dataGridView1.Columns[5].Visible = false;
                dataGridView1.Columns[7].Visible = false;
                dataGridView1.Columns[0].Width = 100;
                dataGridView1.AllowUserToAddRows = false;
                sqlConnection.Close();
            }
            else
            {
                sqlConnection.Open();
                DataSet dataSet = new DataSet();
                SqlDataAdapter command = new SqlDataAdapter($@"Select idbid as [Номер заявки], bid.idtenant,(tenant.firstname +' '+tenant.name+' '+tenant.lastname) as Житель,
bid.idworker,(worker.firstname +' '+worker.name+' '+worker.lastname) as Ответственный, bid.idservice, (service.name+'- '+description+', '+ CONVERT(NVARCHAR(10),cost)) as Услуга, 
bid.idobject, (type+', '+address) as Объект, poz as Комментарий, status as Статус, 
date as [Дата заявки], datec as [Дата выполнения] 
from bid inner join object on bid.idobject=object.idobject inner join tenant on bid.idtenant=tenant.idtenant 
inner join worker on bid.idworker=worker.idworker inner join service on bid.idservice=service.idservice
union
Select idbid as [Номер заявки], bid.idtenant,(tenant.firstname +' '+tenant.name+' '+tenant.lastname) as Житель,
bid.idworker,NULL as Ответственный, bid.idservice, (service.name+'- '+description+', '+ CONVERT(NVARCHAR(10),cost)) as Услуга, 
bid.idobject, (type+', '+address) as Объект, poz as Комментарий, status as Статус, 
date as [Дата заявки], datec as [Дата выполнения] 
from bid inner join object on bid.idobject=object.idobject inner join tenant on bid.idtenant=tenant.idtenant 
 inner join service on bid.idservice=service.idservice
 where bid.idworker is NULL", sqlConnection);
                command.Fill(dataSet);
                dataGridView1.DataSource = dataSet.Tables[0].DefaultView;
                sqlConnection.Close();
                dataGridView1.Columns[1].Visible = false;
                dataGridView1.Columns[3].Visible = false;
                dataGridView1.Columns[5].Visible = false;
                dataGridView1.Columns[7].Visible = false;
                dataGridView1.Columns[0].Width = 100;
                dataGridView1.AllowUserToAddRows = false;
                sqlConnection.Close();
            }
        }
      

        public void comboBoxcity3()
        {
            sqlConnection.Close();
            sqlConnection.Open();
            string query = $@"select idservice,name as p from service";
            SqlDataAdapter sqlDbDataAdapter = new SqlDataAdapter(query, sqlConnection);
            DataSet dataSet = new DataSet();
            sqlDbDataAdapter.Fill(dataSet);
            comboBox4.DataSource = dataSet.Tables[0];
            comboBox4.DisplayMember = "p";
            comboBox4.ValueMember = "idservice";
            comboBox4.SelectedIndex = -1;
            sqlConnection.Close();

        }
        public void objecti()
        {
            sqlConnection.Close();
            sqlConnection.Open();
            DataSet dataSet = new DataSet();
            SqlDataAdapter command = new SqlDataAdapter($@"Select idobject, type as [Тип объекта],address as Адрес, square as Площадь, com as Описание from object", sqlConnection);
            command.Fill(dataSet);
            dataGridView2.DataSource = dataSet.Tables[0].DefaultView;
            sqlConnection.Close();
            dataGridView2.Columns[0].Visible = false;
            dataGridView2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView2.AllowUserToAddRows = false;
            dataGridView2.Columns[1].Width = 120;
            dataGridView2.Columns[2].Width = 100;
            dataGridView2.Columns[3].Width = 80;
            dataGridView2.Columns[4].Width = 500;
            sqlConnection.Close();
        }
        public void comboBoxcity2()
        {if (policyholder != null)
            {
                sqlConnection.Close();
                sqlConnection.Open();
                string query = $@"select idtenant,(firstname+' '+name+' '+lastname) as p from tenant where idtenant={policyholder.idakk}";
                SqlDataAdapter sqlDbDataAdapter = new SqlDataAdapter(query, sqlConnection);
                DataSet dataSet = new DataSet();
                sqlDbDataAdapter.Fill(dataSet);
                comboBox1.DataSource = dataSet.Tables[0];
                comboBox1.DisplayMember = "p";
                comboBox1.ValueMember = "idtenant";
                comboBox1.SelectedIndex = -1;
                sqlConnection.Close();
                comboBox1.Enabled = false;
            }
            else
            {
                sqlConnection.Close();
                sqlConnection.Open();
                string query = "select idtenant,(firstname+' '+name+' '+lastname) as p from tenant";
                SqlDataAdapter sqlDbDataAdapter = new SqlDataAdapter(query, sqlConnection);
                DataSet dataSet = new DataSet();
                sqlDbDataAdapter.Fill(dataSet);
                comboBox1.DataSource = dataSet.Tables[0];
                comboBox1.DisplayMember = "p";
                comboBox1.ValueMember = "idtenant";
                comboBox1.SelectedIndex = -1;
                sqlConnection.Close();
                comboBox1.Enabled = true;
            }

        }
        public void clear()
        {
            textBox2.Text = "";
            comboBox1.SelectedIndex = -1;
            comboBox4.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;
            comboBox1.Enabled = true;
           
        }
        private void objectpolicyholder_Load(object sender, EventArgs e)
        {
            if(insurerak!=null) { 
                button1.Visible = true;
                button2.Visible = false;
                button6.Visible = false;
                button4.Visible = false;
                
                button7.Visible = false;
                insurerak.button6.Visible = true;
                insurerak.button10.Visible = true;
                insurerak.button9.Visible = true;
                
            }
            else
            { 
            //comboBoxcity2();
            //comboBoxcity();
               
                button1.Visible = true;
                button2.Visible = true;
                button6.Visible = true;
                button7.Visible = true;
            }
            dataGridView2.RowsDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridView1.RowsDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            dateTimePicker1.MaxDate= DateTime.Today;
            dateTimePicker1.MinDate= DateTime.Today;
            panel3.Visible = false;
            panel2.Visible = false;
            Bid_load();
            
        }

        private void button4_Click(object sender, EventArgs e)
        {
        
            if (panel2.Visible == false)
            {objecti();
                comboBoxcity3();
                comboBoxcity2();
               
                button7.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (48).png");
                clear();
                panel3.Visible = false;
                panel2.Visible = true;
                label6.Text = "Добавить заявку";
                button11.Text = "Добавить";
          comboBox2.Visible = false;
                label9.Visible = false;
                dateTimePicker2.Visible = false;
                
                Bid_load();
                dataGridView1.Enabled = true;
            button11.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (31).png");
                if (policyholder != null)
                { comboBox1.SelectedIndex = 0; }
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
         
            //try
            //{
                if (panel2.Visible == false)
                {
               
                    button7.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (48).png");
                panel3.Visible = false;
                clear();
                if (id != 0||policyholder!=null)
                    {
                        comboBoxcity3();
                        comboBoxcity2();
                        if (dataGridView1.CurrentRow.Cells[10].Value.ToString() != "Выполнено")
                        {
                            dataGridView1.Enabled = false;
                            panel2.Visible = true;
                            label6.Text = "Редактировать заявку";
                            comboBox1.SelectedValue = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                            comboBox1.Enabled = false;
                            comboBox4.SelectedValue = dataGridView1.CurrentRow.Cells[5].Value.ToString();
                            if (dataGridView1.CurrentRow.Cells[3].Value.ToString()!="")
                            {
                            label5.Visible = true;
                            comboBox2.Visible = true;
                                dateTimePicker2.Visible= true;
                                label9.Visible = true;
                            textBox2.Enabled = false;
                            comboBox2.SelectedValue = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                                dateTimePicker2.Value = Convert.ToDateTime(dataGridView1.CurrentRow.Cells[12].Value.ToString());
                            }
                            else
                            {
                                label9.Visible = false;
                                comboBox2.Visible = false;
                                dateTimePicker2.Visible = false;
                            textBox2.Enabled = false;
                        }
                        dateTimePicker1.MaxDate = DateTime.Today;
                        dateTimePicker1.MinDate = DateTime.Today.AddDays(-30);
                        dateTimePicker1.Value = Convert.ToDateTime(dataGridView1.CurrentRow.Cells[11].Value.ToString());
                        dateTimePicker1.Enabled = false;
                            textBox2.Text= dataGridView1.CurrentRow.Cells[9].Value.ToString();
                            if (policyholder!=null)
                            {
                                label5.Visible = false;
                                label9.Visible= false;
                                comboBox2.Visible = false;
                                dateTimePicker2.Visible = false;
                            textBox2.Enabled = true;

                            }
                            

                            sqlConnection.Close();
                            sqlConnection.Open();
                            DataSet dataSet = new DataSet();
                            SqlDataAdapter command = new SqlDataAdapter($@"Select idobject, type as [Тип объекта],address as Адрес, square as Площадь, com as Описание from object where idobject='{Convert.ToInt32( dataGridView1.CurrentRow.Cells[7].Value.ToString())}'", sqlConnection);
                            command.Fill(dataSet);
                            dataGridView2.DataSource = dataSet.Tables[0].DefaultView;
                            sqlConnection.Close();
                            dataGridView2.Columns[0].Visible = false;
                            dataGridView2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                            dataGridView2.AllowUserToAddRows = false;
                            dataGridView2.Columns[1].Width = 120;
                            dataGridView2.Columns[2].Width = 100;
                            dataGridView2.Columns[3].Width = 80;
                            dataGridView2.Columns[4].Width = 500;
                            sqlConnection.Close();

                          
                            button11.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (29).png");
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
                        dataGridView1.Enabled = true;
                    }
                }
                else
                {
                    dataGridView1.Enabled = true;
                    clear();
                    panel2.Visible = false;
                }
            //}
            //catch { }
        }

        private void button2_Click(object sender, EventArgs e)
        {
           
            button7.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (48).png");
            clear();
            dataGridView1.Enabled = true;
            panel2.Visible = false;
            panel3.Visible = false;
            if (id != 0)
            {
                try
                {
                    if (MessageBox.Show($@"Вы уверены что хотите удалить заявку?", "Удаление", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        if (dataGridView1.CurrentRow.Cells[10].Value.ToString() == "Обрабатывается")
                        {
                            sqlConnection.Open();
                            string query = $@"DELETE FROM [bid] WHERE [idbid] = 
                             {dataGridView1.CurrentRow.Cells[0].Value.ToString()}";
                            SqlCommand command = new SqlCommand(query, sqlConnection);
                            command.ExecuteNonQuery();
                            sqlConnection.Close(); Bid_load();
                            
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
        int k = 0; int j = 0;
        private void button11_Click(object sender, EventArgs e)
        {
            //try
            //{
            if (label6.Text == "Добавить заявку")
            {
                
                //Код на добавление информации в БД об заявках страхователей
                if (dateTimePicker2.Visible==true) {
                    if (comboBox1.SelectedIndex != -1 && id2 != 0 &&comboBox4.SelectedIndex!=-1 && comboBox2.SelectedIndex!=-1 )
                    {
                        k = 0;
                        for (int i = 0; i < dataGridView1.RowCount; i++)
                        {
                            if (dataGridView1[10, i].Value.ToString()=="Принят"|| dataGridView1[10, i].Value.ToString() == "Обрабатывается") {
                                if (comboBox4.SelectedValue + Convert.ToString(id2) == dataGridView1[5, i].Value.ToString() + dataGridView1[7, i].Value.ToString())
                                {
                                    k++;

                                }
                            }
                        }
                        if (k == 0 )
                        {
                            k = 0; 
                            sqlConnection.Open();
                    SqlCommand command = new SqlCommand($@"INSERT INTO [bid](idtenant,idworker,idservice,idobject,status,date,datec,poz)
                            VALUES (@idt,@idw,@ids,@ido,@s,@date,@datec,@poz);", sqlConnection);
                    command.Parameters.AddWithValue("@idt", comboBox1.SelectedValue);
                    command.Parameters.AddWithValue("@idw", comboBox2.SelectedValue);
                    command.Parameters.AddWithValue("@ids", comboBox4.SelectedValue);
                    command.Parameters.AddWithValue("@ido", id2);
                    command.Parameters.AddWithValue("@s", ("Принят"));
                    command.Parameters.AddWithValue("@date", (dateTimePicker1.Value));
                    command.Parameters.AddWithValue("@datec", (dateTimePicker2.Value));
                    command.Parameters.AddWithValue("@poz", textBox2.Text);
                    command.ExecuteNonQuery();
                    sqlConnection.Close();
                           
                            Bid_load();
                            clear();
                            panel2.Visible = false;
                        }
                        else
                        {
                            MessageBox.Show("Такая услуга уже заказана есть!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
            else
            {
                MessageBox.Show("Заполните поля!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
                }
                else

                {
                    if (comboBox1.SelectedIndex != -1 && id2 != 0 && comboBox4.SelectedIndex !=-1)
                    {
                        k = 0; j = 0;
                        for (int i = 0; i < dataGridView1.RowCount; i++)
                        {
                            if (dataGridView1[10, i].Value.ToString() == "Принят" || dataGridView1[10, i].Value.ToString() == "Обрабатывается") { 
                                if (comboBox4.SelectedValue + Convert.ToString(id2) == dataGridView1[5, i].Value.ToString() + dataGridView1[7, i].Value.ToString())
                                {
                                    k++;
                                    j = Convert.ToInt32(dataGridView1[0, i].Value.ToString());
                                }
                            }
                        }
                        if (k == 0 || j == id)
                        {
                            k = 0; j = 0;
                            sqlConnection.Open();
                    SqlCommand command = new SqlCommand($@"INSERT INTO [bid](idtenant,idservice,idobject,status,date,poz)
                            VALUES (@idt,@ids,@ido,@s,@date,@poz);", sqlConnection);
                    command.Parameters.AddWithValue("@idt", comboBox1.SelectedValue);
                    command.Parameters.AddWithValue("@ids", comboBox4.SelectedValue);
                    command.Parameters.AddWithValue("@ido", id2);
                    command.Parameters.AddWithValue("@s", ("Обрабатывается"));
                    command.Parameters.AddWithValue("@date", (dateTimePicker1.Value));
                    command.Parameters.AddWithValue("@poz", textBox2.Text);
                    command.ExecuteNonQuery();
                    sqlConnection.Close();
                            
                            Bid_load();
                            clear();
                            panel2.Visible = false;
                        }
                        else
                        {
                            MessageBox.Show("Такая услуга уже заказана есть!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Заполните поля!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }





           
        }
                else
                {
                    
                    if (dateTimePicker2.Visible == true)
                    {if (comboBox2.SelectedIndex != -1 && comboBox4.SelectedIndex!=-1)
                    {
                        sqlConnection.Open();
                        SqlCommand command = new SqlCommand($@"UPDATE bid SET idworker=@idw," +
                        $"idservice=@ids,status=@s,datec=@d, poz=@p WHERE idbid=@id", sqlConnection);
                        command.Parameters.AddWithValue("@d", (dateTimePicker2.Value));
                        command.Parameters.AddWithValue("@idw", (comboBox2.SelectedValue));
                        command.Parameters.AddWithValue("@ids", (comboBox4.SelectedValue));
                        command.Parameters.AddWithValue("@s", "Принят");
                        command.Parameters.AddWithValue("@id", (id));
                        command.Parameters.AddWithValue("@p", (textBox2.Text));
                        command.ExecuteNonQuery();
                        sqlConnection.Close();
                        dataGridView1.Enabled = true;
                        clear();
                        Bid_load();
                        comboBox1.Enabled = true;
                        dateTimePicker1.Enabled = true;
                        panel2.Visible = false;
                        id2 = 0;
                    }else
                    {
                        MessageBox.Show("Заполните поля!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                }
                else
                {
                    if ( comboBox4.SelectedIndex != -1)
                    {
                        sqlConnection.Open();
                        SqlCommand command = new SqlCommand($@"UPDATE bid SET " +
                        $"idservice=@ids,date=@d, poz=@p  WHERE idbid=@id", sqlConnection);
                        command.Parameters.AddWithValue("@d", (dateTimePicker1.Value));
                      
                        command.Parameters.AddWithValue("@ids", (comboBox4.SelectedValue));
                        command.Parameters.AddWithValue("@p", (textBox2.Text));
                        command.Parameters.AddWithValue("@id", (id));
                        command.ExecuteNonQuery();
                        sqlConnection.Close();
                        dataGridView1.Enabled = true;
                        clear();
                        Bid_load();
                        comboBox1.Enabled = true;
                        dateTimePicker1.Enabled = true;
                        panel2.Visible = false;
                        id2 = 0;
                    }
                    else
                    {
                        MessageBox.Show("Заполните поля!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                    
                }
            //}
            //catch { }
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
                Bid_load();
                button7.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (48).png");
            }
            else
            {
                panel3.Visible = true;
                button7.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (47).png");
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



        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
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
            for (int i = 0; i < dataGridView1.ColumnCount; i++)
            {
                ExcelApp.Cells[2, i+1] = dataGridView1.Columns[i].HeaderText;
            }
            int y = 0;
            for (int j = 0; j < dataGridView1.ColumnCount; j++)
            {
                y = 0;
                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    if (dataGridView1.Rows[i].Cells[j].Visible == true)
                    {
                        if (j == 1)
                        {
                            ExcelApp.Cells[y + 3, j+1] = dataGridView1.Rows[i].Cells[j].Value.ToString().Substring(0,10);
                            y++;
                        }
                        else
                        {
                            ExcelApp.Cells[y+ 3, j+1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                            y++;
                        }
                        }
                }
            }
            for (int i = 0; i < visible; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    Microsoft.Office.Interop.Excel.Range range = ExcelWorkSheet.Range[$"A1:H{visible + 3}"];
                    ExcelWorkSheet.Range[$"A1:H{visible + 3}"].Cells.WrapText = true;
                    range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders.Weight = Excel.XlBorderWeight.xlMedium;
                }
            }
            ExcelApp.Columns["C"].Delete();
            ExcelApp.Columns["E"].Delete();
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            sqlConnection.Close();
            sqlConnection.Open();
            string query = $@"select post.name from post, worker,service where post.idpost=worker.idpost and service.idpost=post.idpost and service.name= '{comboBox4.Text}'";
            SqlCommand command = new SqlCommand(query, sqlConnection);
            string result = (string)command.ExecuteScalar();
            sqlConnection.Close();
            if (insurerak != null)
            {
                string query1 = $@"select worker.idworker,(firstname+' '+worker.name+' '+lastname) as p from worker, post where worker.idpost=post.idpost and  worker.idworker='{insurerak.idakk}'";
                SqlDataAdapter sqlDbDataAdapter = new SqlDataAdapter(query1, sqlConnection);
                DataSet dataSet = new DataSet();
                sqlDbDataAdapter.Fill(dataSet);
                comboBox2.DataSource = dataSet.Tables[0];
                comboBox2.DisplayMember = "p";
                comboBox2.ValueMember = "idworker";
                comboBox2.SelectedIndex = -1;
                sqlConnection.Close();
                comboBox2.Enabled= false;
               
            }
            else
            {
                string query1 = $@"select worker.idworker,(firstname+' '+worker.name+' '+lastname) as p from worker, post where worker.idpost=post.idpost and post.name='{result}'";
                SqlDataAdapter sqlDbDataAdapter = new SqlDataAdapter(query1, sqlConnection);
                DataSet dataSet = new DataSet();
                sqlDbDataAdapter.Fill(dataSet);
                comboBox2.DataSource = dataSet.Tables[0];
                comboBox2.DisplayMember = "p";
                comboBox2.ValueMember = "idworker";
                comboBox2.SelectedIndex = -1;
                sqlConnection.Close();
                comboBox2.Enabled = true;
            }
        
        }

        private void label9_Click(object sender, EventArgs e)
        {
            
        }

        private void label5_Click(object sender, EventArgs e)
        {
            if (comboBox2.Visible == true)
            {
                comboBox2.Visible = false;
                label9.Visible= false;
                    dateTimePicker2.Visible = false;
            }
            else
            { 
                    label9.Visible = true;
                dateTimePicker2.Visible = true;
                comboBox2.Visible = true; 
                if (insurerak != null)
                { comboBox2.SelectedIndex = 0; }
               
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            dateTimePicker1.Value = DateTime.Today;
        }

        private void dataGridView1_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if ((dataGridView1.CurrentRow.Cells[10].Value.ToString() != "Обрабатывается"))
            { 

                if (MessageBox.Show($@"Задача выполнена?", "Выполнение", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    if ((dataGridView1.CurrentRow.Cells[10].Value.ToString() == "Принят"))
                    {
                        if (Convert.ToDateTime(dataGridView1.CurrentRow.Cells[12].Value.ToString().Substring(0, 10)) == Convert.ToDateTime(Convert.ToString(DateTime.Now).ToString().Substring(0, 10)))
                        {
                            sqlConnection.Open();
                            SqlCommand command = new SqlCommand($@"UPDATE bid SET datec=@d ," +
                            $"status=@s WHERE idbid=@id", sqlConnection);
                            command.Parameters.AddWithValue("@s", ("Выполнено"));
                            command.Parameters.AddWithValue("@d", (dateTimePicker2.Value));
                            command.Parameters.AddWithValue("@id", (id));
                            command.ExecuteNonQuery();
                            sqlConnection.Close();
                            dataGridView1.Enabled = true;
                            clear();
                            Bid_load();
                        }
                        else if (Convert.ToDateTime(dataGridView1.CurrentRow.Cells[12].Value.ToString().Substring(0, 10)) < Convert.ToDateTime(Convert.ToString(DateTime.Now).ToString().Substring(0, 10)))

                        {
                            sqlConnection.Open();
                            SqlCommand command = new SqlCommand($@"UPDATE bid SET datec=@d ," +
                            $"status=@s WHERE idbid=@id", sqlConnection);
                            command.Parameters.AddWithValue("@s", ("Позже срока"));
                            command.Parameters.AddWithValue("@d", (dateTimePicker2.Value));
                            command.Parameters.AddWithValue("@id", (id));
                            command.ExecuteNonQuery();
                            sqlConnection.Close();
                            dataGridView1.Enabled = true;
                            clear();
                            Bid_load();

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
            else
            {
                MessageBox.Show("Оформите задачу на себя!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
    }
}
