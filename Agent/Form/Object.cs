using Microsoft.Office.Interop.Word;
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
using static System.Net.Mime.MediaTypeNames;
using Excel = Microsoft.Office.Interop.Excel;

namespace Agent.Form
{
    public partial class Object : UserControl
    {
        datebase db = new datebase();
        public SqlConnection sqlConnection = null;
        Tenantak policyholder;
        public Object(Tenantak policyholder1)
        {
            InitializeComponent();
            sqlConnection = new SqlConnection(db.connection);
            policyholder = policyholder1;
        }
        Workerak insurerak;
        public Object(Workerak insurerak1)
        {
            InitializeComponent();
            sqlConnection = new SqlConnection(db.connection);
            insurerak = insurerak1;
        }
        Form1 form1;
        public Object(Form1 form)
        {
            InitializeComponent();
            sqlConnection = new SqlConnection(db.connection);
            form1 = form;
        }
        int id = 0;
        public void Vid_load()
        {
            sqlConnection.Open();
            DataSet dataSet = new DataSet();
            SqlDataAdapter command = new SqlDataAdapter($@"Select idobject, type as [Тип объекта],address as Адрес, square as Площадь, com as Описание from object", sqlConnection);
            command.Fill(dataSet);
            dataGridView1.DataSource = dataSet.Tables[0].DefaultView;
            sqlConnection.Close();
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.Columns[4].Width = 1000;
            sqlConnection.Close();
        }
        public void Vid_load2()
        {
            //sqlConnection.Open();
            //DataSet dataSet = new DataSet();
            //SqlDataAdapter command = new SqlDataAdapter($@"Select idvida, vid.name as [Вид страхования],
            //note as Примечание from vid, policyholder,position,work
            //where policyholder.idwork=position.idposition and work.idwork=position.idwork and 
            //idpolicyholder={policyholder.idakk}
            //and ((vid.name NOT Like '%гражданской%' and vid.name NOT Like '%пассажиров%'  
            //and work.name NOT Like '%бжд%' ) 
            //or (work.name  Like '%бжд%' and vid.name NOT Like '%профессиональной%')
            //or (work.name  Like '%бжд%' and position.name  Like '%начальник%' )) ", sqlConnection);
            //command.Fill(dataSet);
            //dataGridView1.DataSource = dataSet.Tables[0].DefaultView;
            //sqlConnection.Close();
            //dataGridView1.Columns[0].Visible = false;
            //dataGridView1.AllowUserToAddRows = false;
            //sqlConnection.Close();
        }
        private void Vid_Load(object sender, EventArgs e)
        {
            if (policyholder != null)
            {
                try
                {
                    button4.Visible = false;
                    button1.Visible = false;
                    button5.Visible = false;
                    button2.Visible = false;
                }
                catch { }
            }
            else if (insurerak != null)
            {
                try
                {
                    button4.Visible = false;
                    button1.Visible = false;
                    button5.Visible = false;
                    button2.Visible = false;
                }
                catch { }
            }
            else
            {
                button4.Visible = true;
                button1.Visible = true;
                button5.Visible = true;
                button2.Visible = true;
                
            }   dataGridView1.RowsDefaultCellStyle.WrapMode = DataGridViewTriState.True;
                dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                Vid_load();
                panel2.Visible = false;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (panel6.Visible == false)
            {
                dataGridView1.Enabled = true;
                textBox2.Text = "";
                panel6.Visible = true;
            }
            else
            {
                dataGridView1.Enabled = true;
                textBox2.Text = "";
                panel6.Visible = false;
            }


            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //try
            //{

                if (panel2.Visible == false)
                {
                    textBox3.Text = "";
                    textBox2.Text = "";
                if (id != 0&& dataGridView1.CurrentRow.Cells[1].Value.ToString()== "Многоквартирный дом")
                {
                    panel2.Left = 117;
                 
                    panel2.Width = 1496;
                    dataGridView1.Enabled = false;
                    panel2.Visible = true;
                    label6.Text = "Редактировать многоквартирный дом";
                    comboBox1.Text = "";
                    comboBox1.Items.Clear();
                    comboBox1.Items.Add("Панельный");
                    comboBox1.Items.Add("Кирпичный");
                    comboBox1.Items.Add("Монолитный");
                    comboBox1.Items.Add("Комбинированный");
                    comboBox1.SelectedIndex = -1;
                    textBox7.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                    textBox8.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();

                      //  text.Contains(wordToCheck, StringComparison.OrdinalIgnoreCase

                        if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("Водоснабжение: централизованное;")) 
                        { 
                        checkBox1.Checked = true;
                        }
                        if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("Водоснабжение: автономное;"))
                        {
                            checkBox2.Checked = true;
                        }

                        if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("Отопление: централизованное;"))
                        {
                            checkBox8.Checked = true;
                        }
                        if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("Отопление: автономное;"))
                        {
                            checkBox7.Checked = true;
                        }

                        if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("Электричество: центральное;"))
                        {
                            checkBox4.Checked = true;
                        }
                        if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("Электричество: автономное;"))
                        {
                            checkBox3.Checked = true;
                        }

                        if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("Канализация: централизованная;"))
                        {
                            checkBox6.Checked = true;
                        }
                        if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("Канализация: автономная;"))
                        {
                            checkBox5.Checked = true;
                        }         
                  
                        if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("Газ: централизованный газопровод;"))
                        {
                            checkBox10.Checked = true;
                        }
                        if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("Газ: cниженный газ в баллонах;"))
                        {
                            checkBox9.Checked = true;
                        }



                        if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("Подвал;"))
                        {
                            
                            checkBox18.Checked = true;
                        }
                        if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("Чердак;"))
                        {
                            
                            checkBox17.Checked = true;
                        }
                        if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("Электросчетавая;"))
                        {
                            
                            checkBox16.Checked = true;
                        }
                        if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("Машинное отделение лифтов;"))
                        {
                            
                            checkBox19.Checked = true;
                        }
                        

                        if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("Детская площадка;"))
                        {
                            
                            checkBox24.Checked = true;
                        }
                        if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("Спортивная площадка;"))
                        {
                        
                            checkBox23.Checked = true;
                        }
                        if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("Зеленые зоны;"))
                        {
                            
                            checkBox21.Checked = true;
                        }
                        if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("Места для парковки;"))
                        {
                            
                            checkBox22.Checked = true;
                        }



                        if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("Пассажирский"))
                        {
                            
                        string stroca = Convert.ToString(dataGridView1.CurrentRow.Cells[4].Value.ToString().Substring(dataGridView1.CurrentRow.Cells[4].Value.ToString().IndexOf("Пассажирский: ") + 1));
                        int finish = Convert.ToInt32(stroca.IndexOf(';'))+ Convert.ToInt32(dataGridView1.CurrentRow.Cells[4].Value.ToString().IndexOf("Пассажирский: ") + 1);
                            textBox6.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString().Substring(Convert.ToInt32(dataGridView1.CurrentRow.Cells[4].Value.ToString().IndexOf("Пассажирский: ") + 14), finish - Convert.ToInt32( dataGridView1.CurrentRow.Cells[4].Value.ToString().IndexOf("Пассажирский: ") + 14));

                        }
                        if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("Грузопассажирский"))
                        {
                            
                           
                        string stroca = Convert.ToString(dataGridView1.CurrentRow.Cells[4].Value.ToString().Substring(dataGridView1.CurrentRow.Cells[4].Value.ToString().IndexOf("Грузопассажирский: ") + 1));
                        int finish = Convert.ToInt32(stroca.IndexOf(';')) + Convert.ToInt32(dataGridView1.CurrentRow.Cells[4].Value.ToString().IndexOf("Грузопассажирский: ") + 1);
                        textBox5.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString().Substring(Convert.ToInt32(dataGridView1.CurrentRow.Cells[4].Value.ToString().IndexOf("Грузопассажирский: ") + 19), finish - Convert.ToInt32(dataGridView1.CurrentRow.Cells[4].Value.ToString().IndexOf("Грузопассажирский: ") + 19));

                    }
                    if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("Пожарный"))
                        {
                        string stroca = Convert.ToString(dataGridView1.CurrentRow.Cells[4].Value.ToString().Substring(dataGridView1.CurrentRow.Cells[4].Value.ToString().IndexOf("Пожарный: ") + 1));
                        int finish = Convert.ToInt32(stroca.IndexOf(';')) + Convert.ToInt32(dataGridView1.CurrentRow.Cells[4].Value.ToString().IndexOf("Пожарный: ") + 1);
                        textBox4.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString().Substring(Convert.ToInt32(dataGridView1.CurrentRow.Cells[4].Value.ToString().IndexOf("Пожарный: ") + 10), finish - Convert.ToInt32(dataGridView1.CurrentRow.Cells[4].Value.ToString().IndexOf("Пожарный: ") + 10));

                    }



                    if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("этажей"))
                    {
                        

                        string stroca = Convert.ToString(dataGridView1.CurrentRow.Cells[4].Value.ToString().Substring(dataGridView1.CurrentRow.Cells[4].Value.ToString().IndexOf("этажей: ") + 1));
                        int finish = Convert.ToInt32(stroca.IndexOf(';')) + Convert.ToInt32(dataGridView1.CurrentRow.Cells[4].Value.ToString().IndexOf("этажей: ") + 1);
                        textBox3.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString().Substring(Convert.ToInt32(dataGridView1.CurrentRow.Cells[4].Value.ToString().IndexOf("этажей: ") + 8), finish - Convert.ToInt32(dataGridView1.CurrentRow.Cells[4].Value.ToString().IndexOf("этажей: ") + 8));

                    }
                    if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("квартир"))
                    {
                       
                        string stroca = Convert.ToString(dataGridView1.CurrentRow.Cells[4].Value.ToString().Substring(dataGridView1.CurrentRow.Cells[4].Value.ToString().IndexOf("квартир: ") + 1));
                        int finish = Convert.ToInt32(stroca.IndexOf(';')) + Convert.ToInt32(dataGridView1.CurrentRow.Cells[4].Value.ToString().IndexOf("квартир: ") + 1);
                        textBox2.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString().Substring(Convert.ToInt32(dataGridView1.CurrentRow.Cells[4].Value.ToString().IndexOf("квартир: ") + 9), finish - Convert.ToInt32(dataGridView1.CurrentRow.Cells[4].Value.ToString().IndexOf("квартир: ") + 9));

                    }


                    if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("строения"))
                    {
                       
                        string stroca = Convert.ToString(dataGridView1.CurrentRow.Cells[4].Value.ToString().Substring(dataGridView1.CurrentRow.Cells[4].Value.ToString().IndexOf("строения: ") + 1));
                        int finish = Convert.ToInt32(stroca.IndexOf(';')) + Convert.ToInt32(dataGridView1.CurrentRow.Cells[4].Value.ToString().IndexOf("строения: ") + 1);
                        comboBox1.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString().Substring(Convert.ToInt32(dataGridView1.CurrentRow.Cells[4].Value.ToString().IndexOf("строения: ") + 10), finish - Convert.ToInt32(dataGridView1.CurrentRow.Cells[4].Value.ToString().IndexOf("строения: ") + 10));

                    }

                    if (checkBox14.Checked == true)
                    {

                        lift += "Пожарный: " + textBox4.Text + "; ";

                    }
                    else
                    {
                        delete = "Пожарный: " + textBox4.Text + "; ";
                        lift = lift.Replace(delete, "");
                        lift = lift.Replace("  ", " ");

                    }
                    if (checkBox13.Checked == true)
                    {
                        lift += "Грузопассажирский: " + textBox5.Text + "; ";

                    }
                    else
                    {
                        delete = "Грузопассажирский: " + textBox5.Text + "; ";
                        lift = lift.Replace(delete, "");
                        lift = lift.Replace("  ", " ");
                    }
            
                    if (checkBox12.Checked == true)
                    {
                        lift += "Пассажирский: " + textBox6.Text + "; ";

                    }
                    else
                    {
                        delete = "Пассажирский: " + textBox6.Text + "; ";
                        lift = lift.Replace(delete, "");
                        lift = lift.Replace("  ", " ");
                    }

                    button11.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (29).png");
                    button11.Text = "Редактировать";
                    button11.Width = 244;
                    button11.Left = 223;
                }
                else if (id != 0 && dataGridView1.CurrentRow.Cells[1].Value.ToString() == "Частный дом (коттедж)")
                {
                    dataGridView1.Enabled = false;
                    panel2.Visible = true;
                    panel11.Visible = true;
                    label6.Text = "Редактировать частный дом (коттедж)";
                    panel2.Width = 973;
                    textBox7.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                    textBox8.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                    panel2.Left = 317;
                    panel11.Left = 317;
                    //  text.Contains(wordToCheck, StringComparison.OrdinalIgnoreCase
                    comboBox1.Visible = true;
                    comboBox1.Text = "";
                    comboBox1.Items.Clear();
                    comboBox1.Items.Add("одноэтажный");
                    comboBox1.Items.Add("двухэтажные");
                    comboBox1.Items.Add("многосекционные");
                    comboBox1.SelectedIndex = -1;
                    if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("Водоснабжение: централизованное;"))
                    {
                        checkBox1.Checked = true;
                    }
                    if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("Водоснабжение: автономное;"))
                    {
                        checkBox2.Checked = true;
                    }

                    if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("Отопление: централизованное;"))
                    {
                        checkBox8.Checked = true;
                    }
                    if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("Отопление: автономное;"))
                    {
                        checkBox7.Checked = true;
                    }

                    if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("Электричество: центральное;"))
                    {
                        checkBox4.Checked = true;
                    }
                    if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("Электричество: автономное;"))
                    {
                        checkBox3.Checked = true;
                    }

                    if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("Канализация: централизованная;"))
                    {
                        checkBox6.Checked = true;
                    }
                    if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("Канализация: автономная;"))
                    {
                        checkBox5.Checked = true;
                    }

                    if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("Газ: централизованный газопровод;"))
                    {
                        checkBox10.Checked = true;
                    }
                    if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("Газ: cниженный газ в баллонах;"))
                    {
                        checkBox9.Checked = true;
                    }


                    if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("Сад;"))
                    {

                        checkBox29.Checked = true;
                    }
                    if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("Хозяйственная постройка;"))
                    {

                        checkBox26.Checked = true;
                    }
                    if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("Газон;"))
                    {

                        checkBox28.Checked = true;
                    }

                    if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("строения"))
                    {

                        string stroca = Convert.ToString(dataGridView1.CurrentRow.Cells[4].Value.ToString().Substring(dataGridView1.CurrentRow.Cells[4].Value.ToString().IndexOf("строения: ") + 1));
                        int finish = Convert.ToInt32(stroca.IndexOf(';')) + Convert.ToInt32(dataGridView1.CurrentRow.Cells[4].Value.ToString().IndexOf("строения: ") + 1);
                        comboBox1.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString().Substring(Convert.ToInt32(dataGridView1.CurrentRow.Cells[4].Value.ToString().IndexOf("строения: ") + 10), finish - Convert.ToInt32(dataGridView1.CurrentRow.Cells[4].Value.ToString().IndexOf("строения: ") + 10));

                    }
                }
                
                    else if (id != 0 && dataGridView1.CurrentRow.Cells[1].Value.ToString() == "Объект благоустройства")
                {
                    panel2.Left = 317;
                    panel2.Width = 509;
                    panel12.Visible = true;
                    panel12.Left = 826;
                    label2.Visible = false;
                    label14.Visible = false;
                    textBox2.Visible = false;
                    textBox3.Visible = false;
                    comboBox2.SelectedIndex = -1;
                    button11.Top = 240;
                    panel6.Visible = false;
                    dataGridView1.Enabled = true;
                    panel2.Visible = true;
                    textBox2.Text = "";
                    button11.Text = "Редактировать";
                    button11.Width = 174;
                    label3.Text = "Тип объекта";
                    comboBox1.Text = "";
                    comboBox1.Items.Clear();
                    comboBox1.Items.Add("парк");
                    comboBox1.Items.Add("сквер");
                    comboBox1.Items.Add("зона отдыха");
                    comboBox1.SelectedIndex = -1;
                    dataGridView1.Enabled = false;
                    panel2.Visible = true;
                    label6.Text = "Редактировать объект благоустройства";

                    textBox7.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                    textBox8.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();

                    if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("Детская площадка;"))
                    {
                        checkBox31.Checked = true;
                    }
                    if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("Скамейки;"))
                    {
                        checkBox27.Checked = true;
                    }

                    if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("Спортивная площадка;"))
                    {
                        checkBox30.Checked = true;
                    }
                    if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("Фонари;"))
                    {
                        checkBox32.Checked = true;
                    }

                    if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("Зеленые насаждения;"))
                    {
                        checkBox25.Checked = true;
                    }
                    if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("Освещение;"))
                    {
                        checkBox15.Checked = true;
                    }

                    if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("Санитарные зоны;"))
                    {
                        checkBox20.Checked = true;
                    }
                    if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("Поливочные системы;"))
                    {
                        checkBox11.Checked = true;
                    }

                    if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("объекта"))
                    {

                        string stroca = Convert.ToString(dataGridView1.CurrentRow.Cells[4].Value.ToString().Substring(dataGridView1.CurrentRow.Cells[4].Value.ToString().IndexOf("объекта: ") + 1));
                        int finish = Convert.ToInt32(stroca.IndexOf(';')) + Convert.ToInt32(dataGridView1.CurrentRow.Cells[4].Value.ToString().IndexOf("объекта: ") + 1);
                        comboBox1.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString().Substring(Convert.ToInt32(dataGridView1.CurrentRow.Cells[4].Value.ToString().IndexOf("объекта: ") + 9), finish - Convert.ToInt32(dataGridView1.CurrentRow.Cells[4].Value.ToString().IndexOf("объекта: ") + 9));

                    }

                }

                else
                {
                    MessageBox.Show("Строка не выбрана!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    panel2.Visible = false;
                }
            }
          
        //}
        //    catch { }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView1.Enabled = true;
            textBox2.Text = "";
            panel2.Visible = false;
            if (id != 0)
            {
                try
                {
                    if (MessageBox.Show($@"Вы уверены что хотите удалить объект {dataGridView1.CurrentRow.Cells[1].Value.ToString()}?", "Удаление", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {

                        sqlConnection.Open();
                        string query = $@"DELETE FROM [object] WHERE [idobject] = 
                             {dataGridView1.CurrentRow.Cells[0].Value.ToString()}";
                        SqlCommand command = new SqlCommand(query, sqlConnection);
                        command.ExecuteNonQuery();
                        sqlConnection.Close();
                        Vid_load();
                    }
                }
                catch { MessageBox.Show("Удаление невозможно, данные все еще нужны!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            }
            else { MessageBox.Show("Выберите строку для удаления!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }

        }

        private void button5_Click(object sender, EventArgs e)
        {
            dataGridView1.Enabled = true;
            try
            {
                //Код вывода в «Excel» 
                Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
                Excel.Worksheet ExcelWorkSheet;
                ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
                ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
                ExcelApp.Columns.NumberFormat = "General";
                ExcelWorkSheet.StandardWidth = 30;
                ExcelWorkSheet.Columns.ColumnWidth = 20;
                ExcelApp.Rows[1].Columns[1] = "Вид страхования";
                ExcelApp.Rows[dataGridView1.RowCount + 3].Columns[1] = "Ридецкая Анна Михайловна";
                for (int i = 1; i < dataGridView1.Columns.Count; i++)
                {
                    ExcelApp.Cells[2, i] = dataGridView1.Columns[i].HeaderText;
                }
                for (int j = 1; j < dataGridView1.ColumnCount; j++)
                {
                    for (int i = 0; i < dataGridView1.RowCount; i++)
                    {
                        if (dataGridView1.Rows[i].Cells[j].Visible == true)
                        {   
                                    ExcelApp.Cells[i + 3, j] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                        }
                    }
                }
                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    for (int j = 0; j < dataGridView1.ColumnCount; j++)
                    {
                        Microsoft.Office.Interop.Excel.Range range = ExcelWorkSheet.Range[$"A1:D{dataGridView1.RowCount + 3}"];
                        ExcelWorkSheet.Range[$"A1:D{dataGridView1.RowCount + 3}"].Cells.WrapText = true;
                        range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders.Weight = Excel.XlBorderWeight.xlMedium;
                    }
                }
                ExcelApp.Visible = true;
                ExcelApp.UserControl = true;
            }
            catch { }
        }
        int k = 0;
        int j = 0;
        private void button11_Click(object sender, EventArgs e)
        {
            //try
            //{
                k = 0;
                if (label6.Text == "Добавить многоквартирный дом")
                {
                if (textBox2.Text != "" && textBox3.Text != ""&& textBox8.Text != "" && textBox7.Text != ""&& comboBox1.Text!="")
                {
                    k = 0; 
                    for (int i = 0; i < dataGridView1.RowCount; i++)
                    {
                        if ("многоквартирный дом" + textBox7.Text.ToLower() == dataGridView1[1, i].Value.ToString().ToLower() + dataGridView1[2, i].Value.ToString().ToLower())
                        {
                            k++;
                        }
                    }
                    if (k == 0 )
                    {
                        k = 0; 
                        //Код на добавление данных в БД о виде страхования
                        sqlConnection.Open();
                            SqlCommand command = new SqlCommand($@"INSERT INTO [object](type,address,square,com) 
                            VALUES (@t,@a,@s,@c);", sqlConnection);
                            command.Parameters.AddWithValue("@t", ("Многоквартирный дом"));
                            command.Parameters.AddWithValue("@a", (textBox7.Text));
                        command.Parameters.AddWithValue("@s", Convert.ToDecimal(textBox8.Text));
                        command.Parameters.AddWithValue("@c", ("- Количество этажей: "+textBox3.Text+"; "+
                            "Количество квартир: " + textBox2.Text + "; " +
                            "Тип строения: " + comboBox1.Text + "; \n" +
                            "- Комунникации: " +voda + otop  +" \n" + elec +canal +gaz + " \n" +
                            "- Наличие лифтов: " + lift + " \n" +
                            "- Технические помещения: " + tex + " \n" +
                            "- Придомовая территория: " + ter
                            ));
                        command.ExecuteNonQuery();
                            sqlConnection.Close();
                            Vid_load();
                            textBox2.Text = "";
                            textBox3.Text = "";
                clear();
                            panel2.Visible = false;

                }
                else
                {
                    MessageBox.Show("Такой многоквартирный дом уже есть!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
                else
                {
                    MessageBox.Show("Заполните поля!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else if(label6.Text == "Редактировать многоквартирный дом")
                {
                if (textBox2.Text != "" && textBox3.Text != "" && textBox8.Text != "" && textBox7.Text != "" && comboBox1.Text != "")
                {
                    k = 0; j = 0;
                    for (int i = 0; i < dataGridView1.RowCount; i++)
                    {
                        if ("многоквартирный дом" + textBox7.Text.ToLower() == dataGridView1[1, i].Value.ToString().ToLower() + dataGridView1[2, i].Value.ToString().ToLower())
                        {
                            k++;
                            j = Convert.ToInt32(dataGridView1[0, i].Value.ToString());
                        }
                    }
                    if (k == 0 || j == id)
                    {
                        k = 0; j = 0;
                        sqlConnection.Open();
                            SqlCommand command = new SqlCommand($@"UPDATE object SET address=@a, square=@s, com=@c WHERE idobject=@id", sqlConnection);
                            command.Parameters.AddWithValue("@a", (textBox7.Text));
                            command.Parameters.AddWithValue("@s", Convert.ToDecimal(textBox8.Text));
                        command.Parameters.AddWithValue("@c", "- Количество этажей: " + textBox3.Text + "; " +
                            "Количество квартир: " + textBox2.Text + "; " +
                            "Тип строения: " + comboBox1.Text + "; \n" +
                            "- Комунникации: " + voda + otop + " \n" + elec + canal + gaz + " \n" +
                            "- Наличие лифтов: " + lift + " \n" +
                            "- Технические помещения: " + tex + " \n" +
                            "- Придомовая территория: " + ter);
                        command.Parameters.AddWithValue("@id", (id));
                            command.ExecuteNonQuery();
                            sqlConnection.Close();
                            Vid_load();
                            textBox2.Text = "";
                            textBox3.Text = "";
                            panel2.Visible = false;
                            dataGridView1.Enabled = true;
                clear();
                }
                else
                {
                    MessageBox.Show("Такой многоквартирный дом уже есть!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
                else
                {
                    MessageBox.Show("Заполните поля!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else if (label6.Text == "Добавить частный дом (коттедж)")
            {
                if (textBox7.Text != "" && textBox8.Text != ""  && comboBox1.Text != "")
                {
                    k = 0;
                    for (int i = 0; i < dataGridView1.RowCount; i++)
                    {
                        if ("частный дом (коттедж)" + textBox7.Text.ToLower() == dataGridView1[1, i].Value.ToString().ToLower() + dataGridView1[2, i].Value.ToString().ToLower())
                        {
                            k++;
                          
                        }
                    }
                    if (k == 0 )
                    {
                        k = 0; 
                        //Код на добавление данных в БД о виде страхования
                        sqlConnection.Open();
                SqlCommand command = new SqlCommand($@"INSERT INTO [object](type,address,square,com) 
                            VALUES (@t,@a,@s,@c);", sqlConnection);
                command.Parameters.AddWithValue("@t", ("Частный дом (коттедж)"));
                command.Parameters.AddWithValue("@a", (textBox7.Text));
                command.Parameters.AddWithValue("@s", Convert.ToDecimal(textBox8.Text));
                command.Parameters.AddWithValue("@c", ("Тип строения: " + comboBox1.Text + "; \n" +
                    "- Комунникации: " + voda + otop + " \n" + elec + canal + gaz + " \n" +
                    "- Придомовая территория: " + ter
                    ));
                command.ExecuteNonQuery();
                sqlConnection.Close();
                Vid_load();
                textBox2.Text = "";
                textBox3.Text = "";
                clear();
                panel2.Visible = false;
                panel11.Visible = false;
                }
                else
                {
                    MessageBox.Show("Такой частный дом (коттедж) уже есть!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
                else
                {
                    MessageBox.Show("Заполните поля!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else if (label6.Text == "Редактировать частный дом (коттедж)")
            {
                if (textBox7.Text != "" && textBox8.Text != "" && comboBox1.Text != "")
                {
                    k = 0; j = 0;
                    for (int i = 0; i < dataGridView1.RowCount; i++)
                    {
                        if ("частный дом (коттедж)" + textBox7.Text.ToLower() == dataGridView1[1, i].Value.ToString().ToLower() + dataGridView1[2, i].Value.ToString().ToLower())
                        {
                            k++;
                            j = Convert.ToInt32(dataGridView1[0, i].Value.ToString());
                        }
                    }
                    if (k == 0 || j == id)
                    {
                        k = 0; j = 0;
                        sqlConnection.Open();
                SqlCommand command = new SqlCommand($@"UPDATE object SET address=@a, square=@s, com=@c WHERE idobject=@id", sqlConnection);
                command.Parameters.AddWithValue("@a", (textBox7.Text));
                command.Parameters.AddWithValue("@s", Convert.ToDecimal(textBox8.Text));
                command.Parameters.AddWithValue("@c",  "Тип строения: " + comboBox1.Text + "; \n" +
                    "- Комунникации: " + voda + otop + " \n" + elec + canal + gaz + " \n" +
                    "- Придомовая территория: " + ter);
                command.Parameters.AddWithValue("@id", (id));
                command.ExecuteNonQuery();
                sqlConnection.Close();
                Vid_load();
                textBox2.Text = "";
                textBox3.Text = "";
                panel2.Visible = false; 
                clear();
                panel11.Visible = false;
                dataGridView1.Enabled = true;
                panel12.Visible = false;

                }
                else
                {
                    MessageBox.Show("Такой частный дом (коттедж) уже есть!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
                else
                {
                    MessageBox.Show("Заполните поля!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else if (label6.Text == "Добавить объект благоустройства")
            {
                if (textBox7.Text != "" && textBox8.Text != "" && comboBox1.Text != "")
                {
                    k = 0; 
                    for (int i = 0; i < dataGridView1.RowCount; i++)
                    {
                        if ("объект благоустройства" + textBox7.Text.ToLower() == dataGridView1[1, i].Value.ToString().ToLower() + dataGridView1[2, i].Value.ToString().ToLower())
                        {
                            k++;
                           
                        }
                    }
                    if (k == 0 )
                    {
                        k = 0; 
                        //Код на добавление данных в БД о виде страхования
                        sqlConnection.Open();
                SqlCommand command = new SqlCommand($@"INSERT INTO [object](type,address,square,com) 
                            VALUES (@t,@a,@s,@c);", sqlConnection);
                command.Parameters.AddWithValue("@t", ("Объект благоустройства"));
                command.Parameters.AddWithValue("@a", (textBox7.Text));
                command.Parameters.AddWithValue("@s", Convert.ToDecimal(textBox8.Text));
                command.Parameters.AddWithValue("@c", ("Тип объекта: " + comboBox1.Text + "; \n" +
                    "- Комунникации: " + tex + " \n" +
                    "- Придомовая территория: " + ter
                    ));
                command.ExecuteNonQuery();
                sqlConnection.Close();
                Vid_load();
                textBox2.Text = "";
                textBox3.Text = "";
                clear();
                panel12.Visible = false;
                panel2.Visible = false;
                panel11.Visible = false;
                    }
                    else
                    {
                        MessageBox.Show("Такой объект благоустройства уже есть!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                }
                else
                {
                    MessageBox.Show("Заполните поля!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else if (label6.Text == "Редактировать объект благоустройства")
            {
                if (textBox7.Text != "" && textBox8.Text != "" && comboBox1.Text != "")
                {
                    k = 0;j = 0;
                    for (int i = 0; i < dataGridView1.RowCount; i++)
                    {
                        if ("объект благоустройства" + textBox7.Text.ToLower() == dataGridView1[1, i].Value.ToString().ToLower()+ dataGridView1[2, i].Value.ToString().ToLower())
                        {
                            k++;
                            j = Convert.ToInt32(dataGridView1[0, i].Value.ToString());
                        }
                    }
                    if (k == 0 || j == id)
                    {
                        k = 0; j = 0;
                        sqlConnection.Open();
                SqlCommand command = new SqlCommand($@"UPDATE object SET address=@a, square=@s, com=@c WHERE idobject=@id", sqlConnection);
                command.Parameters.AddWithValue("@a", (textBox7.Text));
                command.Parameters.AddWithValue("@s", Convert.ToDecimal(textBox8.Text));
                command.Parameters.AddWithValue("@c", "Тип объекта: " + comboBox1.Text + "; \n" +
                    "- Комунникации: " + tex + " \n" +
                    "- Придомовая территория: " + ter);
                command.Parameters.AddWithValue("@id", (id));
                command.ExecuteNonQuery();
                sqlConnection.Close();
                Vid_load();
                textBox2.Text = "";
                textBox3.Text = "";
                panel2.Visible = false;
                panel11.Visible = false;
                panel12.Visible = false;
                dataGridView1.Enabled = true;
                clear();
                    }
                    else
                    {
                        MessageBox.Show("Такой объект благоустройства уже есть!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("Заполните поля!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            //}
            //    catch { }
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            e.Handled = !((c == 8 || c == '1' || c == '2' || c == '3' || c == '4' || c == '5' || c == '6' || c == '7' || c == '8' || c == '9'));
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            id = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
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

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            
           
        }
        public void clear()
        {
            voda = "";
            elec = "";
            canal = "";
            otop = "";
            gaz = "";
            ter = "";
            tex = "";
            lift = "";
            del = "";
            d = "";
            delete = "";
            id= 0;
            checkBox1.Checked = false;
            checkBox2.Checked = false;
            checkBox3.Checked = false;
            checkBox4.Checked = false;
            checkBox5.Checked = false;
            checkBox6.Checked = false;
            checkBox7.Checked = false;
            checkBox8.Checked = false;
            checkBox9.Checked = false;
            checkBox10.Checked = false;
            checkBox11.Checked = false;
            checkBox12.Checked = false;
            checkBox13.Checked = false;
            checkBox14.Checked = false;
            checkBox15.Checked = false;
            checkBox16.Checked = false;
            checkBox17.Checked = false;
            checkBox18.Checked = false;
            checkBox19.Checked = false;
            checkBox20.Checked = false;
            checkBox21.Checked = false;
            checkBox22.Checked = false;
            checkBox23.Checked = false;
            checkBox24.Checked = false;
            checkBox25.Checked = false;
            checkBox27.Checked = false;
            checkBox29.Checked = false;
            checkBox26.Checked = false;
            checkBox28.Checked = false;
            checkBox30.Checked = false;
            checkBox31.Checked = false;
            checkBox32.Checked = false;
            textBox4.Text = "";
            textBox6.Text = "";
            textBox5.Text = "";
            textBox7.Text = "";
            textBox8.Text = "";
            textBox3.Text = "";
            textBox2.Text = "";
            comboBox1.SelectedIndex = -1;

        }
    string voda = "";
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                checkBox2.Checked = false;
                voda = "Водоснабжение: централизованное; ";
            }
            else
            {
                checkBox1.Checked = false;
                voda = "";
            }
        }
        
    
        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                checkBox1.Checked = false;
                voda = "Водоснабжение: автономное; ";
            }
            else
            {
                checkBox2.Checked = false;
                voda = "";
            }
        }

        string otop = "";
        private void checkBox8_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox8.Checked == true)
            {
                checkBox7.Checked = false;
                otop = "Отопление: централизованное; ";

            }
            else
            {
                checkBox8.Checked = false;
                otop = "";
            }
        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox7.Checked == true)
            {
                checkBox8.Checked = false;
                otop = "Отопление: автономное; ";
            }
            else
            {
                checkBox7.Checked = false;
                otop = "";
            }
        }
        string elec = "";
        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked == true)
            {
                elec = "Электричество: центральное; ";
                checkBox3.Checked = false;
            }
            else
            {
                checkBox4.Checked = false;
                elec = "";
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked == true)
            {
                elec = "Электричество: автономное; ";
                checkBox4.Checked = false;

            }
            else
            {
                checkBox3.Checked = false;
                 elec = "";
            }
        }
        string canal = "";
        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox6.Checked == true)
            {
                canal= "Канализация: централизованная; ";
                checkBox5.Checked = false;
            }
            else
            {
                checkBox6.Checked = false;
                canal = "";
            }
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox5.Checked == true)
            {
                canal = "Канализация: автономная; ";
                checkBox6.Checked = false;
            }
            else
            {
                checkBox5.Checked = false;
                canal = "";
            }
        }
        string gaz = "";
        private void checkBox10_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox10.Checked == true)
            {
                gaz = "Газ: централизованный газопровод; ";
                checkBox9.Checked = false;
            }
            else
            {
                gaz = "";
                checkBox10.Checked = false;
            }
        }

        private void checkBox9_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox9.Checked == true)
            {
                gaz = "Газ: cниженный газ в баллонах; ";
                checkBox10.Checked = false;
            }
            else
            {
                gaz = "";
                checkBox9.Checked = false;
            }
        }

       

        private void textBox8_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && e.KeyChar != 8)
                e.Handled = true;
        }

        private void textBox7_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            e.Handled = !((c >= 'а' && c <= 'я') || (c >= 'А' && c <= 'Я') || (c == 'ё' || c == '-' || c == '.' || c == ',' || c == 8 || c == 32));
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            e.Handled = !((c == 8 || c == '0' || c == '1' || c == '2' || c == '3' || c == '4' || c == '5' || c == '6' || c == '7' || c == '8' || c == '9'));
        }

        private void textBox6_KeyPress(object sender, KeyPressEventArgs e)
        {
            delete = "Пассажирский: " + textBox6.Text + "; ";
            char c = e.KeyChar;
            e.Handled = !((c == 8 || c == '0' || c == '1' || c == '2' || c == '3' || c == '4' || c == '5' || c == '6' || c == '7' || c == '8' || c == '9'));

        }

        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            delete = "Грузопассажирский: " + textBox5.Text + "; ";
            char c = e.KeyChar;
            e.Handled = !((c == 8 || c == '0' || c == '1' || c == '2' || c == '3' || c == '4' || c == '5' || c == '6' || c == '7' || c == '8' || c == '9'));

        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            delete = "Пожарный: " + textBox4.Text + "; ";
            char c = e.KeyChar;
            e.Handled = !((c == 8 || c == '0' || c == '1' || c == '2' || c == '3' || c == '4' || c == '5' || c == '6' || c == '7' || c == '8' || c == '9'));

        }
        string lift = "";string delete = "";
        private void checkBox12_CheckedChanged(object sender, EventArgs e)
        {
            
          
           
                             
        }

        private void checkBox13_CheckedChanged(object sender, EventArgs e)
        {
            
         
        }

        private void checkBox14_CheckedChanged(object sender, EventArgs e)
        {
           
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            if (textBox6.Text!="")
            {
                
                checkBox12.Checked = true;

            }
            else
            {
                checkBox12.Checked = false;
            }
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

            if (textBox5.Text != "")
            {
                checkBox13.Checked = true;
            }
            else
            {
                
                checkBox13.Checked = false;
            }
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            if (textBox4.Text != "")
            {
                checkBox14.Checked = true;
            }
            else
            {
                
                checkBox14.Checked = false;
            }
        }

        private void textBox4_Leave(object sender, EventArgs e)
        {
            if (checkBox14.Checked == true)
            {
                
                lift += "Пожарный: " + textBox4.Text + "; ";
                
            }
            else
            {
                lift = lift.Replace(delete, "");
                lift = lift.Replace("  ", " ");

            }
        }

        private void textBox5_Leave(object sender, EventArgs e)
        {
            if (checkBox13.Checked == true)
            {
                lift += "Грузопассажирский: " + textBox5.Text + "; ";
                
            }
            else
            {
                lift = lift.Replace(delete, "");
                lift = lift.Replace("  ", " ");
            }
        }

        private void textBox6_Leave(object sender, EventArgs e)
        {
            if (checkBox12.Checked == true)
            {
                lift += "Пассажирский: " + textBox6.Text + "; ";
              
            }
            else
            {
  
                lift = lift.Replace(delete, "");
                lift = lift.Replace("  ", " ");
            }
        }
        string tex = ""; string d = "";
        private void checkBox18_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox18.Checked == true)
            {
                tex += "Подвал; ";
              
            }
            else
            {  d = "Подвал; ";

                tex = tex.Replace(d, "");
                tex = tex.Replace("  ", " ");
            }
        }

        private void checkBox17_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox17.Checked == true)
            {
                tex += "Чердак; ";
               
            }
            else
            { d = "Чердак; ";

                tex = tex.Replace(d, "");
                tex = tex.Replace("  ", " ");
            }
        }

        private void checkBox16_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox16.Checked == true)
            {
                tex += "Электросчетавая; ";
              
            }
            else
            {  d = "Электросчетавая; ";

                tex = tex.Replace(d, "");
                tex = tex.Replace("  ", " ");
            }
        }

        private void checkBox19_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox19.Checked == true)
            {
                tex += "Машинное отделение лифтов; ";
               
            }
            else
            {
 d = "Машинное отделение лифтов; ";
                tex = tex.Replace(d, "");
                tex = tex.Replace("  ", " ");
            }
        }
        string ter = ""; string del = "";
        private void checkBox24_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox24.Checked == true)
            {
                ter += "Детская площадка; ";
               
            }
            else
            { del = "Детская площадка; ";

                ter = ter.Replace(del, "");
                ter = ter.Replace("  ", " ");
            }
        }

        private void checkBox23_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox23.Checked == true)
            {
                ter += "Спортивная площадка; ";
               
            }
            else
            { del = "Спортивная площадка; ";

                ter = ter.Replace(del, "");
                ter = ter.Replace("  ", " ");
            }
        }

        private void checkBox21_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox21.Checked == true)
            {
                ter += "Зеленые зоны; ";
               
            }
            else
            { del = "Зеленые зоны; ";

                ter = ter.Replace(del, "");
                ter = ter.Replace("  ", " ");
            }
        }

        private void checkBox22_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox22.Checked == true)
            {
                ter += "Места для парковки; ";
               
            }
            else
            { del = "Места для парковки; ";

                ter = ter.Replace(del, "");
                ter = ter.Replace("  ", " ");
            }
        }

        private void textBox5_Layout(object sender, LayoutEventArgs e)
        {

        }

        private void textBox5_KeyUp(object sender, KeyEventArgs e)
        {
            delete = "Грузопассажирский: " + textBox5.Text + "; ";

        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (comboBox2.SelectedIndex == 0)
            {
                label2.Visible = true;
                label14.Visible = true;
                textBox2.Visible = true;
                textBox3.Visible = true;
                panel2.Left = 117;
                comboBox2.SelectedIndex = -1;
                dataGridView1.Enabled = true;
                panel6.Visible = false;
                panel2.Visible = true;
                panel11.Visible = false;
                textBox2.Text = "";
                label6.Text = "Добавить многоквартирный дом";
                button11.Text = "Добавить";
                button11.Width = 174;
                button11.Top = 349;
                panel2.Width = 1496;
                label3.Text = "Тип строения";
                comboBox1.Text = "";
                comboBox1.Items.Clear();
                comboBox1.Items.Add("Панельный");
                comboBox1.Items.Add("Кирпичный");
                comboBox1.Items.Add("Монолитный");
                comboBox1.Items.Add("Комбинированный"); 
                comboBox1.SelectedIndex = -1;
                Vid_load();
                clear();
                button11.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (31).png");
            }
            else if (comboBox2.SelectedIndex == 1)
            {
                panel2.Left = 317;
                panel11.Left = 317;
                comboBox2.SelectedIndex = -1;
                panel6.Visible = false;
                dataGridView1.Enabled = true;
                panel2.Visible = true;
                panel11.BringToFront();
                panel11.Visible = true;
                textBox2.Text = "";
                label6.Text = "Добавить частный дом (коттедж)";
                button11.Text = "Добавить";
                button11.Width = 174;
                button11.Top = 349;
                panel2.Width = 973;
                comboBox1.Visible = true;
                label3.Text = "Тип строения";
                comboBox1.Text = "";
                comboBox1.Items.Clear();
                comboBox1.Items.Add("одноэтажный");
                comboBox1.Items.Add("двухэтажные");
                comboBox1.Items.Add("многосекционные");
                comboBox1.SelectedIndex = -1;
                Vid_load();
                clear();
                button11.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (31).png");
            }

            else if (comboBox2.SelectedIndex == 2)
            {
                panel2.Left = 317;
                panel2.Width = 509;
                panel12.Visible = true;
                panel12.Left = 826;
                label2.Visible = false;
                label14.Visible = false;
                textBox2.Visible = false;
                textBox3.Visible = false;
                comboBox2.SelectedIndex= -1;
                button11.Top = 240;
                panel6.Visible = false;
                dataGridView1.Enabled = true;
                panel2.Visible = true;
                textBox2.Text = "";
                label6.Text = "Добавить объект благоустройства";
                button11.Text = "Добавить";
                button11.Width = 174;
                label3.Text = "Тип объекта";
                comboBox1.Text = "";
                comboBox1.Items.Clear();
                comboBox1.Items.Add("парк");
                comboBox1.Items.Add("сквер");
                comboBox1.Items.Add("зона отдыха");
                comboBox1.SelectedIndex = -1;
                Vid_load();
                clear();
                button11.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (31).png");
            }
        }

        private void label17_Click(object sender, EventArgs e)
        {

        }

        private void label18_Click(object sender, EventArgs e)
        {

        }

        private void checkBox29_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox29.Checked == true)
            {
                ter += "Сад; ";

            }
            else
            {
                del = "Сад; ";

                ter = ter.Replace(del, "");
                ter = ter.Replace("  ", " ");
            }
        }

        private void checkBox26_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox26.Checked == true)
            {
                ter += "Хозяйственная постройка; ";

            }
            else
            {
                del = "Хозяйственная постройка; ";

                ter = ter.Replace(del, "");
                ter = ter.Replace("  ", " ");
            }
        }

        private void checkBox28_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox28.Checked == true)
            {
                ter += "Газон; ";

            }
            else
            {
                del = "Газон; ";

                ter = ter.Replace(del, "");
                ter = ter.Replace("  ", " ");
            }
        }

        private void checkBox31_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox31.Checked == true)
            {
                tex += "Детская площадка; ";

            }
            else
            {
                d = "Детская площадка; ";

                tex = tex.Replace(d, "");
                tex = tex.Replace("  ", " ");
            }
        }

        private void checkBox27_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox27.Checked == true)
            {
                tex += "Скамейки; ";

            }
            else
            {
                d = "Скамейки; ";


                tex = tex.Replace(d, "");
                tex = tex.Replace("  ", " ");
            }
        }

        private void checkBox30_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox30.Checked == true)
            {
                tex += "Спортивная площадка; ";

            }
            else
            {
                d = "Спортивная площадка; ";


                tex = tex.Replace(d, "");
                tex = tex.Replace("  ", " ");
            }
        }

        private void checkBox32_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox32.Checked == true)
            {
                tex += "Фонари; ";

            }
            else
            {
                d = "Фонари; ";

                tex = tex.Replace(d, "");
                tex = tex.Replace("  ", " "); ;
            }
        }

        private void checkBox25_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox25.Checked == true)
            {
                tex += "Зеленые насаждения; ";

            }
            else
            {
                d = "Зеленые насаждения; ";

                tex = tex.Replace(d, "");
                tex = tex.Replace("  ", " ");
            }
        }

        private void checkBox15_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox15.Checked == true)
            {
                ter += "Освещение; ";

            }
            else
            {
                del = "Освещение; ";

                ter = ter.Replace(del, "");
                ter = ter.Replace("  ", " ");
            }
        }

        private void checkBox20_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox20.Checked == true)
            {
                ter += "Санитарные зоны; ";

            }
            else
            {
                del = "Санитарные зоны; ";

                ter = ter.Replace(del, "");
                ter = ter.Replace("  ", " ");
            }
        }

        private void checkBox11_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox11.Checked == true)
            {
                ter += "Поливочные системы; ";

            }
            else
            {
                del = "Поливочные системы; ";

                ter = ter.Replace(del, "");
                ter = ter.Replace("  ", " ");
            }

        }

        private void button7_Click(object sender, EventArgs e)
        {
            panel1.Enabled = true;
            dataGridView1.Enabled = true;
            panel2.Visible = false;
            panel12.Visible = false;
            panel11.Visible = false;
            panel6.Visible = false;
        }
    }
}
