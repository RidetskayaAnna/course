using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace Agent
{
    public partial class Form1 : System.Windows.Forms.Form
    {
        datebase db = new datebase();
        public SqlConnection sqlConnection = null;
        public Form1()
        {
            InitializeComponent();
            sqlConnection = new SqlConnection(db.connection);

        }
        private void addControll(UserControl uc)
        {

            panel5.Controls.Clear();
            panel5.Controls.Add(uc);

        }
        public Boolean press = true;
        private System.Drawing.Point mouseOffset;
        private bool isMouseDown = false;
        private void Form1_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) this.Close();
        }

        private void Form1_MouseDown(object sender, MouseEventArgs e)
        {
            int xOffset;
            int yOffset;

            if (e.Button == MouseButtons.Left)
            {
                xOffset = -e.X - SystemInformation.FrameBorderSize.Width;
                yOffset = -e.Y - SystemInformation.CaptionHeight -
                    SystemInformation.FrameBorderSize.Height;
                mouseOffset = new System.Drawing.Point(xOffset, yOffset);
                isMouseDown = true;
            }
        }

        private void Form1_MouseMove(object sender, MouseEventArgs e)
        {
            if (isMouseDown)
            {
                System.Drawing.Point mousePos = Control.MousePosition;
                mousePos.Offset(mouseOffset.X, mouseOffset.Y);
                Location = mousePos;
            }
        }

        private void Form1_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                isMouseDown = false;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form.Working uc = new Form.Working(this);
            button7.Visible = false;
            button10.BringToFront();
            button10.Visible = true;
            button8.Visible = false;
            button9.Visible = false;
            panel4.Left = 204;
            panel4.Top = button1.Top;
            panel4.Height = button1.Height;
            addControll(uc);
        }

        private void button2_Click(object sender, EventArgs e)
        {
    
            Form.Worker insurer = new Form.Worker();
            addControll(insurer);
            button10.Visible = false;
            button7.Visible = false;     
            button8.BringToFront();
            button8.Visible = true;
            button9.Visible = false;
            panel4.Left = 204;
            panel4.Top = button2.Top;
            panel4.Height = button2.Height;
        }



        private void button4_Click(object sender, EventArgs e)
        {

            Form.Tenant policyholder = new Form.Tenant();
            addControll(policyholder);
            panel4.Top = button4.Top;
            panel4.Left = 204;
            button8.Visible = false;
            panel4.Height = button4.Height;
            button10.Visible = false;
            button7.Visible = false;
            button9.Visible = false;
        }



        private void button6_Click(object sender, EventArgs e)
        {
  
            Form.Static staticc = new Form.Static();
            addControll(staticc);
            panel4.Height = button6.Height;
            panel4.Top = button6.Top;
            panel4.Left = 204;
            button10.Visible = false;
            button8.Visible = false;
            button7.Visible = false;
            button9.Visible = false;

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Form.Working uc = new Form.Working(this);
            button7.Visible = false;
      
            button10.BringToFront();
            button10.Visible = true;
            button9.Visible = false;
            button8.Visible = false;
            panel4.Top = button1.Top;
            panel4.Left = 204;
            panel4.Height = button1.Height;
            addControll(uc);
        }

        private void button7_Click_2(object sender, EventArgs e)
        {
         
            Form.Pays uc = new Form.Pays(this);
            panel4.Top = button7.Top;
            panel4.Left = 1491;
            panel4.Height = button7.Height;
            button7.BringToFront();
            panel4.BringToFront();
            button7.Visible = true;
            button9.Visible = true;
            button8.Visible = false;
            button10.Visible = true;
            addControll(uc);
        }



        private void button9_Click(object sender, EventArgs e)
        {
         
            Form.Back uc = new Form.Back(this);
            panel4.Top = button9.Top;
            panel4.Left = 1346;
            panel4.Height = button9.Height;
            button9.Visible = true;
            button9.BringToFront();
            panel4.BringToFront();
            button7.Visible = true;
            button8.Visible = false;
            addControll(uc);
        }

        private void button11_Click(object sender, EventArgs e)
        {
         
            Form.Object uc = new Form.Object(this);
            button8.Visible = false;
            button10.Visible = false;
            button7.Visible = false;
            button9.Visible = false;
            button9.BringToFront();
            panel4.Top = button11.Top;
            panel4.Left = 204;
            panel4.Height = button11.Height;
            addControll(uc);
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
          
            Form.Service uc = new Form.Service(this);
            button8.Visible = false;
            button7.Visible = false;
            button9.Visible = false;
            button10.Visible = false;
            panel4.Top = button3.Top;
            panel4.Left = 204;
            panel4.Height = button3.Height;
            addControll(uc);
        }

        private void button5_Click_1(object sender, EventArgs e)
        {
         
            Notification uc = new Notification(this);
            button8.Visible = false;
            button7.Visible = false;
            button10.Visible = false;
            button9.Visible = false;
            panel4.Top = button5.Top;
            panel4.Left = 204;
            panel4.Height = button5.Height;
            addControll(uc);
        }
        private void button8_Click_1(object sender, EventArgs e)
        {
      
            Form.Work uc = new Form.Work();
            button8.Visible = true;
            button7.Visible = false;
            button9.Visible = false;
            panel4.Top = button8.Top;
            button10.Visible = false;
            panel4.Height = button8.Height;
            button8.BringToFront();
            panel4.Left = 1205;
            addControll(uc);
        }

        private void button10_Click(object sender, EventArgs e)
        {
            Form.Bid uc = new Form.Bid(this);
            button8.Visible = false;
            button7.BringToFront();
            button7.Visible = true;
            button10.BringToFront();    
            button10.Visible = true;  
            button9.BringToFront();
            button9.Visible = true;
            panel4.Top = button10.Top;
            panel4.Left = 1205;
            panel4.Height = button10.Height;
            addControll(uc);
        }

      
        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            //sqlConnection.Open();
            //string query = $@"DELETE FROM [sostav] WHERE [idtreaty] is null";
            //SqlCommand command = new SqlCommand(query, sqlConnection);
            //command.ExecuteNonQuery();
            //sqlConnection.Close();
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
          
        }

        private void metroSetControlBox1_Click(object sender, EventArgs e)
        {

        }

        private void button13_Click(object sender, EventArgs e)
        {
           
            Form.Calendar calendar = new Form.Calendar();
            addControll(calendar);
            panel4.Height = button13.Height;
            panel4.Top = button13.Top;
            panel4.Left = 204;
            button10.Visible = false;
            button8.Visible = false;
            button7.Visible = false;
            button9.Visible = false;
        }
    }
}
