using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace Agent
{
    public partial class Form1 :System.Windows.Forms.Form
    {
        public Form1()
        {
            InitializeComponent();
            
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
            Form.Treaty uc = new Form.Treaty(this);  
            button7.BringToFront(); 
            button7.Visible = true;
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
            Form.Insurer insurer = new Form.Insurer();
            addControll(insurer);
            button10.Visible = false;
            button7.Visible = false;
            button8.Visible = false;
            button9.Visible = false;
            panel4.Left = 204;
            panel4.Top = button2.Top;
            panel4.Height = button2.Height;
        }

     

        private void button4_Click(object sender, EventArgs e)
        {
            Form.Policyholder policyholder = new Form.Policyholder();
            addControll(policyholder);
            panel4.Top = button4.Top;
            panel4.Left = 204;
            button8.BringToFront();
            button8.Visible = true;
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
            Form.Treaty uc = new Form.Treaty(this);
            button7.BringToFront();
            button7.Visible = true;
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
            panel4.Left = 1038;
            panel4.Height = button7.Height;
            button7.BringToFront();
            panel4.BringToFront();
            button7.Visible = true;
            button9.Visible = false;
            button8.Visible = false;
            button10.Visible = false;
            addControll(uc);
        }

    

        private void button9_Click(object sender, EventArgs e)
        {
            Form.Correctionfactor uc = new Form.Correctionfactor();
            panel4.Top = button9.Top;
            panel4.Left = 1286;
            panel4.Height = button9.Height;
            button9.Visible = true;
            button9.BringToFront();
            panel4.BringToFront();
            button7.Visible = false;
            button8.Visible = false;
            addControll(uc);
        }

        private void button11_Click(object sender, EventArgs e)
        {
            Form.Vid uc = new Form.Vid(this);
            button8.Visible = false;
            button10.Visible = false;
            button7.Visible=false;
            button9.Visible = true;
            button9.BringToFront();
            panel4.Top = button11.Top;
            panel4.Left = 204;
            panel4.Height = button11.Height;
            addControll(uc);
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            Form.City uc = new Form.City();
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

        private void panel5_Click(object sender, EventArgs e)
        {
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
            button7.Visible = false;

            button10.Visible = true;
            button9.Visible = false;
            button10.BringToFront();
            panel4.Top = button10.Top;
            panel4.Left = 1202;
            panel4.Height = button10.Height;
            addControll(uc);
        }
    }
}
