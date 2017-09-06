using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel; 

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        int flag = 0;
        public Form2 form2 = new Form2();
        public Form1()
        {
            InitializeComponent();
            //timer1.Start();
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
            form2.Tag = this;
           form2.Show(this);
           // Hide();
           
        
        }

        private void button2_Click(object sender, EventArgs e)
        {
            
           textBox1.Text = form2.excelfile();

        }

  /*      private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                panel2.Enabled = false;
                panel3.Enabled = false;
                panel4.Enabled = true;
                button1.Enabled = true;
            }
            else
            {
                panel2.Enabled = true;
                panel3.Enabled = true;
                panel4.Enabled = false;
                this.button1.BackColor = System.Drawing.Color.DarkOliveGreen;
                this.button1.Text = "PLAY";
                button1.Enabled = false;
            }

        
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                panel3.Enabled = false;
                panel1.Enabled = false;
                panel5.Enabled = true;
                button1.Enabled = true;
            }
            else
            {
                panel1.Enabled = true;
                panel3.Enabled = true;
                panel5.Enabled = false;
                this.button1.BackColor = System.Drawing.Color.DarkOliveGreen;
                this.button1.Text = "PLAY";
                button1.Enabled = false;
            }


        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked == true)
            {
                panel2.Enabled = false;
                panel1.Enabled = false;
                panel6.Enabled = true;
                button1.Enabled = true;
            }
            else
            { 
                panel1.Enabled = true;
                panel2.Enabled = true;
                panel6.Enabled = false;
                this.button1.BackColor = System.Drawing.Color.DarkOliveGreen;
                this.button1.Text = "PLAY";
                button1.Enabled = false;
                
            }
        }*/

        private void trackBar1_Scroll(object sender, EventArgs e)
        {
            form2.speeds(Convert.ToInt32(trackBar1.Value));
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            if (textBox1.Text == "")
            { MessageBox.Show("PLEASE SELECT ANY FILE"); }
            else
            {
                if (this.button1.Text == "PLAY")
                {
                    this.button1.BackColor = System.Drawing.Color.IndianRed;
                    this.button1.Text = "PAUSE";
                        if (flag == 0)
                            form2.excelprocess();
                        else
                        { form2.starttimer(); flag = 0; }
                  

                }
                else
                {
                    this.button1.BackColor = System.Drawing.Color.DarkOliveGreen;
                    this.button1.Text = "PLAY";
                    flag = 1;
                    form2.tymstopper();

                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (radioButton1.Checked == true)
               form2.Color(1);
            else if (radioButton2.Checked == true)
               form2.Color(2); 
            else
                form2.Color(3); 
 }

        private void button6_Click(object sender, EventArgs e)
        {
            if (textBox2.Text == "")
            { MessageBox.Show("PLEASE SELECT ANY FILE"); }
            else
            {
                if (this.button1.Text == "PAUSE")
                {
                    this.button1.BackColor = System.Drawing.Color.DarkOliveGreen;
                    this.button1.Text = "PLAY";
                    flag = 1;
                    form2.tymstopper();
                }
                form2.image();

            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            textBox2.Text = form2.img_file();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            textBox3.Text = form2.flash_file();
        }

     

   /*     private void button1_Click(object sender, EventArgs e)
        {
             Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }


            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlWorkSheet.Cells[1, 1] = "Sheet 1 contentarun";

            xlWorkSheet.Cells[2, 1] = "hiran";
            xlWorkSheet.Cells[1, 2] = "hima";
            xlWorkBook.SaveAs("d:\\aruncsharp-Excel.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);

            MessageBox.Show("Excel file created , you can find the file d:\\csharp-Excel.xls");
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

       private void timer1_Tick_1(object sender, EventArgs e)
        {
            label1.Location = new Point(label1.Location.X, label1.Location.Y - 5);

            if (label1.Location.Y < 0)
            {
                label1.Location = new Point(label1.Location.X, this.Height);
            }

            label2.Location = new Point(label2.Location.X + 5, label2.Location.Y);

            if (label2.Location.X  > this.Width)
            {
                label2.Location = new Point(0 - label2.Width, label2.Location.Y);
            }

        }

      */

      /* ORGINAL COPY    
        private void timer1_Tick(object sender, EventArgs e)
        {
            label2.Location = new Point(label2.Location.X + 5, label2.Location.Y);
 
            if(label2.Location.X  > this.Width)
            {
                label2.Location = new Point(0 - label2.Width, label2.Location.Y);
            }
        }*/


    }
}
