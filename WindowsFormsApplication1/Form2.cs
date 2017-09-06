using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;


namespace WindowsFormsApplication1
{
    public partial class Form2 : Form
    {
        public string[] comp;
        public string[] last;
        public double[] per;
        public string path,path_img,path_flash;
        public int count = 0, i = 1, colour,speed = 1;
        
        
        public Form2()
        {
            InitializeComponent();

        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }

        public string excelfile()
        {
            DialogResult dlgResult = openFileDialog1.ShowDialog();
            if (dlgResult.Equals(DialogResult.OK))
            {
                path = openFileDialog1.FileName;
                return openFileDialog1.FileName;
            }
            return null;
        }
        public string img_file()
        {
            //openFileDialog1.Filter = "Image Files (JPG,PNG,GIF)|*.JPG;*.PNG;*.GIF";
            DialogResult dlgResult = openFileDialog1.ShowDialog();
            if (dlgResult.Equals(DialogResult.OK))
            {
                path_img = openFileDialog1.FileName;
                return openFileDialog1.FileName;
            }
            return null;
        }

        public string flash_file()
        {
            DialogResult dlgResult = openFileDialog1.ShowDialog();
            if (dlgResult.Equals(DialogResult.OK))
            {
                path_flash = openFileDialog1.FileName;
                return openFileDialog1.FileName;
            }
            return null;
        }

        public void Color(int c)
        {

            colour = c;

        }
        public void speeds(int s)
        {
           // if (s == 0)
            //    speed = 0.5;
           // else if (s == 1)
            //    speed = 1;
            //else
                speed = s+1;

        }
        public void tymstopper()
        {
           
            timer1.Stop();
            
        }
        public void starttimer()
        {
            this.BackgroundImage = null;
            timer1.Start();
           
        }
        public void image()
        {
            panel1.Visible = false; panel2.Visible = false; panel3.Visible = false;
            panel4.Visible = false; panel5.Visible = false; panel6.Visible = false;
            this.BackgroundImage = new Bitmap(@path_img);
        }


        public void excelprocess()
        {
        
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }

            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;

            string str;
            int rCnt = 0;
            int cCnt = 0;
            count = 0;
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(path, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;
            comp = new string[range.Rows.Count + 1];
            last = new string[range.Rows.Count + 1];
            per = new double[range.Rows.Count + 1];


            for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
            {
                count++;
                for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
                {
                    try
                    {
                        str = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                    }
                    catch (System.Exception excep)
                    { str = ((range.Cells[rCnt, cCnt] as Excel.Range).Value2).ToString("0.00"); }


                    if (cCnt == 1)
                    { comp[count] = str; }
                    else if (cCnt == 2)
                    { last[count] = str; }
                    else
                    { per[count] = Convert.ToDouble(str); }

                    // MessageBox.Show(str);

                }

            }

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
            panelstart();


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


        private void panelstart()
        {
            this.BackgroundImage = null;
            Control[] controls;
            Double sign;


            for (int j = 1; j <= 6; j++)
            {
                if (i > count)
                    i = 1;
                for (int k = 1; k <= 4; k++)
                {

                    //  MessageBox.Show("Eokeeeeeeeeeeeeeee");
                    controls = this.Controls.Find("Label" + j.ToString() + "_" + k, true);
                    if (controls.Length == 1) // 0 means not found, more - there are several controls with the same name
                    {
                        Label control = controls[0] as Label;
                        if (control != null)
                        {

                            if (colour == 1)
                                control.ForeColor = System.Drawing.Color.Red;
                            else if (colour == 2)
                                control.ForeColor = System.Drawing.Color.Green;
                            else
                                control.ForeColor = System.Drawing.Color.Yellow;

                            if (k == 1)
                                control.Text = comp[i];
                            else if (k == 2)
                                control.Text = last[i];
                            else if (k == 3)
                            {

                                if (per[i] < 0)
                                { control.Image = global::WindowsFormsApplication1.Properties.Resources.green1; }
                                else if (per[i] > 0)
                                    control.Image = global::WindowsFormsApplication1.Properties.Resources.red2;
                                else
                                    control.Image = global::WindowsFormsApplication1.Properties.Resources.yellow12;

                            }
                            else
                            {
                                if (per[i] < 0)
                                    sign = -1;
                                else
                                    sign = 1;
                                control.Text = (per[i] * sign).ToString("0.00") + "%";
                            }



                            //   MessageBox.Show("arunnn");

                        }
                    }
                }


                i++;

            }
            timer1.Start();
        }


        private void timer1_Tick_1(object sender, EventArgs e)
        {
           
            Control[] controls, panel;
            Double sign;
           
              //  panel1.Visible = true; panel2.Visible = true; panel3.Visible = true;
               // panel4.Visible = true; panel5.Visible = true; panel6.Visible = true;
               // panel1.Location = new Point(panel1.Location.X, (int)(panel1.Location.Y - speed));
               // panel2.Location = new Point(panel2.Location.X, (int)(panel2.Location.Y - speed));
               // panel3.Location = new Point(panel3.Location.X, (int)(panel3.Location.Y - speed));
               // panel4.Location = new Point(panel4.Location.X,(int)( panel4.Location.Y - speed));
                //panel5.Location = new Point(panel5.Location.X, (int)(panel5.Location.Y - speed));
                //panel6.Location = new Point(panel6.Location.X, (int)(panel6.Location.Y - speed));
               

                for (int j = 1; j <= 6; j++)       //select panel
                {
                    panel = this.Controls.Find("panel" + j, true);
                    if (panel.Length == 1)
                    {
                        Panel p = panel[0] as Panel;
                        if (p != null)
                        {
                            p.Visible = true;
                      p.Location = new Point(p.Location.X, (int)(p.Location.Y - speed));
                            if (p.Location.Y + p.Height < 0)
                            {
                                if (i > count)
                                    i = 1;
                                for (int k = 1; k <= 4; k++)                 //select label
                                {
                                    controls = this.Controls.Find("Label" + j + "_" + k, true);
                                    if (controls.Length == 1)
                                    {
                                        Label control = controls[0] as Label;
                                        if (control != null)
                                        {
                                            if (colour == 1)
                                                control.ForeColor = System.Drawing.Color.Red;
                                            else if (colour == 2)
                                                control.ForeColor = System.Drawing.Color.Green;
                                            else
                                                control.ForeColor = System.Drawing.Color.Yellow; //colour selection

                                            if (k == 1)
                                                control.Text = comp[i];
                                            else if (k == 2)
                                                control.Text = last[i];
                                            else if (k == 3)
                                            {

                                                if (per[i] < 0.00)
                                                { control.Image = global::WindowsFormsApplication1.Properties.Resources.red2; }
                                                else if (per[i] > 0.00)
                                                    control.Image = global::WindowsFormsApplication1.Properties.Resources.green1;
                                                else
                                                    control.Image = global::WindowsFormsApplication1.Properties.Resources.yellow12;

                                            }
                                            else
                                            {
                                                if (per[i] < 0)
                                                    sign = -1;
                                                else
                                                    sign = 1;
                                                control.Text = (per[i] * sign).ToString("0.00") + "%";
                                            }                                             //insert label values



                                        }
                                    }
                                }
                                i++;
                                p.Location = new Point(p.Location.X, this.Height);

                            }
                        }
                    }
                }
            
             Thread.Sleep((int)(15.00));

        }
    }
}
