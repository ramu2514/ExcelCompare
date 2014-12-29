using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace Excel_Compare
{
     
    public partial class UserInterface : Form
    {

//        static Panel activePanel= new Panel();
        static bool sheetIgnore, trim, ignore, highlight;
        static int mode;
        string newFile = "";
        string sheetCompareFile = "";
        string oldFile="";
        string outputFile = "";
        //activePanel = this.panel2;
        static bool launch = true;
        public UserInterface()
        {
            Type officeType = Type.GetTypeFromProgID("Excel.Application");
            if (officeType == null)
            {
                MessageBox.Show("MS Ofice Not found.\nAborting Program");
                return;
            }
            else
            {
                InitializeComponent();
                this.comboBox1.SelectedIndex = 0;
                panel1.Visible = false; panel2.Visible = true; panel3.Visible = false; panel4.Visible = false;
            }
       }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult result = openNewFile.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                string file = openNewFile.FileName;
                try
                {
                    newFile = file;
                    label1.Text = Path.GetFileName(file);
                }
                catch (IOException)
                {
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult result = openNewFile.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                string file = openNewFile.FileName;
                try
                {
                    oldFile = file;
                    label2.Text = Path.GetFileName(file);
                }
                catch (IOException)
                {
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            button3.Enabled = false;
            pictureBox1.Visible = true;
            label3.Text = "Processing.....";
            string output;
            int status=99;
            mode = comboBox1.SelectedIndex + 1;
            launch=checkBox1.Checked;
            trim = checkBox3.Checked;
            ignore = checkBox5.Checked;
            sheetIgnore = checkBox2.Checked;
            highlight = checkBox4.Checked;
            output = System.Environment.GetEnvironmentVariable("TEMP");
            if (mode == 2)
                output += "\\temp.xml";
            else if (mode == 1)
                output += "\\temp.html";
            else //will never be possible unless program is hacked
            {
                MessageBox.Show("Illegal output mode.Setting xml output mode");
                mode = 1;
                output += "\\temp.xml";
            }

            if (newFile == "" || oldFile == "")
                MessageBox.Show("Both Files are needed to compare.");
            else
                status = CompareExcel(oldFile, newFile, output);
            //MessageBox.Show(""+status+launch);
            if(status==0)
            {
                outputFile = output;
                button6.Visible = true;
            }
            if ((status == 0) && launch)
            {
                System.Diagnostics.Process.Start(@output);
            }
            button3.Enabled = true;
            pictureBox1.Visible = false;
            label3.Text = "";
        }

/*        public static int CompareExcel2HTML(string oldFile, string newFile, string outputFile)
        {
            int status = 0;
            Excel.Application objExcel;
            Excel.Workbook objWorkbook1 = null, objWorkbook2 = null;
            Excel.Worksheet objWorksheet1, objWorksheet2;

            objExcel = new Excel.Application();
            int WScount1, WScount2;
            objExcel.Visible = false;
            objExcel.DisplayAlerts = false;

            try
               {
                    if (!File.Exists(oldFile))
                    {
                        MessageBox.Show("Old File " + oldFile + " Not Found");
                        objExcel.Quit();
                        return 1;
                    }
                    if (!File.Exists(newFile))
                    {
                        MessageBox.Show("New File " + oldFile + " Not Found");
                        objExcel.Quit();
                        return 2;
                    }
                    if (oldFile == newFile)
                    {
                        MessageBox.Show("Both input files are same.You can not compare a file with itself.");
                        objExcel.Quit();
                        return 10;
                    }
                    try
                    {
                        objWorkbook1 = objExcel.Workbooks.Open(newFile);
                        objWorkbook2 = objExcel.Workbooks.Open(oldFile);
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("MS office should be insalled.");
                        return 100;
                    }
                    using (StreamWriter sw = new StreamWriter(outputFile))
                    {

                        sw.WriteLine("<!DOCTYPE HTML><html><head><title>Excel Compare Reslt</title><style>th,td{word-wrap: break-word;max-width: 300px;} body{margin:0px;font:12px verdana;} table,tr,th,td{font:12px verdana;border:1px solid black;border-collapse:collapse;padding:5px;text-align:center} th{color:white;background:green;}</style></head><body><center>");

                        WScount1 = objWorkbook1.Worksheets.Count;
                        WScount2 = objWorkbook2.Worksheets.Count;
                        if (WScount1 != WScount2)
                        {
                            MessageBox.Show("Sheet count not matched.\nCan not Compare Files.");
                            objWorkbook1.Close(true, null, null);
                            objWorkbook2.Close(true, null, null);
                            objExcel.Quit();
                            return 100;
                        }

                        for (int i = 1; i <= WScount1; i++)
                        {
                            string Name = objWorkbook1.Worksheets.get_Item(i).Name;
                            sw.WriteLine("<h1>" + Name + "</h1><table border='1'><tr><th>Location</th><th>Old File <br>" + Path.GetFileName(oldFile) + "</th><th>New File <br>" + Path.GetFileName(newFile) + "</th></tr>");
                            objWorksheet1 = objWorkbook1.Worksheets.get_Item(i);
                            objWorksheet2 = objWorkbook2.Worksheets.get_Item(i);

                            foreach (Excel.Range cell in objWorksheet1.UsedRange)
                            {

                                string changedCell = cell.get_Address();

                                if (cell.Text != objWorksheet2.get_Range(changedCell).Text)
                                {
                                    //cell.Interior.ColorIndex = 3 'Highlights in red color if any changes in cells
                                    sw.WriteLine("<tr><td>" + changedCell + "</td><td>" + cell.Text + "</td><td>" + objWorksheet2.get_Range(changedCell).Text + "</td></tr>");
                                }
                                // else
                                //    {
                                //        cell.Interior.ColorIndex = 0; 
                                //    }
                            }
                            sw.WriteLine("</table>");
                        }
                        sw.WriteLine("</cener></body></html>");
                    }
                }
                catch (System.IO.DirectoryNotFoundException)
                {
                    MessageBox.Show("Directory of output file does not exist or you may not have permissions");
                    objWorkbook1.Close(true, null, null);
                    objWorkbook2.Close(true, null, null);
                    objExcel.Quit();
                    return 50;
                }
                catch (System.IO.IOException Ex)
                {
                    string exception = "";
                    Match m = Regex.Match(Ex.ToString(), "^(.*)", RegexOptions.Multiline);
                    if (m.Success)
                        exception = m.Groups[0].Value;
                    MessageBox.Show("IO Exceception." + exception);
                    objWorkbook1.Close(true, null, null);
                    objWorkbook2.Close(true, null, null);
                    objExcel.Quit();
                    return 55;
                }
                catch (Exception Ex)
                {
                    string exception = "";
                    Match m = Regex.Match(Ex.ToString(), "^(.*)", RegexOptions.Multiline);
                    if (m.Success)
                        exception = m.Groups[0].Value;
                     MessageBox.Show("Unknow Exceception.\n" + exception);
                     objWorkbook1.Close(true, null, null);
                     objWorkbook2.Close(true, null, null);
                     objExcel.Quit();
                     return 99;
                }
            objWorkbook1.Close(true, null, null);
            objWorkbook2.Close(true, null, null);
            objExcel.Quit();
            return status;
        }
    */
        private void button4_Click(object sender, EventArgs e)
        {
            DialogResult result = openNewFile.ShowDialog(); // Show the dialog.
            comboBox2.Items.Clear();
            comboBox3.Items.Clear();
            if (result == DialogResult.OK) // Test result.
            {
                string file = openNewFile.FileName;
                try
                {
                    oldFile = newFile = "";
                    sheetCompareFile = file;
                    label4.Text = Path.GetFileName(file);
                }
                catch (IOException)
                {
                    return;
                }
                loadSheets();
            }
        }
        private void loadSheets()
        {
            pictureBox2.Visible = true;
            label6.Text = "Processing.....";

            Excel.Application objExcel;
            Excel.Workbook objWorkbook = null;
            // Excel.Worksheet objWorksheet1, objWorksheet2;
            try
            {
                objExcel = new Excel.Application();
                int WScount;
                objExcel.Visible = false;
                objExcel.DisplayAlerts = false;
                if (!File.Exists(sheetCompareFile))
                {
                    MessageBox.Show("File " + sheetCompareFile + " Not Found");
                    objExcel.Quit();
                    return;
                }
                try
                {
                    objWorkbook = objExcel.Workbooks.Open(sheetCompareFile);
                }
                catch (Exception)
                {
                    MessageBox.Show("MS office should be insalled.");
                    return;
                }
                WScount = objWorkbook.Worksheets.Count;
                string Name;
                for (int i = 0; i < WScount; i++)
                {
                    Name = objWorkbook.Worksheets.get_Item(i+1).Name;
                    comboBox2.Items.Add(Name);
                    comboBox3.Items.Add(Name);
                }
                comboBox2.SelectedIndex = 0;
                comboBox3.SelectedIndex = 0;
                comboBox2.Enabled = true;
                comboBox3.Enabled = true;
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Exception:" + Ex.Message);
                return;
            }
            pictureBox2.Visible = false;
            label6.Text = "";
            button5.Enabled = true;
            }

        private void button5_Click(object sender, EventArgs e)
        {
            int status = 99;
            string output;
            pictureBox2.Visible = true;
            label6.Text = "Comparing Sheets";
            button5.Enabled = true;
            mode = comboBox1.SelectedIndex + 1;
            launch = checkBox1.Checked;
            trim = checkBox3.Checked;
            ignore = checkBox5.Checked;
            sheetIgnore = checkBox2.Checked;
            highlight = checkBox4.Checked;
            output = System.Environment.GetEnvironmentVariable("TEMP");
            if (mode == 2)
                output += "\\temp.xml";
            else if (mode == 1)
                output += "\\temp.html";
            else //will never be possible unless program is hacked
            {
                MessageBox.Show("Illegal output mode.Setting xml output mode");
                mode = 1;
                output += "\\temp.html";
            }

            int left = comboBox2.SelectedIndex + 1;
            int right = comboBox3.SelectedIndex + 1;
            //MessageBox.Show(""+left+right);
            if (left==0)
                MessageBox.Show("Please choose a sheet(left side)");
            else if (right==0)
                MessageBox.Show("Please choose a sheet(right side)");
            else if (left==right)
                MessageBox.Show("You Crazy! Select Two different sheets");
            else
                status = CompareSheets(left,right,output);
            //MessageBox.Show(""+status+launch);
            if (status == 0)
            {
                outputFile = output;
                button6.Visible = true;
            }
            if ((status == 0) && launch)
            {
                System.Diagnostics.Process.Start(@output);
            }
            pictureBox2.Visible = false;
            label6.Text = "";
        }
        public int CompareSheets(int left, int right, string outputFile)
        {
            //.return 1;
            Excel.Application objExcel;
            Excel.Workbook objWorkbook1;
            Excel.Worksheet objWorksheet1, objWorksheet2;
            objExcel = new Excel.Application();
           // int WScount1, WScount2;
            objExcel.Visible = false;
            objExcel.DisplayAlerts = false;
            int status = 0;
           // sheetCompareFile = Path.GetFullPath(sh);
            try
            {
                objWorkbook1 = objExcel.Workbooks.Open(sheetCompareFile);
                using (StreamWriter sw = new StreamWriter(outputFile))
                {
                    // Excel.Range range ;
                    if (mode == 1) sw.WriteLine("<!DOCTYPE HTML><html><head><title>Excel Compare Reslt</title><style>th,td{word-wrap: break-word;max-width: 300px;} body{margin:0px;font:12px verdana;} table,tr,th,td{font:12px verdana;border:1px solid black;border-collapse:collapse;padding:5px;text-align:center} th{color:white;background:green;}</style></head><body><center>");
                    if (mode == 2)
                    {
                        sw.WriteLine("<?xml version=\"1.0\"?>");
                        sw.WriteLine("<ExcelCompare>");
                        sw.WriteLine(" <info>");
                        sw.WriteLine("  <date>" + DateTime.Now.ToString("MM-dd-yyyy HH:mm ss tt :") + "</date>");
                        sw.WriteLine("  <leftSheet>" + objWorkbook1.Worksheets.get_Item(left).Name + "</leftSheet>");
                        sw.WriteLine("  <rightSheet>" + objWorkbook1.Worksheets.get_Item(right).Name + "</rightSheet>");
                        sw.WriteLine("  <OutputMode>XML</OutputMode>");
                        sw.WriteLine("  <OutputPath>" + outputFile + "</OutputPath>");
                        sw.WriteLine(" </info>");
                    }
                    int i = 1;
                        if (mode == 1) sw.WriteLine("<h1>" + Name + "</h1><table border='1'><tr><th>Location</th><th>Sheet1 <br>" + objWorkbook1.Worksheets.get_Item(left).Name + "</th><th>Sheets<br>" + objWorkbook1.Worksheets.get_Item(right).Name + "</th></tr>");
                        objWorksheet1 = objWorkbook1.Worksheets.get_Item(left);
                        objWorksheet2 = objWorkbook1.Worksheets.get_Item(right);
                        foreach (Excel.Range cell in objWorksheet1.UsedRange)
                        {
                            string changedCell = cell.get_Address();
                            bool result = false;
                            if (trim && ignore) result = !(String.Equals(cell.Text.Trim(), objWorksheet2.get_Range(changedCell).Text.Trim(), StringComparison.OrdinalIgnoreCase));
                            else if (trim) result = !(String.Equals(cell.Text.Trim(), objWorksheet2.get_Range(changedCell).Text.Trim()));
                            else if (ignore) result = !(String.Equals(cell.Text, objWorksheet2.get_Range(changedCell).Text, StringComparison.OrdinalIgnoreCase));
                            else result = !(String.Equals(cell.Text, objWorksheet2.get_Range(changedCell).Text));
                            if (result)
                            {
                                if (highlight) cell.Interior.ColorIndex = 3;
                                if (mode == 2)
                                {
                                    sw.WriteLine(" <result id=\"" + i + "\">");
                                    sw.WriteLine("  <leftSheetName>" + objWorkbook1.Worksheets.get_Item(left).Name + "</leftSheetName>");
                                    sw.WriteLine("  <rightSheetName>" + objWorkbook1.Worksheets.get_Item(right).Name + "</rightSheetName>");
                                    sw.WriteLine("  <row>" + cell.Row + "</row>");
                                    sw.WriteLine("  <column>" + cell.Column + "</column>");
                                    sw.WriteLine("  <cell>" + cell.get_Address() + "</cell>");
                                    sw.WriteLine("  <newValue><![CDATA[" + cell.Text + "]]></newValue>");
                                    sw.WriteLine("  <oldValue><![CDATA[" + objWorksheet2.get_Range(changedCell).Text + "]]></oldValue>");
                                    sw.WriteLine(" </result>");
                                    i++;
                                }
                                if (mode == 1) sw.WriteLine("<tr><td>" + changedCell + "</td><td>" + cell.Text + "</td><td>" + objWorksheet2.get_Range(changedCell).Text + "</td></tr>");
                            }
                        if (mode == 1) sw.WriteLine("</table>");

                    }
                    if (mode == 1) sw.WriteLine("</cener></body></html>");
                    if (mode == 2) sw.WriteLine("</ExcelCompare>");
                    if (highlight)
                    {
                        objWorkbook1.Save();
                    }
                    else
                    {
                        objWorkbook1.Close(true, null, null);
                    }
                    objExcel.Quit();
                }
            }
            catch (System.IO.IOException Ex)
            {
                status = 55;
                //string exception = "";
                //Match m = Regex.Match(Ex.ToString(), "^(.*)", RegexOptions.Multiline);
                //if (m.Success)
                //    exception = m.Groups[0].Value;
                //Console.WriteLine("IO Exceception." + exception);
                MessageBox.Show("IO Exceception." + Ex.Message);
                objExcel.Quit();
            }
            catch (Exception Ex)
            {
                status = 99;
                //string exception = "";
                //Match m = Regex.Match(Ex.ToString(), "^(.*)", RegexOptions.Multiline);
                //if (m.Success)
                //    exception = m.Groups[0].Value;
                //MessageBox.Show("Unknow Exceception." + exception);
                MessageBox.Show("Unknow Exceception." + Ex.Message);
                objExcel.Quit();
            }
            objExcel.Quit();
            return status;
        }
        public static int CompareExcel(string oldFile, string newFile, string outputFile)
        {
            Excel.Application objExcel;
            Excel.Workbook objWorkbook1, objWorkbook2;
            Excel.Worksheet objWorksheet1, objWorksheet2;
            objExcel = new Excel.Application();
            int WScount1, WScount2;
            objExcel.Visible = false;
            objExcel.DisplayAlerts = false;
            int status = 0;
          
                    if (!File.Exists(oldFile))
                    {
                        status = 1;
                        MessageBox.Show("Please select old Excel File.\n " + oldFile + " Not Found");
                        return status;
                    }
                    if (!File.Exists(newFile))
                    {
                        status = 2;
                        MessageBox.Show("Please select new Excel File.\n " + newFile + " Not Found");
                        return status;
                    }
                    if (oldFile == newFile)
                    {
                        status = 5;
                        MessageBox.Show("You Crazy! Please select different files");
                        return status;
                    }
                    oldFile = Path.GetFullPath(oldFile);
                    newFile = Path.GetFullPath(newFile);
                    try
                    {
                    objWorkbook1 = objExcel.Workbooks.Open(newFile);
                    objWorkbook2 = objExcel.Workbooks.Open(oldFile);
                    }
                    catch (Exception ex) 
                    {
                        status=100;
                        MessageBox.Show("This application need MS office to be installed.\nIf Office is already installed please mail below file to us.\n"+outputFile);
                        using (StreamWriter sw = new StreamWriter(outputFile))
                        {
                            sw.WriteLine( DateTime.Now.ToString("MM-dd-yyyy HH:mm ss tt :") + ex);
                        }
                        return status;
                    }
                    try
                    {
                      using (StreamWriter sw = new StreamWriter(outputFile))
                      {
                        // Excel.Range range ;
                        if (mode == 1) sw.WriteLine("<!DOCTYPE HTML><html><head><title>Excel Compare Reslt</title><style>th,td{word-wrap: break-word;max-width: 300px;} body{margin:0px;font:12px verdana;} table,tr,th,td{font:12px verdana;border:1px solid black;border-collapse:collapse;padding:5px;text-align:center} th{color:white;background:green;}</style></head><body><center>");
                        if (mode == 2)
                        {
                            sw.WriteLine("<?xml version=\"1.0\"?>");
                            sw.WriteLine("<ExcelCompare>");
                            sw.WriteLine(" <info>");
                            sw.WriteLine("  <date>" + DateTime.Now.ToString("MM-dd-yyyy HH:mm ss tt :") + "</date>");
                            sw.WriteLine("  <NewExcelPath>" + newFile + "</NewExcelPath>");
                            sw.WriteLine("  <OldExcelPath>" + oldFile + "</OldExcelPath>");
                            sw.WriteLine("  <OutputMode>XML</OutputMode>");
                            sw.WriteLine("  <OutputPath>" + outputFile + "</OutputPath>");
                            sw.WriteLine(" </info>");
                        }
                        WScount1 = objWorkbook1.Worksheets.Count;
                        WScount2 = objWorkbook2.Worksheets.Count;
                      if (!sheetIgnore)
                        {
                            if (WScount1 != WScount2)
                            {
                                status = 100;
                                objWorkbook1.Close(true, null, null);
                                objWorkbook2.Close(true, null, null);
                                objExcel.Quit();
                                MessageBox.Show("Similar sheets can only be compared.\nIf you want to compare anyway, change ignore sheet count option in settings");
                                return status;
                            }
                        }
                        else
                        {
                            WScount1 = (WScount1 < WScount2) ? WScount1 : WScount2;
                        }
                        for (int i = 1, id = 1; i <= WScount1; i++)
                        {
                            string Name = objWorkbook1.Worksheets.get_Item(i).Name;
                            if (mode == 1) sw.WriteLine("<h1>" + Name + "</h1><table border='1'><tr><th>Location</th><th>Old File <br>" + Path.GetFileName(oldFile) + "</th><th>New File <br>" + Path.GetFileName(newFile) + "</th></tr>");

                            objWorksheet1 = objWorkbook1.Worksheets.get_Item(i);
                            objWorksheet2 = objWorkbook2.Worksheets.get_Item(i);

                            //int rows = (objWorksheet1.Rows.Count >= objWorksheet2.Rows.Count) ? objWorksheet1.Rows.Count : objWorksheet2.Rows.Count;
                            //int cols = (objWorksheet1.Columns.Count >= objWorksheet2.Columns.Count) ? objWorksheet1.Columns.Count : objWorksheet2.Columns.Count;
                            //Excel.Range cells = objWorksheet1.get_Range(objWorksheet1.Cells[1, 1], objWorksheet1.Cells[rows,cols]);

                            //object excelObject1,excelObect2;
                            //if (objWorksheet1.UsedRange.Rows.Count * objWorksheet1.UsedRange.Columns.Count > objWorksheet1.UsedRange.Rows.Count * objWorksheet1.UsedRange.Columns.Count) { }
                            foreach (Excel.Range cell in objWorksheet1.UsedRange)
                            {

                                string changedCell = cell.get_Address();
                                bool result = false;
                                if (trim && ignore) result = !(String.Equals(cell.Text.Trim(), objWorksheet2.get_Range(changedCell).Text.Trim(), StringComparison.OrdinalIgnoreCase));
                                else if (trim) result = !(String.Equals(cell.Text.Trim(), objWorksheet2.get_Range(changedCell).Text.Trim()));
                                else if (ignore) result = !(String.Equals(cell.Text, objWorksheet2.get_Range(changedCell).Text, StringComparison.OrdinalIgnoreCase));
                                else result = !(String.Equals(cell.Text, objWorksheet2.get_Range(changedCell).Text));
                                if (result)
                                {
                                    if(highlight) cell.Interior.ColorIndex = 3;
                                    if (mode == 2)
                                    {
                                        sw.WriteLine(" <result id=\"" + id + "\">");
                                        sw.WriteLine("  <newFleSheetName>"+objWorkbook1.Worksheets.get_Item(i).Name+"</newFleSheetName>");
                                        sw.WriteLine("  <oldFleSheetName>"+objWorkbook2.Worksheets.get_Item(i).Name+"</oldFleSheetName>");
                                        sw.WriteLine("  <row>" + cell.Row + "</row>");
                                        sw.WriteLine("  <column>" + cell.Column + "</column>");
                                        sw.WriteLine("  <cell>" + cell.get_Address() + "</cell>");
                                        sw.WriteLine("  <newValue><![CDATA[" + cell.Text + "]]></newValue>");
                                        sw.WriteLine("  <oldValue><![CDATA[" + objWorksheet2.get_Range(changedCell).Text + "]]></oldValue>");
                                        sw.WriteLine(" </result>");
                                        id++;
                                    }
                                    if (mode == 1) sw.WriteLine("<tr><td>" + changedCell + "</td><td>" + cell.Text + "</td><td>" + objWorksheet2.get_Range(changedCell).Text + "</td></tr>");
                                }
                                //else if(highlight)
                                //{
                                //        cell.Interior.ColorIndex = 0; 
                                //}
                            }
                            if (mode == 1) sw.WriteLine("</table>");

                        }
                        if (mode == 1) sw.WriteLine("</cener></body></html>");
                        if (mode == 2) sw.WriteLine("</ExcelCompare>");
                        if (highlight)
                        {
                            objWorkbook1.Save();
                            objWorkbook2.Save();
                        }
                        else
                        {
                            objWorkbook1.Close(true, null, null);
                            objWorkbook2.Close(true, null, null);
                        }
                          objExcel.Quit();
                    }
                }
                catch (System.IO.IOException Ex)
                {
                    status = 55;
                    //string exception = "";
                    //Match m = Regex.Match(Ex.ToString(), "^(.*)", RegexOptions.Multiline);
                    //if (m.Success)
                    //    exception = m.Groups[0].Value;
                    //Console.WriteLine("IO Exceception." + exception);
                    MessageBox.Show("IO Exceception." + Ex.Message);
                    objExcel.Quit();
                }
                catch (Exception Ex)
                {
                    status = 99;
                    //string exception = "";
                    //Match m = Regex.Match(Ex.ToString(), "^(.*)", RegexOptions.Multiline);
                    //if (m.Success)
                    //    exception = m.Groups[0].Value;
                    //MessageBox.Show("Unknow Exceception." + exception);
                    MessageBox.Show("Unknow Exceception." + Ex.Message);
                    objExcel.Quit();
                }
            objExcel.Quit();
            return status;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (mode == 1) this.saveFileDialog1.DefaultExt = ".html";
            if (mode == 2) this.saveFileDialog1.DefaultExt = ".xml";
            DialogResult result = saveFileDialog1.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                string name = saveFileDialog1.FileName;
                if (File.Exists(name)) { MessageBox.Show("File " + name + " already exist.\nOverwriting the file....."); }
                try
                {
                    File.Copy(outputFile, name, true);
                }
                catch (Exception Ex)
                {
                    //string exception = "";
                    //Match m = Regex.Match(Ex.ToString(), "^(.*)", RegexOptions.Multiline);
                    //if (m.Success)
                    //    exception = m.Groups[0].Value;
                    //MessageBox.Show("Unknow Exceception." + exception);
                    MessageBox.Show("Unknow Exceception." + Ex.Message);
                }
            }
        }


        private void button7_Click(object sender, EventArgs e)
        {
            if (mode == 1) this.saveFileDialog1.DefaultExt = ".html";
            if (mode == 2) this.saveFileDialog1.DefaultExt = ".xml";
            DialogResult result = saveFileDialog1.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                string name = saveFileDialog1.FileName;
                if (File.Exists(name)) { MessageBox.Show("File " + name + " already exist.\nOverwriting the file....."); }
                try
                {
                    File.Copy(outputFile, name, true);
                }
                catch (Exception Ex)
                {
                    //string exception = "";
                    //Match m = Regex.Match(Ex.ToString(), "^(.*)", RegexOptions.Multiline);
                    //if (m.Success)
                    //    exception = m.Groups[0].Value;
                    //MessageBox.Show("Unknow Exceception." + exception);
                    MessageBox.Show("Unknow Exceception." + Ex.Message);
                }
            }
        }
        private void sheetCompareToolStripMenuItem_Click(object sender, EventArgs e)
        {
            panel1.Visible = false; panel2.Visible = false; panel3.Visible = true; panel4.Visible = false;
        }
        private void fileCompareToolStripMenuItem_Click(object sender, EventArgs e)
        {
            panel1.Visible = true; panel2.Visible = false; panel3.Visible = false; panel4.Visible = false;
        }
        private void settingsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            panel1.Visible = false; panel2.Visible = false; panel3.Visible = false; panel4.Visible = true;
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Excel Compare for MS Office\nProduct of Ramu Creations\nwww.webapps-tricks.com");
        }

        private void helpToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.webapps-tricks.com/rc/downloadcenter/windows/Excel-Compare-for-MSOffice/");
        }

        private void donateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("http://webapps-tricks.com/rc/about/donate.php");
        }

        private void reportBugToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Please Raise a ticket at the link that is opened.\nWe will resolve the bug as early as possible.Thank you");
            System.Diagnostics.Process.Start("https://sourceforge.net/p/excelcompare-msoffice/tickets/");
        }

        private void tellAFriendToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process proc = new System.Diagnostics.Process();
            proc.StartInfo.FileName = "mailto:?subject=Try Excel Compare Software&body=I am using Excel Compare Software. It will be used for comparing excel files very easily";
            proc.Start();
        }



    }
}

