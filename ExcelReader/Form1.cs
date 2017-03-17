using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelReader
{
    public partial class Form1 : Form
    {
        public static String file_to_open;

        public static Excel.Application xlApp = null;
        public static Excel.Workbook xlWorkbook = null;
        public static Excel.Range xlRange = null;
        public static Excel._Worksheet xlWorksheet = null;
        public static int start_row = 1;
        public static int col_count = 10;

        object saveas = 5;

        public Form1()
        {
            InitializeComponent();
            xlApp = new Excel.Application();
        }

        public static int GetColumnNumber(string name)
        {
            int number = 0;
            int pow = 1;
            for (int i = name.Length - 1; i >= 0; i--)
            {
                number += (name[i] - 'A' + 1) * pow;
                pow *= 26;
            }

            return number;
        }

        public static string ReadDropDownValues(Excel._Worksheet xlWorkSheet, Excel.Range dropDownCell, int orig_col_num, bool dict_format)
        {
            string result = "";
            string col_name = "";
            int col_num = 0;
            int row_num = 0;

            try
            {
                result = dropDownCell.Validation.Formula1;

                if (result == "")
                    return result;

                if (result.Contains("INDIRECT"))
                {
                    col_name = result.Split('$')[1];
                    col_num = GetColumnNumber(col_name);
                    row_num = Int32.Parse(result.Split('$')[2].TrimEnd(')'));

                    result = ReadDropDownValues(xlWorkSheet, xlWorkSheet.UsedRange.Cells[row_num, col_num], col_num, true);
                }
                else
                {
                    if (result.Contains("=h_"))
                    {
                        Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[result.TrimStart('=')];
                        Excel.Range xlRange = xlWorksheet.UsedRange;

                        int rowCount = xlRange.Rows.Count;
                        int colCount = xlRange.Columns.Count;

                        if (dict_format)
                        {
                            result = "INDIRECT:" + "C" + orig_col_num.ToString() + ",";

                            //Writing dictionary for INDIRECT reference, flanked by two slashes.
                            result += "/" + "{";

                            //We're assuming that the format is first column in every row, is followed by its corresponding values.
                            for (int i = 1; i <= rowCount; i++)
                            {
                                result += "'" + xlRange.Cells[i, 1].Value2.ToString() + "'" + ":[";
                                for (int j = 2; j <= colCount; j++)
                                {
                                    if (xlRange.Cells[i, j].Value2 != null)
                                        result += "'" + xlRange.Cells[i, j].Value2.ToString() + "'" + ",";
                                }
                                result = result.TrimEnd(',') + "],";
                            }
                            result = result.TrimEnd(',') + "}/";
                        }
                        else
                        {
                            result = "CHOICE:";

                            //We're assuming that the data is stored in the first column, across multiple rows.
                            for (int i = 1; i <= rowCount; i++)
                            {
                                result += xlRange.Cells[i, 1].Value2.ToString() + ",";
                            }
                            result = result.TrimEnd(',');
                        }
                    }
                    else //items separated by commas
                    {
                        string temp = "CHOICE:";
                        foreach (string each in result.Split(','))
                        {
                            temp += "'" + each + "',";
                        }
                        result = temp.TrimEnd(',');
                    }
                }
            }
            catch (Exception Ex)
            {
            }

            return result;
        }

        public static string getExcelFile(string file_to_open, string config_file)
        {
            StreamWriter writer = new StreamWriter(config_file);

            List<String> headers = new List<String>();
            List<String> values = new List<String>();

            try
            {
                xlWorkbook = xlApp.Workbooks.Open(file_to_open);

                xlWorksheet = xlWorkbook.Sheets[1];
                xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;

                String cell_value = "", cell_address = "", cell_address_header = "";

                int counter = 1;

                for (int i = 1; i <= rowCount; i++)
                {
                    values.Add("\n");
                    headers.Add("\n");
                    for (int j = 1; j <= colCount; j++)
                    {
                        if (xlRange.Cells[i, j].Value2 != null)
                        {
                            cell_value = xlRange.Cells[i, j].Value2.ToString();
                            cell_address = xlRange.Cells[i, j].AddressLocal(true, true, Excel.XlReferenceStyle.xlR1C1, null, null);

                            var row_replaced = new StringBuilder(cell_address);
                            row_replaced.Remove(1, 1);
                            row_replaced.Insert(1, "0");
                            cell_address_header = row_replaced.ToString();

                            values.Add("\n");
                            values.Add(cell_address + "|" + cell_value + "|");
                            values.Add(counter.ToString());

                            headers.Add("\n");
                            headers.Add(cell_address_header + "|" + cell_value + "|");
                            headers.Add(counter.ToString());

                            counter = 1;

                            //Read the data validation (and print if it exists)
                            string Values = ReadDropDownValues(xlWorksheet, xlRange.Cells[i + 1, j], j, false);

                            values.Add("|");
                            headers.Add("|");

                            if (Values != "")
                                values.Add("(" + Values.ToString() + ")");

                            headers.Add(cell_value);

                            values.Add("\n");
                            headers.Add("\n");
                        }
                        else
                            counter += 1;
                    }
                }

                writer.WriteLine("#Headers");
                write_list(writer, headers);

                writer.WriteLine();
                writer.WriteLine("#Values");
                write_list(writer, values);

                return "Configuration file creation completed!";
            }
            catch (Exception Ex)
            {
                return Ex.Message;
            }
            finally
            {
                writer.Close();

                cleanup();
            }
        }

        private void BrowseTemplate_Click(object sender, EventArgs e)
        {
            int row_count = int.Parse(rowcount.Text);
            string config_file = "";

            DialogResult result_config = DialogResult.Cancel;
            DialogResult result_populate = DialogResult.No;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                file_to_open = openFileDialog1.FileName;

                string python_path = @"C:\Python27\python.exe";
                string path = Path.GetDirectoryName(file_to_open);
                string main = "";

                if (create_config.Checked)
                    config_file = Path.ChangeExtension(file_to_open, "txt");
                else
                    config_file = file_to_open;

                string output_file = path + @"\Template_latest.xls";

                if (!create_config.Checked)
                    result_config = MessageBox.Show("You have chosen to NOT create text template.\nMake sure, it already exists in the path, before proceeding!", "Configuration file", MessageBoxButtons.OKCancel);
                else
                {
                    xlWorkbook = xlApp.Workbooks.Open(file_to_open);
                    toolStripStatusLabel1.Text = "Creating configuration file...";
                    toolStripStatusLabel1.Text = getExcelFile(file_to_open, config_file);
                    result_populate = MessageBox.Show("Modify the configuration if required\nand hit 'Yes' to continue populating the template!", "Configuration file", MessageBoxButtons.YesNo);
                }

                if (result_config == DialogResult.OK || result_populate == DialogResult.Yes)
                {
                    string arguments = " --config \"" + config_file + "\" --rowcount " + row_count.ToString() + " --colcount " + col_count.ToString() + " --startrow " + start_row.ToString() + " --output \"" + output_file.ToString() + "\"";

                    toolStripStatusLabel1.Text = "Populating Excel using config.txt...";

                    Process process = new Process();
                    try
                    {
                        main = Application.StartupPath + @"\ExcelWriter\ExcelWriter.exe";
                        process.StartInfo.FileName = main;
                        process.StartInfo.Arguments = arguments;
                        process.StartInfo.UseShellExecute = true;
                        process.Start();
                    }
                    catch
                    {
                        //Executes main.py in python, if ExcelWriter.exe doesn't exist
                        main = path + @"\ExcelWriter\main.py";
                        process.StartInfo.FileName = python_path;
                        process.StartInfo.Arguments = main + arguments;
                        process.StartInfo.UseShellExecute = true;
                        process.Start();
                    }

                    toolStripStatusLabel1.Text = "Rename Template_latest.xls to avoid overwriting next time!";
                }
            }
        }

        private void KillExcel_Click(object sender, EventArgs e)
        {
            foreach (var process in Process.GetProcessesByName("EXCEL"))
            {
                process.Kill();
            }
        }

        private static void write_list(StreamWriter writer, List<String> list)
        {
            foreach (string line in list)
            {
                if (line == "\n")
                    writer.WriteLine();
                else
                    writer.Write(line);
            }
        }
        private static void cleanup ()
        {
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            try
            {
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);
            }
            catch (Exception Ex)
            {
                //nothing
            }
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }

        private void excelMapToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Configuration Configuration = new Configuration();
            Configuration.Show();
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutBox AboutBox = new AboutBox();
            AboutBox.Show();
        }

        private void documentationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process.Start(Application.StartupPath + "\\Help\\readme.mht");
        }

        private void create_config_Click(object sender, EventArgs e)
        {
            if (create_config.Checked)
            {
                BrowseTemplate.Text = "Browse Template.xls";
                openFileDialog1.FileName = "Template.xls";
            }
            else
            {
                BrowseTemplate.Text = "Browse Template.txt";
                openFileDialog1.FileName = "Template.txt";
            }
        }

        private void createTemplateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "Text file|*.txt|Excel file|*.xls";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if (Path.GetExtension (saveFileDialog1.FileName) == ".xls")
                {
                    xlWorkbook = xlApp.Workbooks.Add();
                    xlWorkbook.SaveAs(saveFileDialog1.FileName, Excel.XlFileFormat.xlExcel8);

                    cleanup();
                }
                if (Path.GetExtension(saveFileDialog1.FileName) == ".txt")
                {
                    File.Create(saveFileDialog1.FileName);
                }
            }
        }

        private void openTemplateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Text file|*.txt|Excel file|*.xls";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Process.Start(openFileDialog1.FileName);
            }
        }
    }
}
