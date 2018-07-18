using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel; 


namespace Excel_Export
{
    public partial class Form1 : Form
    {


        Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;

        String connectionString = "Data Source=AZIZ-PC;Initial Catalog=aziz;Integrated Security=True";

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            
      try
            {

                for (int i = 1; i < 100; i++)
                {              
                    object misValue = System.Reflection.Missing.Value;
                    xlWorkBook = xlApp.Workbooks.Add(misValue);
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);



                    if (xlApp == null)
                    {
                        MessageBox.Show("Excel is not properly installed!!");
                        return;
                    }


                   
                        xlWorkSheet.Cells[1, 1] = "Kuvet";
                        xlWorkSheet.Cells[1, 2] = "Position";
                        xlWorkSheet.Cells[1, 3] = "Olcum Zamani";


                        //Add values for Kuvet only
                        xlWorkSheet.Cells[2, 1] = "50";
                        xlWorkSheet.Cells[3, 1] = "5";
                        xlWorkSheet.Cells[4, 1] = "25";


                        xlWorkSheet.Cells[2, 2] = "40";
                        xlWorkSheet.Cells[3, 2] = "50";
                        xlWorkSheet.Cells[4, 2] = "60";


                        xlWorkSheet.Cells[2, 3] = "40";
                        xlWorkSheet.Cells[3, 3] = "50";
                        xlWorkSheet.Cells[4, 3] = "60";



                        xlWorkBook.SaveAs("c:\\excel\\csharp-Excel"+i+".xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                        xlWorkBook.Close(true, misValue, misValue);




                        xlApp.Quit();

                    }
  
                 //MessageBox.Show("Excel file created , you can find the file c:\\csharp-Excel.xls");

                             

            }
                catch(Exception ex)
                 {
                    MessageBox.Show(ex.Message);

                }

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

        protected override void OnClosing(CancelEventArgs e)
        {


            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);

            base.OnClosing(e);
        }






    }
}
