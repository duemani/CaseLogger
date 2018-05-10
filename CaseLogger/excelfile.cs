using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseLogger
{
    public class excelfile
    {
        private Excel.Application excel_App;
        private Excel.Workbook excel_Wkbk;
        private Excel.Worksheet excel_Wkst;
        private object misvalue = System.Reflection.Missing.Value;
        public string filename; 

        public excelfile()
        {
            
        }

        //constructor for a old patient
        public excelfile(string path)
        {
            excel_App = new Excel.Application();
            filename = path;
            if (!(File.Exists(filename)))
            {
                excel_Wkbk = excel_App.Workbooks.Add();
            }
            else
            {
                try
                {
                    excel_Wkbk = excel_App.Workbooks.Open(path, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "\t",false, false, 0, true, 1, 0);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error opening excel file" + Environment.NewLine + ex.Message);
                }
                excel_Wkst = (Excel.Worksheet)excel_Wkbk.Worksheets.get_Item(1);
            }
        }

        //write to a cell
        public bool WriteCell(string celladdress, string value)
        {
            try
            {
                Excel.Range range = excel_Wkst.get_Range(celladdress, celladdress);
                range.set_Value(misvalue,value);
            }
            catch
            {
                return false;
            }
            return true;
        }

        //write to more than one cell
        public bool WriteCells_multiple(string celladdress_beg, string celladdress_end, string value)
        {
            try
            {
                Excel.Range range = excel_Wkst.get_Range(celladdress_beg, celladdress_end);
                range.set_Value(misvalue,value);
            }
            catch
            {
                return false;
            }
            return true;
        }

        //get value from a cell
        public string GetCellValue (string address)
        {
            string cellvalue;
            try
            {
                cellvalue = excel_Wkst.get_Range(address, address).Value2.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error getting cell" + address + Environment.NewLine + ex.Message);
                return "";
            }
            return cellvalue;
        }

        public string GetCellValueText(string address)
        {
            string cellvalue;
            try
            {
                cellvalue = excel_Wkst.get_Range(address, address).Text.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error getting cell" + address + Environment.NewLine + ex.Message);
                return "";
            }
            return cellvalue;
        }

        //get values from more than one cell, NOT WORKING CORRECTLY
        public string GetCellValues_multiple (string address_beg, string address_end)
        {
            string cellvalues;
            try
            {
                cellvalues = excel_Wkst.get_Range(address_beg, address_end).Value2.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error getting cells" + address_beg + "-" + address_end + Environment.NewLine + ex.Message);
                return "";
            }
            return cellvalues;
        }

        public int getlastrow()
        {
            Excel.Range last = excel_Wkst.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            Excel.Range range = excel_Wkst.get_Range("A1", last);
            int lastrow = last.Row;
            return lastrow;
        }


        public void savefile()
        {
            excel_Wkbk.SaveAs(filename);
         
        }

        //close file
        public void closefile()
        {
            excel_Wkbk.Close(true, misvalue, misvalue);
            excel_App.Quit();          
            releaseObject(excel_Wkst);
            releaseObject(excel_Wkbk);
            releaseObject(excel_App);
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
         //       MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        } 
    }
}