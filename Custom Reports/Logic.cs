using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
//using Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace Custom_Reports
{
    class Logic
    {
        public string currnet { get; set; }
        public string PoColumn { get; set; }
        public string PoTotal { get; set; }
        public string ShippingA { get; set; }
        public string Company { get; set; }
        public string InvoiceD { get; set; }
        public string MonthYear { get; set; }
        public string InvoiceN { get; set; }
        public string PoDate { get; set; }
        public string GroupNumber { get; set; }
        public bool JoinC { get; set; }


        public void MemberReconcile()
        {

            // imports data from source file using connection calas and then finds the matching values.


            Excel.Worksheet repinfo = (Excel.Worksheet)Globals.ThisAddIn.Application.Sheets["Import"];

            Excel.Range repLastRow = repinfo.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            int replasrow2 = repLastRow.Row;

            Excel.Range rng2 = repinfo.Range["I2:I" + replasrow2];

            System.Data.DataTable tbl = new System.Data.DataTable();
            Connection cn = new Connection();
            tbl = cn.Tble();


            try
            {
                foreach (DataRow vala in tbl.Rows)
            {
                foreach (Excel.Range valb in rng2.Rows)
                {
                    if (vala["VendorName"].ToString().ToUpper().Trim().Contains(valb.Value.ToString().ToUpper().Trim()))
                    {
                        valb.Offset[0, -4].Value = vala["VendorName"] + "-" + vala["ContractName"];
                        valb.Offset[0, -6].Value = vala["ContractId"].ToString();
                        valb.Offset[0, -7].Value = vala["VendorId"].ToString();

                    }

                    //if (vala["VendorName"].ToString().ToUpper().Trim().Contains(valb.Value.ToString().ToUpper().Trim()))
                    if (valb.Value.ToString().ToUpper().Trim().Contains(vala["VendorName"].ToString().ToUpper().Trim()))
                    {
                        valb.Offset[0, -4].Value = vala["VendorName"] + "-" + vala["ContractName"];
                        valb.Offset[0, -6].Value = vala["ContractId"].ToString();
                        valb.Offset[0, -7].Value = vala["VendorId"].ToString();

                    }

                }

            }

            }
            catch (Exception)
            {
                MessageBox.Show("Somthing went Wrong while Reconciling, Please dont submit the report");
                repinfo.Range["A1:H1"].Value = "Do not Submit this Report it Has Errors";
                repinfo.Range["A1:H1"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                repinfo.Range["A1:H1"].Font.Size = "19";
                repinfo.Range["A1:H1"].Font.Color = Color.Yellow;

            }


        }

        public void VendorReconcile()
        {

            Excel.Worksheet repinfo = (Excel.Worksheet)Globals.ThisAddIn.Application.Sheets["Import"];

            Excel.Range repLastRow = repinfo.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            int replasrow2 = repLastRow.Row;

            Excel.Range rng2 = repinfo.Range["B2:B" + replasrow2];

            System.Data.DataTable tbl = new System.Data.DataTable();
            Connection cn = new Connection();
            tbl = cn.TbleVenReport();




            try
            {
                foreach (DataRow vala in tbl.Rows)
            {
                foreach (Excel.Range valb in rng2.Rows)
                {
                    if (vala["Group_Name"].ToString().ToUpper().Trim().Contains(valb.Value.ToString().ToUpper().Trim()))
                    {
                        valb.Offset[0, 3].Value = vala["Group_Name"].ToString();
                        valb.Offset[0, -1].Value = vala["Group_ID"].ToString();
                        //valb.Offset[0, -7].Value = vala["VendorId"].ToString();

                    }

                        if (valb.ToString().ToUpper().Trim().Contains(vala["Group_Name"].ToString().ToUpper().Trim()))
                        {
                            valb.Offset[0, 3].Value = vala["Group_Name"].ToString();
                            valb.Offset[0, -1].Value = vala["Group_ID"].ToString();
                            //valb.Offset[0, -7].Value = vala["VendorId"].ToString();

                        }

                    }

                   

                }
            }
            catch (Exception)
            {
                repinfo.Range["A1:H1"].Value = "Do not Submit this Report it Has Errors";
                repinfo.Range["A1:H1"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                repinfo.Range["A1:H1"].Font.Size = "19";
                repinfo.Range["A1:H1"].Font.Color = Color.Yellow;
                MessageBox.Show("Somthing went wrong Please dont submit the report");
                
            }

        }
        public void Dup2()
        {
            int replasrow2 = 0;
            Excel.Worksheet repinfo = (Excel.Worksheet)Globals.ThisAddIn.Application.Sheets["Import"];

            Excel.Range repLastRow = repinfo.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            replasrow2 = repLastRow.Row;
            Excel.Range rng2 = repinfo.Range["I2:I" + replasrow2];


            try
            {


                for (var x = replasrow2; x >= 2; x += -1)
                {
                    for (var y = 2; y <= replasrow2; y++)
                    {

                        if (Convert.ToString(repinfo.Cells[x, 6].Value) == Convert.ToString(repinfo.Cells[y, 6].Value) && Convert.ToString(repinfo.Cells[x, 9].Value) == Convert.ToString(repinfo.Cells[y, 9].Value) && Convert.ToString(repinfo.Cells[x, 11].Value) == Convert.ToString(repinfo.Cells[y, 11].Value) && x > y) // this one looks into Po number Invoice number and Company name
                                                                                                                                                                                                                                                                                                                               //if (Convert.ToString(repinfo.Cells[x, 6].Value) == Convert.ToString(repinfo.Cells[y, 6].Value) && Convert.ToString(repinfo.Cells[x, 11].Value) == Convert.ToString(repinfo.Cells[y, 11].Value) && x > y) // this one was looking into PO number and Invoice number only
                        {

                            decimal kk;

                            if ((decimal.TryParse(repinfo.Cells[x, 7].Value.ToString(), out kk)))
                            {
                                repinfo.Cells[y, 7].Value = Convert.ToDecimal(repinfo.Cells[x, 7].Value) + Convert.ToDecimal(repinfo.Cells[y, 7].Value);
                                repinfo.Cells[y, 10].Value = Convert.ToDecimal(repinfo.Cells[x, 10].Value) + Convert.ToDecimal(repinfo.Cells[y, 10].Value);
                                repinfo.Rows[x].EntireRow.Delete(Type.Missing);
                            }

                            else
                            {
                                repinfo.Cells[x, 8].Value = "This number on the left is not Convertable";
                                repinfo.Cells[y, 8].Value = "This number looks good but The duplicate of this number was not convertable";
                                repinfo.Cells[x, 8].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                                repinfo.Cells[y, 8].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

                            }

                            break;
                        }
                    }
                }



            }
            catch (Exception)
            {
                repinfo.Range["A1:H1"].Value = "Do not Submit this Report it Has Errors";
                repinfo.Range["A1:H1"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                repinfo.Range["A1:H1"].Font.Size = "19";
                repinfo.Range["A1:H1"].Font.Color = Color.Yellow;

                MessageBox.Show("Somthing went wrong While removing Duplicates");

            }

        }

        public void Dup22()
        {
            // This one looks into Company name and po number
            int replasrow2 = 0;
            Excel.Worksheet repinfo = (Excel.Worksheet)Globals.ThisAddIn.Application.Sheets["Import"];

            Excel.Range repLastRow = repinfo.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            replasrow2 = repLastRow.Row;
            Excel.Range rng2 = repinfo.Range["I2:I" + replasrow2];


            try
            {


                for (var x = replasrow2; x >= 2; x += -1)
                {
                    for (var y = 2; y <= replasrow2; y++)
                    {


                        if (Convert.ToString(repinfo.Cells[x, 6].Value) == Convert.ToString(repinfo.Cells[y, 6].Value) && Convert.ToString(repinfo.Cells[x, 9].Value) == Convert.ToString(repinfo.Cells[y, 9].Value) && x > y)
                        {

                            decimal kk;

                            if ((decimal.TryParse(repinfo.Cells[x, 7].Value.ToString(), out kk)))
                            {
                                repinfo.Cells[y, 7].Value = Convert.ToDecimal(repinfo.Cells[x, 7].Value) + Convert.ToDecimal(repinfo.Cells[y, 7].Value);
                                repinfo.Cells[y, 10].Value = Convert.ToDecimal(repinfo.Cells[x, 10].Value) + Convert.ToDecimal(repinfo.Cells[y, 10].Value);
                                repinfo.Rows[x].EntireRow.Delete(Type.Missing);
                            }

                            else
                            {
                                repinfo.Cells[x, 8].Value = "This number on the left is not Convertable";
                                repinfo.Cells[y, 8].Value = "This number looks good but The duplicate of this number was not convertable";
                                repinfo.Cells[x, 8].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                                repinfo.Cells[y, 8].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);


                            }

                            break;
                        }
                    }
                }



            }
            catch (Exception)
            {

                repinfo.Range["A1:H1"].Value = "Do not Submit this Report it Has Errors";
                repinfo.Range["A1:H1"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                repinfo.Range["A1:H1"].Font.Size = "19";
                repinfo.Range["A1:H1"].Font.Color = Color.Yellow;
                MessageBox.Show("Somthing went wrong While removing Duplicates 22");

            }


        }

        public void VendorCombineCol()
        {
            int replasrow2 = 0;
            Excel.Worksheet repinfo = (Excel.Worksheet)Globals.ThisAddIn.Application.Sheets["Import"];

            Excel.Range repLastRow = repinfo.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            replasrow2 = repLastRow.Row;
            Excel.Range rng2 = repinfo.Range["C2:C" + replasrow2];
            try
            {
                if (InvoiceN == "")
                {
                    if (MonthYear == "")
                    {
                        //MessageBox.Show("Inv Empty month empty");
                        foreach (Excel.Range val in rng2)
                        {
                            val.Value = val.Value.ToString() + "-" + MonthYear;
                        }
                    }

                    else if (MonthYear != "")
                    {
                        foreach (Excel.Range val in rng2)
                        {
                            val.Value = val.Value.ToString() + "-" + MonthYear;
                        }
                    }
                }
                if (InvoiceN != "")
                {
                    //MessageBox.Show("Not empty");
                    if (MonthYear.ToString() == "")
                    {

                        foreach (Excel.Range val in rng2)
                        {
                            if (val.Offset[0, 4].Value.ToString() != "")
                            {
                                if (JoinC == false)
                                {
                                    val.Value = val.Value.ToString() + "-Inv " + val.Offset[0, 4].Value.ToString();
                                }
                                if (JoinC == true)
                                {
                                    val.Value = val.Value.ToString() + "-" + MonthYear;
                                }

                            }

                            

                        }
                    }

                    if (MonthYear != "")
                    {
                        foreach (Excel.Range val in rng2)
                        {
                            if (val.Offset[0, 4].Value.ToString() != "")
                            {
                                if (JoinC == false)
                                {
                                    val.Value = val.Value.ToString() + "-Inv " + val.Offset[0, 4].Value.ToString() + "-" + MonthYear;
                                }
                                if (JoinC == true)
                                {
                                    val.Value = val.Value.ToString() + "-" + MonthYear;
                                }
                            }
                        }
                    }
                }
            }
            catch
            {
                MessageBox.Show("Somthing went wrong Please dont submit this report");
            }
        }

        public void VendorDuplicate()
        {
            int replasrow2 = 0;
            Excel.Worksheet repinfo = (Excel.Worksheet)Globals.ThisAddIn.Application.Sheets["Import"];

            Excel.Range repLastRow = repinfo.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            replasrow2 = repLastRow.Row;
            Excel.Range rng2 = repinfo.Range["C2:C" + replasrow2];


            try
            {


                for (var x = replasrow2; x >= 2; x += -1)
                {
                    for (var y = 2; y <= replasrow2; y++)
                    {
                        // 6 is PO number 9 is company name in member re

                        if (Convert.ToString(repinfo.Cells[x, 3].Value) == Convert.ToString(repinfo.Cells[y, 3].Value) && Convert.ToString(repinfo.Cells[x, 2].Value) == Convert.ToString(repinfo.Cells[y, 2].Value) && x > y)
                        {

                            decimal kk;

                            if ((decimal.TryParse(repinfo.Cells[x, 4].Value.ToString(), out kk)))
                            {
                                repinfo.Cells[y, 4].Value = Convert.ToDecimal(repinfo.Cells[x, 4].Value) + Convert.ToDecimal(repinfo.Cells[y, 4].Value);
                                repinfo.Cells[y, 8].Value = Convert.ToDecimal(repinfo.Cells[x, 8].Value) + Convert.ToDecimal(repinfo.Cells[y, 8].Value);
                                repinfo.Rows[x].EntireRow.Delete(Type.Missing);

                                repinfo.Cells[x, 5].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                                repinfo.Cells[y, 5].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                            }

                            else
                            {
                                repinfo.Cells[x, 5].Value = "This number on the left is not Convertable";
                                repinfo.Cells[y, 5].Value = "This number looks good but The duplicate of this number was not convertable";
                                repinfo.Cells[x, 5].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                                repinfo.Cells[y, 5].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

                            }

                            break;
                        }
                    }
                }



            }
            catch (Exception)
            {

                repinfo.Range["A1:H1"].Value = "Do not Submit this Report it Has Errors";
                repinfo.Range["A1:H1"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                repinfo.Range["A1:H1"].Font.Size = "19";
                repinfo.Range["A1:H1"].Font.Color = Color.Yellow;
                repinfo.Name = "Errors in Report";
                MessageBox.Show("Somthing went wrong While removing Duplicates 22");

            }

        }
    }
}
    
