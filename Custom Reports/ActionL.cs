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
    class ActionL
    {


        public string FirstSheet
        {
            get; set;
        }
        public string Currnet { get; set; }
        public string PoColumn { get; set; }
        public string PoTotal { get; set; }
        public string ShippingA { get; set; }
        public string Company { get; set; }
        public string InvoiceD { get; set; }
        public string MonthYear { get; set; }
        public string InvoiceN { get; set; }
        public string PoDate { get; set; }
        public string GroupNumber { get; set; }
        private int i2;
        public string Txbox { get; set; }
        public string txbox1 { get; set; }
        public decimal result { get; set; }

        public void Copy()
        {
            try
            {



                Worksheet currnet = Globals.ThisAddIn.Application.ActiveSheet;
                FirstSheet = currnet.Name.ToString();
                Excel.Range lastCell2 = currnet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
                int lasrow2 = lastCell2.Row;
                MessageBox.Show(lasrow2.ToString());
                Reports rt = new Reports();
             


                Excel.Application excelAPP = new Excel.Application();
                //Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add();
                Excel.Worksheet newWorksheet;
                newWorksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets.Add();
                newWorksheet.Name = "Import";
                newWorksheet.Range["A1"].Value = "Group";
                newWorksheet.Range["B1"].Value = "Vendor";
                newWorksheet.Range["C1"].Value = "Contract";
                newWorksheet.Range["D1"].Value = "Description";
                newWorksheet.Range["E1"].Value = "Check and Delete";
                newWorksheet.Range["F1"].Value = "PO number";
                newWorksheet.Range["G1"].Value = "Total";
                newWorksheet.Range["H1"].Value = "Cusomer Number";
                newWorksheet.Range["I1"].Value = "Company";
                newWorksheet.Range["J1"].Value = "Shipping Amount";
                newWorksheet.Range["K1"].Value = "Invoice Number";
                newWorksheet.Range["L1"].Value = "PO Date";
                newWorksheet.Range["M1"].Value = "Invoice Date";
                //string ponumber = PoColumn;
                string range1 = PoColumn + "2:" + PoColumn + lasrow2;
                string range2 = "F2:" + "F" + lasrow2;

                // range for Total
                string rngtotal1 = PoTotal + "2:" + PoTotal + lasrow2;
                string rngTotal2 = "G2:" + "G" + lasrow2;
                //currnet.Copy(currnet.Range[range1], newWorksheet.Range["F2"]) ;

                Excel.Range from = currnet.Range[range1];
                Excel.Range to = newWorksheet.Range[range2];

                Excel.Range fromtotal = currnet.Range[rngtotal1];
                Excel.Range tototal = newWorksheet.Range[rngTotal2];

                // company
                string rngcomp1 = Company + "2:" + Company + lasrow2;
                string rngcomp2 = "I2:" + "I" + lasrow2;

                Excel.Range fromComp = currnet.Range[rngcomp1];
                Excel.Range toComp = newWorksheet.Range[rngcomp2];

                //Shipping

                string rngship1 = ShippingA + "2:" + ShippingA + lasrow2;
                string rngship2 = "J2:" + "J" + lasrow2;

                Excel.Range fromship = currnet.Range[rngship1];
                Excel.Range toship = newWorksheet.Range[rngship2];

                //Invoice

                string rnginv1 = InvoiceN + "2:" + InvoiceN + lasrow2;
                string rnginv2 = "K2:" + "K" + lasrow2;

                Excel.Range frominv = currnet.Range[rnginv1];
                Excel.Range toiv = newWorksheet.Range[rnginv2];



            try
            {
                from.Copy(to);
                fromtotal.Copy(tototal);
                    fromComp.Copy(toComp);
                    if (ShippingA != "")
                    {
                        fromship.Copy(toship);
                    }
                    if (InvoiceN != "")
                     {
                        frominv.Copy(toiv);
                     }

        }
                catch (Exception)
                {
                    newWorksheet.Range["A1:H1"].Value = "Do not Submit this Report it Has Errors";
                    newWorksheet.Range["A1:H1"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    newWorksheet.Range["A1:H1"].Font.Size = "19";
                    newWorksheet.Range["A1:H1"].Font.Color = Color.Yellow;

                    MessageBox.Show("Somthing went wrong with copy, please try again");
                }

}
            catch (Exception)
            {

                MessageBox.Show("Somthing went Wrong Please check the Sheet names");
            }

        }

        public void cleanColumn(string CheckColumn, string ColumnName)
        {
            //this method will pin point other not convertable cells which dont have any duplicate.
            int replasrow2 = 0;
            Worksheet reporiginal = Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Range repLastRow = reporiginal.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);

            replasrow2 = repLastRow.Row;

            Excel.Range rng2 = reporiginal.Range[CheckColumn + "2:" + CheckColumn + replasrow2.ToString()];

            i2 = 0;

            if (CheckColumn != "")
            {

                foreach (Excel.Range vala in rng2)
                {
                    decimal kk;

                    //decimal shwo ;

                    if (vala.Cells.Value == null)
                    {
                        vala.Cells.Value = "0";
                        vala.Cells.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                    }

                    else if (decimal.TryParse(vala.Cells.Value.ToString(), out kk))

                    {
                        vala.Cells.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                    }
                    else
                    {
                        i2++;
                        vala.Cells.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

                    }

                }
            }

            
         
            if (i2 > 0 && ColumnName == "PoTotal" )
            {

                MessageBox.Show("The value in the RED are not convertable to number please make corrections and try again", i2.ToString() + " Errors on PO Total");
                //MessageBox.Show(i2.ToString() + "Errors - The value in the RED are not convertable to number please make corrections and try again");
                Txbox = "No";
            }
           else if (i2 > 0 && ColumnName == "ShipingA")
            {

                MessageBox.Show("The value in the RED are not convertable to number please make corrections and try again", i2.ToString() + " Errors on Shipping Amounn");
                //MessageBox.Show(i2.ToString() + "Errors - The value in the RED are not convertable to number please make corrections and try again");
                txbox1 = "No";
            }


        }


        public void Vcopy()
        {
            try
            {
                Worksheet currnet = Globals.ThisAddIn.Application.ActiveSheet;
                string OriginalSheet = currnet.Name.ToString();
                Excel.Range lastCell2 = currnet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
                int lasrow2 = lastCell2.Row;
                MessageBox.Show(lasrow2.ToString());

                Excel.Application excelAPP = new Excel.Application();
                //Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add();
                Excel.Worksheet newWorksheet;
                newWorksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets.Add();
                newWorksheet.Name = "Import";
                newWorksheet.Range["A1"].Value = "Group";
                newWorksheet.Range["B1"].Value = "Description";
                newWorksheet.Range["C1"].Value = "PO Number";
                newWorksheet.Range["D1"].Value = "Total";
                newWorksheet.Range["E1"].Value = "Check and Delete";
                newWorksheet.Range["F1"].Value = "Customer Number";
                newWorksheet.Range["G1"].Value = "Invoice Number";
                newWorksheet.Range["H1"].Value = "Shipping Amount";
                newWorksheet.Range["I1"].Value = "PO Date";
                newWorksheet.Range["J1"].Value = "Invoice Date";

                //string ponumber = PoNumber;
                string range1 = PoColumn + "2:" + PoColumn + lasrow2;
                string range2 = "C2:" + "C" + lasrow2;

                Excel.Range fromPoNum = currnet.Range[range1];
                Excel.Range toPoNum = newWorksheet.Range[range2];

                // range for Total
                string rngtotal1 = PoTotal + "2:" + PoTotal + lasrow2;
                string rngTotal2 = "D2:" + "D" + lasrow2;
                //currnet.Copy(currnet.Range[range1], newWorksheet.Range["F2"]) ;

                Excel.Range fromtotal = currnet.Range[rngtotal1];
                Excel.Range tototal = newWorksheet.Range[rngTotal2];

                // company
                string rngcomp1 = Company + "2:" + Company + lasrow2;
                string rngcomp2 = "B2:" + "B" + lasrow2;

                Excel.Range fromComp = currnet.Range[rngcomp1];
                Excel.Range toComp = newWorksheet.Range[rngcomp2];

                //Shipping

                string rngship1 = ShippingA + "2:" + ShippingA + lasrow2;
                string rngship2 = "H2:" + "H" + lasrow2;

                Excel.Range fromship = currnet.Range[rngship1];
                Excel.Range toship = newWorksheet.Range[rngship2];

                //Invoice

                string rnginv1 = InvoiceN + "2:" + InvoiceN + lasrow2;
                string rnginv2 = "G2:" + "G" + lasrow2;

                Excel.Range frominv = currnet.Range[rnginv1];
                Excel.Range toiv = newWorksheet.Range[rnginv2];

                //Po Daate

                string ranPdate = PoDate + "2:" + PoDate + lasrow2;
                string rangPdate2 = "I2:" + "I" + lasrow2;

                Excel.Range fromPoDate = currnet.Range[ranPdate];
                Excel.Range toPOdate = newWorksheet.Range[rangPdate2];

                //Invoice date
                string InvDate1 = InvoiceD + "2:" + InvoiceD + lasrow2;
                string InvDate2 = "J2:" + "J" + lasrow2;

                Excel.Range fromInvDate = currnet.Range[InvDate1];
                Excel.Range toInvDate = newWorksheet.Range[InvDate2];

                try
                {
                    fromPoNum.Copy(toPoNum);
                    fromtotal.Copy(tototal);
                    fromComp.Copy(toComp);
                    if (ShippingA != "")
                    {
                        fromship.Copy(toship);
                    }
                    if (InvoiceN != "")
                    {
                        frominv.Copy(toiv);
                    }
                    if (PoDate != "")
                    {
                        fromPoDate.Copy(toPOdate);
                    }

                    if (InvoiceD != "")
                    {
                        fromInvDate.Copy(toInvDate);
                    }
                }
                catch (Exception)
                {
                    newWorksheet.Range["A1:H1"].Value = "Do not Submit this Report it Has Errors";
                    newWorksheet.Range["A1:H1"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    newWorksheet.Range["A1:H1"].Font.Size = "19";
                    newWorksheet.Range["A1:H1"].Font.Color = Color.Yellow;

                    MessageBox.Show("Somthing went wrong with copy, please try again");
                    MessageBox.Show("Somthing went wrong, please try again");
                }

            }
            catch (Exception)
            {
                MessageBox.Show("Somthing went Wrong Please check the Sheet names");
            }



        }

        public void MemberaddZeroOnblank()
        {
            int replasrow2 = 0;
            Excel.Worksheet repinfo = (Excel.Worksheet)Globals.ThisAddIn.Application.Sheets["Import"];

            Excel.Range repLastRow = repinfo.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            replasrow2 = repLastRow.Row;
            Excel.Range rng2 = repinfo.Range["G2:G" + replasrow2];



            Excel.Range rng7 = repinfo.Range["G2:G" + replasrow2];
            Excel.Range rng10 = repinfo.Range["J2:J" + replasrow2];

            foreach (Range valb in rng10)
            {
                if (valb.Cells.Value == null)
                {
                    valb.Value = "0";

                }
            }
            foreach (Excel.Range vala in rng7)
            {


                if (vala.Cells.Value == null)
                {
                    vala.Value = "0";

                }



            }
        }

        public void VendoraddZeroOnblank()
        {
            int replasrow2 = 0;
            Excel.Worksheet repinfo = (Excel.Worksheet)Globals.ThisAddIn.Application.Sheets["Import"];

            Excel.Range repLastRow = repinfo.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            replasrow2 = repLastRow.Row;
            Excel.Range rng2 = repinfo.Range["D2:D" + replasrow2];



            Excel.Range rng7 = repinfo.Range["D2:D" + replasrow2];
            Excel.Range rng10 = repinfo.Range["H2:H" + replasrow2];

            foreach (Range valb in rng10)
            {
                if (valb.Cells.Value == null)
                {
                    valb.Value = "0";

                }
            }
            foreach (Excel.Range vala in rng7)
            {


                if (vala.Cells.Value == null)
                {
                    vala.Value = "0";

                }



            }
        }

        

    }



}







  

