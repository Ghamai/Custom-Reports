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
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace Custom_Reports
{
    public partial class Reports
    {
        public string test { get; set; }
        public string test1 { get; set; }
        public string shitrans { get; set; }
        private void Custome_Reports_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            checkColumn();
            CheckShiping();

            //MessageBox.Show(test + "And" + test1); // this line is just for troubleshootign to see test and test1 are modified

            if (test != "No" && test1 != "No")
            {
                VCopy();
                VCombine();
                Vreconcil();
                VendorDups();
            }

        }

        private void Invoiced_TextChanged(object sender, RibbonControlEventArgs e)
        {

        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            checkColumn();
            CheckShiping();

            //MessageBox.Show(test);

            //MessageBox.Show(test1);


            if (test != "No" && test1 != "No")

            {
                Copy();

                ActionL ac = new ActionL();
                ac.MemberaddZeroOnblank();
                Logic lc = new Logic();
                lc.MemberReconcile();
                if (invoicen.Text.ToString() != "")
                {
                    lc.Dup2();
                }
                else if (invoicen.Text.ToString() == "")
                {
                    lc.Dup22();
                }
            }


        }

        private void checkColumn()
        {
            ActionL ac = new ActionL();
            ac.cleanColumn(PoTotal.Text.ToString(), PoTotal.Name.ToString());
            test = ac.Txbox;
            
        }

      
        private void CheckShiping()
        {
            ActionL ac = new ActionL();
            ac.cleanColumn(ShipingA.Text.ToString(),ShipingA.Name.ToString());
            test1 = ac.txbox1;
            
        }
        private void VCopy()
        {
            ActionL Pk = new ActionL();
            Pk.Company = Company.Text.ToString();
            Pk.PoTotal = PoTotal.Text.ToString();
            Pk.InvoiceD = Invoiced.Text.ToString();
            Pk.MonthYear = month.Text.ToString();
            Pk.InvoiceN = invoicen.Text.ToString();
            Pk.PoDate = PoDate.Text.ToString();
            Pk.GroupNumber = GroupN.Text.ToString();
            Pk.PoColumn = PoNumber.Text.ToString();
            Pk.ShippingA = ShipingA.Text.ToString();
            Pk.Vcopy();
        }
        private void Copy()

        {
            


            Connection cn = new Connection();
            ActionL Pk = new ActionL();
            Pk.Company = Company.Text.ToString();
            Pk.PoTotal = PoTotal.Text.ToString();
            Pk.InvoiceD = Invoiced.Text.ToString();
            Pk.MonthYear = month.Text.ToString();
            Pk.InvoiceN = invoicen.Text.ToString();
            Pk.PoDate = PoDate.Text.ToString();
            Pk.GroupNumber = GroupN.Text.ToString();
            Pk.PoColumn = PoNumber.Text.ToString();
            Pk.ShippingA = ShipingA.Text.ToString();





            Pk.Copy();
            
           
        }

        private void VCombine()
        {
            Logic lg = new Logic();
            lg.JoinC = JoinCheck.Checked;
            lg.MonthYear = month.Text.ToString();
            lg.InvoiceN = invoicen.Text.ToString();
            lg.VendorCombineCol();
        }

        private void Vreconcil()
        {
            Logic lg = new Logic();
            lg.VendorReconcile();
        }

        private void VendorDups()
        {
            Logic lg = new Logic();
            lg.VendorDuplicate();
        }

        private void PoNumber_TextChanged(object sender, RibbonControlEventArgs e)
        {

        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            
            checkColumn();
            CheckShiping();

            MessageBox.Show(test);
            
            MessageBox.Show(test1);
        

            if (test != "No" && test1 != "No")

            {
                Copy();

                ActionL ac = new ActionL();
                ac.MemberaddZeroOnblank();
                Logic lc = new Logic();
                lc.MemberReconcile();
                if (invoicen.Text.ToString() != "")
                {
                    lc.Dup2();
                }
                else if (invoicen.Text.ToString() == "")
                {
                    lc.Dup22();
                }
            }

        }

        private void Test2_Click(object sender, RibbonControlEventArgs e)
        {
            ActionL ac = new ActionL();
            
           
        }

        private void PoTotal_TextChanged(object sender, RibbonControlEventArgs e)
        {

        }
    }
}
