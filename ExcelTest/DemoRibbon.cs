using ExcelTest.ExcelObj;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using WindowsFormsApp1;


namespace ExcelTest
{
    public partial class DemoRibbon
    {
        private void DemoRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void DemoButton1_Click(object sender, RibbonControlEventArgs e)
        {
            var form = new Form1();
            form.Show();
            MessageBox.Show("Button1");
        }
        
        private void DemoButton2_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("Button2");
            var excelCom = new ExcelCommands(Globals.ThisAddIn.Application);
            excelCom.CompareRangeUsage();

        }

        private void DemoButton3_Click(object sender, RibbonControlEventArgs e)
        {
            
        }
    }
}
