using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelTest.ExcelObj
{
    internal class ExcelCommands
    {
        //Can be get from static global variable created in ThisAddin 
        //Application Application = Globals.ThisAddIn.Application;


        Application _application;

        internal ExcelCommands(Application application) { _application = application; }


        //Возвращает объект приложение Excel
        public Excel.Application GetApplication()
        {
            return _application;
        }
        //Возвращает объект активная рабочаяя книга
        public Excel.Workbook GetActiveWorkBook()
        {
            return (Excel.Workbook)_application.ActiveWorkbook;
        }
        //Возвращает объект активный рабочий лист
        public Excel.Worksheet GetActiveWorksheet()
        {
            return (Excel.Worksheet)_application.ActiveSheet;
        }
        //Возвращает объект активная ячейка
        public Excel.Range GetActiveCell()
        {
            return (Excel.Range)_application.Selection;
        }
    }
}
