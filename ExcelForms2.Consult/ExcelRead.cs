using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace ExcelForms2.Consult
{
    internal class ExcelRead
    {
        int[] Stolbci { get; set; }
        string Name { get; set; }
        public ExcelRead(string name)
        {
            Name = name;
        }

        public string ReadDokument(params ExcelCell[] stolbci)
        {
            var xLApp = new Application();
            xLApp.Visible = true;

            Workbook workbook = xLApp.Workbooks.Open(Name);
            Worksheet workSheet = workbook.Sheets[1];


            int stroka = 2;
            string result = "";
            while (CheckEnd(workSheet, stroka))
            {
                for(int i = 0; i < stolbci.Length; i++)
                {
                    string cellText = Convert.ToString(workSheet.Cells[stroka, stolbci[i].Cell].Value2); 
                    if(stolbci[i].Extract == true)
                    {                        
                        cellText= Detect(cellText, stolbci[i].StartElement, stolbci[i].EndElement);
                    
                    }

                    result += cellText +"  <->   ";
                }
                result += "\n";
                stroka++;
            }


            workbook.Close(0);
            xLApp.Quit();

                return result;
        }       

        bool CheckEnd(Worksheet workSheet, int stroka)
        {
            if (Convert.ToString(workSheet.Cells[stroka, 1].Value2) != null) return true;
            return false;            
        }
        private string Detect(string doc, string startElement, string endElement)
        {
            int indexStart = doc.IndexOf(startElement);            
            string result = doc.Substring(indexStart + startElement.Length);

            int indexEnd = result.IndexOf(endElement)-2;
            result=result.Substring(0, indexEnd);
            return result;
        }
    }
}