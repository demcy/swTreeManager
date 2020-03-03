using System;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using SW;

namespace XL
{
    public class Xl
    {
        
        
        private readonly Excel.Application _xlApp = new Excel.Application();
        private Excel.Workbook _xlBook;
        private Excel.Worksheet _xlSheet;

        private SwTools _swTools;
        public void OpenExcel(SwTools swTools)
        {
            _swTools = swTools;
            _xlApp.Visible = true;
            _xlBook = _xlApp.Workbooks.Add();
            _xlSheet = (Excel.Worksheet)_xlBook.Worksheets[1];
           
            _xlSheet.Cells[1, 1].Value = "Component name";
            _xlSheet.Cells[1, 2].Value = "Description";
            _xlSheet.Cells[1, 3].Value = "Company No";
            _xlSheet.Cells[1, 4].Value = "Material";
            
            _xlSheet.Range[_xlSheet.Cells[1, 1],_xlSheet.Cells[1, 4]].Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            _xlSheet.Range[_xlSheet.Cells[1, 1], _xlSheet.Cells[1, 4]].Cells.Font.Bold = true;
            
            int i = 2;
            foreach (var comp in swTools.Comps)
            {
                _xlSheet.Cells[i, 1].Value = comp.Name;
                _xlSheet.Cells[i, 2].Value = comp.Description;
                _xlSheet.Cells[i, 3].Value = comp.CompanyNo;
                _xlSheet.Cells[i, 4].Value = comp.Material;
                i++;
            }
            _xlSheet.Columns.AutoFit();
            _xlSheet.Columns.Locked = false;
            _xlSheet.Range[_xlSheet.Cells[1, 1],_xlSheet.Cells[1, 4]].Locked = true;
            _xlSheet.Columns[1].Locked = true;
            _xlSheet.Protect(); 
            
            _xlSheet.Change += new Excel.DocEvents_ChangeEventHandler(ChangExcel);
            
        }
        private void ChangExcel(Excel.Range target)
        {
            foreach (Range v in target.Rows)
            {
                bool result = true;
                var comp = _swTools.Comps[v.Row-2];
                _swTools.SwWrite(comp, v.Text, target.Column, out result);
                if (!result)
                {
                    _xlApp.Undo();
                }
            }
        }
    }
}

/*
        xlBook.Save();
        xlBook.Close();
        XL.Quit();*/