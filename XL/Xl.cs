using System;
using System.Collections.Generic;
using System.Linq;
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
        private List<SwComps> orderedComps;
        private Dictionary<Tuple<int, string>, List<SwComps>> orderedDic;
        public void OpenExcel(SwTools swTools)
        {
            _swTools = swTools;
            _xlApp.Visible = true;
            _xlBook = _xlApp.Workbooks.Add();
            _xlSheet = (Excel.Worksheet)_xlBook.Worksheets[1];
           
            _xlSheet.Cells[1, 1].Value = "Item No";
            _xlSheet.Cells[1, 2].Value = "Component name";
            _xlSheet.Cells[1, 3].Value = "Description";
            _xlSheet.Cells[1, 4].Value = "Company No";
            _xlSheet.Cells[1, 5].Value = "Material";
            
            _xlSheet.Range[_xlSheet.Cells[1, 1],_xlSheet.Cells[1, 5]].Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            _xlSheet.Range[_xlSheet.Cells[1, 1], _xlSheet.Cells[1, 5]].Cells.Font.Bold = true;
            
            int i = 2;
            orderedDic = swTools.MainAss.OrderBy(item => item.Key.Item1)
                .ToDictionary(t => t.Key, v => v.Value);
            foreach (var dicItem in orderedDic)
            {
                _xlSheet.Cells[i, 1].Value = dicItem.Key.Item1;
                _xlSheet.Cells[i, 2].Value = dicItem.Key.Item2;
                i++;
            }
            /*orderedComps = swTools.Comps.GroupBy(item => item.Name)
                .Select(item => item.FirstOrDefault())
                .OrderBy(item => item.Name).ToList();
            foreach (var comp in orderedComps)
            {
                _xlSheet.Cells[i, 1].Value = comp.Name;
                _xlSheet.Cells[i, 2].Value = comp.Description;
                _xlSheet.Cells[i, 3].Value = comp.CompanyNo;
                _xlSheet.Cells[i, 4].Value = comp.Material;
                i++;
            }
            */
            
            _xlSheet.Columns.AutoFit();
            _xlSheet.Columns.Locked = false;
            _xlSheet.Range[_xlSheet.Cells[1, 1],_xlSheet.Cells[1, 5]].Locked = true;
            _xlSheet.Columns[1].Locked = true;
            _xlSheet.Columns[2].Locked = true;
            _xlSheet.Protect(); 
            
            _xlSheet.Change += new Excel.DocEvents_ChangeEventHandler(ChangExcel);
            
        }
        private void ChangExcel(Excel.Range target)
        {
            foreach (Range v in target.Rows)
            {
                bool result = true;
                var comp = orderedComps[v.Row-2];
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