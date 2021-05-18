using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
// using System.Reflection;
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
        private Dictionary<Tuple<double, SwAssy>, List<SwComps>> orderedDic;
        private double p = 1;
        private int level = 1;
        private string a = "";
        
        //public event Excel.DocEvents_ChangeEventHandler Change;
        
        public void OpenExcel(SwTools swTools)
        {
            
            _swTools = swTools;
            _xlApp.Visible = true;
            _xlBook = _xlApp.Workbooks.Open(
                @"S:\Programs\Macros\SW TreeManager\SW TreeManager\sw-tree-manager_Template.xlsx");
            _xlSheet = (Excel.Worksheet)_xlBook.Worksheets[1];
           
            //.Cells[1, 1] = "Item No";
            // _xlSheet.Cells[1, 2] = "Component name";
            // _xlSheet.Cells[1, 3] = "Description";
            // _xlSheet.Cells[1, 4] = "Company No";
            // _xlSheet.Cells[1, 5] = "Material";
            
            // _xlSheet.Range[_xlSheet.Cells[1, 1],_xlSheet.Cells[1, 5]].Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            // _xlSheet.Range[_xlSheet.Cells[1, 1], _xlSheet.Cells[1, 5]].Cells.Font.Bold = true;
            
            int i = 2;
            //string lastP;
            orderedDic = swTools.MainAss.OrderBy(item => item.Key.Item1)
                //.ThenBy(item => item.Key.Item2)
                .ToDictionary(t => t.Key, v => v.Value);
            foreach (var dicItem in orderedDic)
            {
                string lastP = ItemNo(dicItem.Key.Item1);
                _xlSheet.Cells[i, 1] = lastP;
                _xlSheet.Cells[i, 2] = dicItem.Key.Item2.Name;
                _xlSheet.Cells[i, 3] = dicItem.Key.Item2.Description;
                _xlSheet.Cells[i, 4] = dicItem.Key.Item2.CompanyNo;
                i++;
                //lastP = p.ToString() + ".";
                orderedComps = dicItem.Value.GroupBy(item => item.Name)
                    .Select(item => item.FirstOrDefault())
                    .OrderBy(item => item.Name).ToList();
                foreach (var comp in orderedComps)
                {
                    _xlSheet.Cells[i, 1] = lastP + "." + p.ToString();
                    _xlSheet.Cells[i, 2] = comp.Name;
                    _xlSheet.Cells[i, 3] = comp.Description;
                    _xlSheet.Cells[i, 4] = comp.CompanyNo;
                    _xlSheet.Cells[i, 5] = comp.Material;
                    //COMP/GETVALUE OF cells(1, x) -< while x in not ""!!!!
                    i++;
                    p++;
                }
            }
            //_xlSheet.Columns.AutoFit();
            //_xlSheet.Columns.Locked = false;
            //_xlSheet.Range[_xlSheet.Cells[1, 1],_xlSheet.Cells[1, 5]].Locked = true;
            //_xlSheet.Columns[1].Locked = true;
            //_xlSheet.Columns[2].Locked = true;
            //_xlSheet.Protect(); 
            
            _xlSheet.Change += new Excel.DocEvents_ChangeEventHandler(ChangExcel);
            
            
        }

        private string ItemNo(double d)
        {
            int curL = 3 + 2 * (d.ToString().Length - 3);
            if (d == 1)
            {
                a = "1";
                return a;
            }

            if (level < curL)
            {
                level = curL;
                a = a + "." + p;
                p = 1;
                return a;
            }
            if (level == curL)
            {
                int lastC = int.Parse(a.Substring(a.Length-1,1));
                p = 1;
                a = a.Substring(0, a.Length - 1) + (lastC + 1).ToString();
                return a;
            }
            if (level > curL)
            {
                int dif = level - curL + 1;
                level = curL;
                int lastC = int.Parse(a.Substring(a.Length-dif,1));
                p = 1;
                a = a.Substring(0, a.Length - dif) + (lastC + 1).ToString();
                return a;
            }

            return "";

        }
        
        private void ChangExcel(Excel.Range target)
        {
            foreach (Excel.Range v in target.Rows)
            {
                string xlName = (_xlSheet.Cells[v.Row, 2] as Excel.Range).Value2.ToString();
                string colName = (_xlSheet.Cells[1, v.Column] as Excel.Range).Value2.ToString();
                bool ifPart = orderedDic.Any(d => d.Value
                     .Any(item => item.Name == xlName));
                if (ifPart)
                {
                    var comp = orderedDic.FirstOrDefault(d => d.Value
                            .Any(item => item.Name == xlName)).Value
                        .FirstOrDefault(item => item.Name == xlName);
                    _swTools.SwWritePart(comp, v.Text.ToString(), colName);
                }
                else
                {
                    var comp = orderedDic.FirstOrDefault(d => d.Key.Item2.Name == xlName)
                             .Key.Item2;
                    _swTools.SwWriteAssy(comp, v.Text.ToString(), colName);
                }
            }
        }
    }
}

/*
 //var comp = orderedComps[v.Row-2];
        xlBook.Save();
        xlBook.Close();
        XL.Quit();*/
        
/*
        //bool result = true;
        //var comp = orderedComps[v.Row-2];
        //_swTools.SwWrite(comp, v.Text, target.Column);//, out result);
        if (!result)
        {
            _xlApp.Undo();
        }*/