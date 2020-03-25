using System;
using System.Collections.Generic;
using System.Globalization;
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
        private Dictionary<Tuple<double, string>, List<SwComps>> orderedDic;
        private double p = 1;
        private int level = 1;
        private string a = "";
        
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
            //string lastP;
            orderedDic = swTools.MainAss.OrderBy(item => item.Key.Item1)
                //.ThenBy(item => item.Key.Item2)
                .ToDictionary(t => t.Key, v => v.Value);
            foreach (var dicItem in orderedDic)
            {
                string lastP = ItemNo(dicItem.Key.Item1);
                _xlSheet.Cells[i, 1].Value = lastP;
                _xlSheet.Cells[i, 2].Value = dicItem.Key.Item2;
                i++;
                //lastP = p.ToString() + ".";
                orderedComps = dicItem.Value.GroupBy(item => item.Name)
                    .Select(item => item.FirstOrDefault())
                    .OrderBy(item => item.Name).ToList();
                foreach (var comp in orderedComps)
                {
                    _xlSheet.Cells[i, 1].Value = lastP + "." + p.ToString();
                    _xlSheet.Cells[i, 2].Value = comp.Name;
                    _xlSheet.Cells[i, 3].Value = comp.Description;
                    _xlSheet.Cells[i, 4].Value = comp.CompanyNo;
                    _xlSheet.Cells[i, 5].Value = comp.Material;
                    i++;
                    p++;
                }
            }
            _xlSheet.Columns.AutoFit();
            _xlSheet.Columns.Locked = false;
            _xlSheet.Range[_xlSheet.Cells[1, 1],_xlSheet.Cells[1, 5]].Locked = true;
            _xlSheet.Columns[1].Locked = true;
            _xlSheet.Columns[2].Locked = true;
            _xlSheet.Protect(); 
            
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
            foreach (Range v in target.Rows)
            {
                string xlName = _xlSheet.Cells[v.Row, 2].Text;
                Console.WriteLine(xlName);
                var comp = orderedDic.FirstOrDefault(d=>d.Value
                        .Any(item=>item.Name==xlName)).Value
                    .FirstOrDefault(item=>item.Name==xlName);
                //var comp = orderedComps[v.Row-2];
                _swTools.SwWrite(comp, v.Text, target.Column);
                
            }
        }

        /*private void ChangExcel(Excel.Range target)
        {
            Console.WriteLine("BBBB");
            foreach (Range v in target.Rows)
            {
                //bool result = true;
                string zz = _xlSheet.Cells[1, 1];
                string b = _xlSheet.Cells[v.Row, 2];
                Console.WriteLine("From excel: " + b);
                Console.WriteLine(zz);
                var gt = orderedDic.FirstOrDefault(q=>q.Value.Any(u=>u.Name==b)).Value
                    .FirstOrDefault(z=>z.Name==b);
                Console.WriteLine(gt.Name);    
                /*var comp = orderedDic
                    .Where(item => item.Value
                        .Any(k => k.Name == _xlSheet.Cells[v.Row, 2]))
                    .Select(t => t.Value.Any(pp => pp.Name == _xlSheet.Cells[v.Row, 2]))
                    .Select(n=>n.)*/
                        
                  //  .Where(k=>k.)
                //orderedComps.Where(item=>item.Name==_xlSheet.Cells[v.Row,2])
                //var comp = orderedComps[v.Row-2];
                //_swTools.SwWrite(comp, v.Text, target.Column);//, out result);
                /*if (!result)
                {
                    _xlApp.Undo();
                }*/
      
    }
}

/*
        xlBook.Save();
        xlBook.Close();
        XL.Quit();*/