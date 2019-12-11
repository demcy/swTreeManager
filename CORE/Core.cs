using System;
using System.Xml.Serialization;
using SW;
using SldWorks;

namespace CORE
{
    public class Core
    {
        SwTools swTools = new SwTools();

        public void swInit()
        {
            if (swTools.SwConnect() && swTools.SwOpenFile())
            {
                Console.WriteLine("READING DATA...");
                Console.WriteLine("IT TAKES SOME TIME...");
                swLoad();
            }
            else
            {
                Console.WriteLine("No file to manage");
            }
        }

        public void swWrite()
        {
            Console.Write("What Part number to get material from? ");
            int n = Convert.ToInt16(Console.ReadLine());
            var getComp = swTools.Comps[n - 1];
            Console.WriteLine("Choosen material: " + getComp.Material);
            foreach (var comp in swTools.Comps)
            {
                SldWorks.ModelDoc2 swModel = (SldWorks.ModelDoc2) comp.Comp.GetModelDoc2();
                SldWorks.Configuration swConf = (SldWorks.Configuration) swModel.GetActiveConfiguration();
                SldWorks.PartDoc swPart = (SldWorks.PartDoc) comp.Comp.GetModelDoc2();
                string Database = "S:/Solidworks Settings/Materials/FD2P Other Materials.sldmat";
                swPart.SetMaterialPropertyName2(swConf.Name, Database, getComp.Material);
            }

            Console.WriteLine("Material is changed");
            swTools.SwRebuildSave();
            swRead();
        }

        public void swLoad()
        {
            swTools.SwRead();
            swRead();
            swWrite();
        }

        public void swRead()
        {
            int i = 1;
            foreach (var comp in swTools.Comps)
            {
                Console.WriteLine(i++ + ".)Component name: " + comp.Name);
                Console.WriteLine("\t" + "Description: " + comp.Description);
                Console.WriteLine("\t" + "Company No: " + comp.CompanyNo);
                Console.WriteLine("\t" + "Material: " + comp.Material);
            }
        }
    }
}