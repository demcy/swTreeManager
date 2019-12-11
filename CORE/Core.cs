

using System;
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

        public void swLoad()
        {
            swTools.SwRead();
            foreach (SldWorks.Component2 comp in swTools.Comps)
            {
                SldWorks.ModelDoc2 swm = (SldWorks.ModelDoc2) comp.GetModelDoc2();
                SldWorks.Configuration swConf = (SldWorks.Configuration)swm.GetActiveConfiguration();
                SldWorks.PartDoc swp =( SldWorks.PartDoc) comp.GetModelDoc2();
                string outDatabase;
                Console.WriteLine("Component name: " + comp.Name);
                Console.WriteLine("\t" +  "Description: " + swm.CustomInfo2[swConf.Name, "Description"]);
                Console.WriteLine( "\t" +  "Company No: " + swm.CustomInfo2[swConf.Name, "Company No"]);
                Console.WriteLine( "\t" +  "Material: " + swm.CustomInfo2[swConf.Name, "Material"]);
                Console.WriteLine( "\t" +  "MaterialPart: " + swp.GetMaterialPropertyName2(swConf.Name,out outDatabase));
                
            }

            if (Console.ReadLine() == "mat")
            {
                int n = Convert.ToInt16(Console.ReadLine());
                SldWorks.Component2 comp = swTools.Comps[n-1];
                SldWorks.ModelDoc2 swm = (SldWorks.ModelDoc2) comp.GetModelDoc2();
                SldWorks.Configuration swConf = (SldWorks.Configuration)swm.GetActiveConfiguration();
                SldWorks.PartDoc swp =( SldWorks.PartDoc) comp.GetModelDoc2();
                string outDatabase;
                Console.WriteLine("Component name: " + comp.Name);
                Console.WriteLine( "\t" +  "Description: " + swm.CustomInfo2[swConf.Name, "Description"]);
                Console.WriteLine( "\t" +  "Company No: " + swm.CustomInfo2[swConf.Name, "Company No"]);
                Console.WriteLine( "\t" +  "Material: " + swm.CustomInfo2[swConf.Name, "Material"]);
                //string m = swm.CustomInfo2[swConf.Name, "Material"];
                string m = swp.GetMaterialPropertyName(out outDatabase);
                foreach (SldWorks.Component2 comp2 in swTools.Comps)
                {
                    SldWorks.ModelDoc2 swm2 = (SldWorks.ModelDoc2) comp2.GetModelDoc2();
                    SldWorks.Configuration swConf2 = (SldWorks.Configuration)swm2.GetActiveConfiguration();
                    SldWorks.PartDoc swp2 =( SldWorks.PartDoc) comp2.GetModelDoc2();
                    string outDatabase2;
                    string Database = "S:/Solidworks Settings/Materials/FD2P Other Materials.sldmat";
                    //swp2.GetMaterialPropertyName2()
                    //swm2.CustomInfo2[swConf2.Name, "Material"] = m;
                    swp2.SetMaterialPropertyName2(swConf2.Name, Database,m);
                    Console.WriteLine("Component name: " + comp2.Name);
                    Console.WriteLine( "\t" +  "Description: " + swm2.CustomInfo2[swConf.Name, "Description"]);
                    Console.WriteLine( "\t" +  "Company No: " + swm2.CustomInfo2[swConf.Name, "Company No"]);
                    Console.WriteLine( "\t" +  "Material: " + swm2.CustomInfo2[swConf.Name, "Material"]);
                }
            }
            
            else
            {
                
            }
            Console.WriteLine();
            
        }



    }
}