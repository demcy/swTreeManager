using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using SldWorks;

namespace SW
{
    public class SwTools
    {
        private SldWorks.SldWorks _swApp;
        private SldWorks.AssemblyDoc _swAss;
        private SldWorks.ModelDoc2 _swModel;
        //public SwAssy swAssy;
        
        public List<Component2> AssList = new List<Component2>();
        public Dictionary<Tuple<double,SwAssy>, List<SwComps>> MainAss = new Dictionary<Tuple<double,SwAssy>, List<SwComps>>();
        
        string Database = "S:/Solidworks Settings/Materials/FD2P Other Materials.sldmat";
        string BuilderTemplate = "S:/Solidworks Settings/Templates/FD2P/FD2P Custom Properties/FD2P Custom Properties Part.prtprp";

        public bool SwConnect()
        {
            try
            {
                _swApp = (SldWorks.SldWorks) Marshal.GetActiveObject("SldWorks.Application");
                Console.WriteLine("SolidWorks is connected");
                return true;
            }
            catch
            {
                Console.WriteLine("SolidWorks not connected");
                return false;
            }
        }
        public bool SwOpenFile(out SwAssy swAssy)
        {
            _swAss = (SldWorks.AssemblyDoc) _swApp.ActiveDoc;
            _swModel = (SldWorks.ModelDoc2) _swApp.ActiveDoc;
            _swAss.ResolveAllLightweight();
            swAssy = new SwAssy(null);
            swAssy.Name = _swModel.GetTitle();
            SldWorks.Configuration swConf = (SldWorks.Configuration) _swModel.GetActiveConfiguration();
            string ConfName = swConf.Name;
            swAssy.Description = _swModel.CustomInfo2[swConf.Name, "Description"];
            swAssy.CompanyNo = _swModel.CustomInfo2[swConf.Name, "Company No"];
            return _swAss != null;
        }
        public void SwRead(double l, double p, double d, SwAssy swAssy)
        {
            double level = l;
            List<SwComps> Comps = new List<SwComps>();
            
            object[] objComponents = (object[]) _swAss.GetComponents(true);
            foreach (var objComponent in objComponents)
            {
                Component2 c = (Component2)objComponent;
                _swModel = (SldWorks.ModelDoc2) c.GetModelDoc2();
                int isToolbox;
                if (!c.IsSuppressed())
                {
                    isToolbox = _swModel.Extension.ToolboxPartType;
                }
                else
                {
                    isToolbox = 3;
                }
                if (isToolbox == 0)
                {
                    if (_swModel.GetType() == 2)
                    {
                        p++;
                        _swAss = (SldWorks.AssemblyDoc)_swApp.ActivateDoc(c.GetPathName());
                        swAssy = new SwAssy((SldWorks.Component2) objComponent);
                        //RECURSION!!!
                        SwRead(level + p/d, 0,d*10, swAssy);
                        _swApp.CloseDoc(c.GetPathName());
                    }
                    SwComps comp = new SwComps((SldWorks.Component2) objComponent);
                    if (comp.isToolbox == 0)
                    {
                        Comps.Add(comp);
                    }
                }
            }
            //_swModel = (SldWorks.ModelDoc2) _swApp.ActiveDoc;
            Tuple<double, SwAssy> header = new Tuple<double, SwAssy>(level, swAssy);//_swModel.GetTitle());
            MainAss.Add(header, Comps);
        }

        public void SwWrite(SwComps comp, string changedText, int colN)//, out bool result)
        {
            //result = true;
            SldWorks.ModelDoc2 swModel = (SldWorks.ModelDoc2) comp.Comp.GetModelDoc2();
            SldWorks.PartDoc swPart = (SldWorks.PartDoc) comp.Comp.GetModelDoc2();
            
            if (swModel.Extension.CustomPropertyBuilderTemplate[false] != BuilderTemplate)
            {
                swModel.Extension.CustomPropertyBuilderTemplate[false] = BuilderTemplate;
                swModel.AddCustomInfo3(comp.ConfName, "Description", 0, "");
                swModel.AddCustomInfo3(comp.ConfName, "Company No", 0, "");
            }
            if (colN == 3)
            {
                swModel.CustomInfo2[comp.ConfName, "Description"]=changedText;
            }
            if (colN == 4)
            {
                swModel.CustomInfo2[comp.ConfName, "Company No"]=changedText;
            }
            if (colN == 5)
            {
                swPart.SetMaterialPropertyName2(comp.ConfName, Database, changedText);
                /*string m = swPart.GetMaterialPropertyName2(comp.ConfName, out Database);
                if (m != changedText)
                {
                    result = false;
                }*/
            }
            _swAss.ForceRebuild2(true);
        }
        

        
    }
}
//private static SldWorks.ModelDoc2 _swModel;
// private Feature _swFeature;
//_swModel = (SldWorks.ModelDoc2) _swApp.ActiveDoc;
//private Component2[] _swCmpnNts; 
//return _swModel.GetPathName();
//FeatureManager fm = swModel.FeatureManager;
// object features = fm.GetFeatures(true);
/*  Console.WriteLine(swModel.GetFeatureCount());
  swFeature = (Feature) swModel.FirstFeature();
  //swFeature.GetNextFeature();*/