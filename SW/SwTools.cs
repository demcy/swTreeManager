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

        public string assAdrs;

        //public List<SldWorks.Component2> Comps = new List<SldWorks.Component2>();
        public List<SwComps> Comps = new List<SwComps>();

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

        public bool SwOpenFile()
        {
            _swAss = (SldWorks.AssemblyDoc) _swApp.ActiveDoc;
            _swAss.ResolveAllLightweight();
            return _swAss != null;
        }

        public void SwRead()
        {
            adrs();
            object[] objComponents = (object[]) _swAss.GetComponents(true);
            foreach (var objComponent in objComponents)
            {
                SwComps comp = new SwComps((SldWorks.Component2) objComponent);
                //SldWorks.Component2 tempComponent = (SldWorks.Component2) objComponent;
                //SldWorks.ModelDoc2 swm = (SldWorks.ModelDoc2) tempComponent.GetModelDoc2();
                Comps.Add(comp);
            }
        }

        public void SwRebuildSave()
        {
            _swAss.ForceRebuild2(true);
        }


        public void adrs()
        {
            _swModel = (SldWorks.ModelDoc2) _swApp.ActiveDoc;
            string a = _swModel.GetPathName();
            assAdrs = a.Substring(0, a.LastIndexOf("\\"));
            Console.WriteLine(assAdrs);
        }

        public string restAdr(string p)
        {
            return p.Substring(assAdrs.Length + 1, p.Length - assAdrs.Length - 1);
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