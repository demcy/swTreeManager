﻿using System;

namespace SW
{
    public class SwComps
    {
        public string Name;
        public string Description;
        public string CompanyNo;
        public string Material;
        public int isToolbox;
        public string ConfName;
        public SldWorks.Component2 Comp;
        private SldWorks.ModelDoc2 swModel;
        private SldWorks.PartDoc swPart;
        private SldWorks.Configuration swConf;
        
        public SwComps(SldWorks.Component2 comp)
        {
            string outDatabase;
            try
            {
                Comp = comp;
                swModel = (SldWorks.ModelDoc2) comp.GetModelDoc2();
                swPart = (SldWorks.PartDoc) comp.GetModelDoc2();
                Name = comp.Name;
                swConf = (SldWorks.Configuration) swModel.GetActiveConfiguration();
                ConfName = swConf.Name;
                Description = swModel.CustomInfo2[swConf.Name, "Description"];
                CompanyNo = swModel.CustomInfo2[swConf.Name, "Company No"];
                Material = swPart.GetMaterialPropertyName2(swConf.Name, out outDatabase);
                isToolbox = swModel.Extension.ToolboxPartType;
                if (comp.IsHidden(true))
                {
                    isToolbox = 3;
                }
            }
            catch(Exception)
            {
                isToolbox = 3;
            }
        }
    }
}