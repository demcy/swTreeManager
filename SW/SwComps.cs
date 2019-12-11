namespace SW
{
    public class SwComps
    {
        public string Name;
        public string Description;
        public string CompanyNo;
        public string Material;
        public SldWorks.Component2 Comp;
        private SldWorks.ModelDoc2 swModel;
        private SldWorks.PartDoc swPart;
        private SldWorks.Configuration swConf;

        public SwComps(SldWorks.Component2 comp)
        {
            Comp = comp;
            swModel = (SldWorks.ModelDoc2) comp.GetModelDoc2();
            swPart = (SldWorks.PartDoc) comp.GetModelDoc2();
            swConf = (SldWorks.Configuration) swModel.GetActiveConfiguration();
            Name = comp.Name;
            Description = swModel.CustomInfo2[swConf.Name, "Description"];
            CompanyNo = swModel.CustomInfo2[swConf.Name, "Company No"];
            string outDatabase;
            Material = swPart.GetMaterialPropertyName2(swConf.Name, out outDatabase);
        }
    }
}