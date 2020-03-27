namespace SW
{
    public class SwAssy
    {
        public string Name;
        public string Description;
        public string CompanyNo;
        private SldWorks.ModelDoc2 swModel;
        private SldWorks.Configuration swConf;
        public string ConfName;
        public SwAssy(SldWorks.Component2 comp)
        {
            if (comp != null)
            {
                swModel = (SldWorks.ModelDoc2) comp.GetModelDoc2();
                Name = swModel.GetTitle();
                swConf = (SldWorks.Configuration) swModel.GetActiveConfiguration();
                ConfName = swConf.Name;
                Description = swModel.CustomInfo2[swConf.Name, "Description"];
                CompanyNo = swModel.CustomInfo2[swConf.Name, "Company No"];
            }
        }
    }
}