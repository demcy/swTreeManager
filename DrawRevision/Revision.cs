using SW;

namespace DrawRevision
{
    public class Revision
    {
        SwTools swTools = new SwTools();
        public void GetFiles(string[] names, string path)
        {
            if (!swTools.SwConnect()) return;
            foreach (var name in names)
            {
                swTools.EasyOpen(name);
                swTools.AddRevision();
                swTools.SaveToPdf(GetName(name, path));
                swTools.CloseDoc(name);
            }
        }

        public string GetName(string name, string path)
        {
            var backIndex = name.LastIndexOf("\\") + 1;
            var pointIndex = name.LastIndexOf(".");
            name = name.Substring(backIndex, pointIndex - backIndex);
            name = path + name + "_rev.Z.pdf";
            return name;
        }
    }
}