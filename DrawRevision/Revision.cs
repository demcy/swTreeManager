﻿using System.Collections.Generic;
using System.IO;
using System.Linq;
using SW;

namespace DrawRevision
{
    public class Revision
    {
        SwTools swTools = new SwTools();
        public void GetFiles(string[] names, string path, List<string> props)
        {
            if (!swTools.SwConnect()) return;
            foreach (var name in names)
            {
                if (CheckExist(name, path, props[0])) continue;
                swTools.EasyOpen(name);
                swTools.AddRevision(props);
                swTools.SaveToPdf(GetName(name, path, props[0]));
                swTools.CloseDoc(name);
            }
        }

        public bool CheckExist(string name, string path, string rev)
        {
            string[] p = Directory.GetFiles(path);
            var backIndex = name.LastIndexOf("\\") + 1;
            var pointIndex = name.LastIndexOf(".");
            name = name.Substring(backIndex, pointIndex - backIndex) + "_rev."+rev;
            return p.Any(v => v.Contains(name));
        }

        public string GetName(string name, string path, string rev)
        {
            var backIndex = name.LastIndexOf("\\") + 1;
            var pointIndex = name.LastIndexOf(".");
            name = name.Substring(backIndex, pointIndex - backIndex);
            name = path + name + "_rev."+rev+".pdf";
            return name;
        }
    }
}