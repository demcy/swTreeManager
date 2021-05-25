﻿using System;
using SW;
using XL;
using System.Configuration;
using System.Collections.Specialized;

namespace CORE
{
    public class Core
    {
        SwTools swTools = new SwTools();
        Xl xl = new Xl();
        public void SwInit()
        {
            string database = ConfigurationManager.AppSettings.Get("material");
            
                
            Console.WriteLine("DB" + database);
            SwAssy swAssy;
            if (swTools.SwConnect() && swTools.SwOpenFile(out swAssy, database))
            {
                Console.WriteLine("READING DATA...");
                Console.WriteLine("IT TAKES SOME TIME...");
                swTools.SwRead(1, 0, 10, swAssy);
                xl.OpenExcel(swTools);
            }
            else
            {
                Console.WriteLine("No file to manage");
            }
        }
    }
}