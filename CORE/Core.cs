using System;
using SW;
using XL;

namespace CORE
{
    public class Core
    {
        SwTools swTools = new SwTools();
        Xl xl = new Xl();
        public void SwInit()
        {
            if (swTools.SwConnect() && swTools.SwOpenFile())
            {
                Console.WriteLine("READING DATA...");
                Console.WriteLine("IT TAKES SOME TIME...");
                swTools.SwRead();
                xl.OpenExcel(swTools);
            }
            else
            {
                Console.WriteLine("No file to manage");
            }
        }
    }
}