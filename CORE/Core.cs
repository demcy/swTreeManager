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
            SwAssy swAssy;
            if (swTools.SwConnect() && swTools.SwOpenFile(out swAssy))
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