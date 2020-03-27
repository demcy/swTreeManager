using System;
using CORE;


namespace ConsoleApp
{
    internal class Program
    {
        public static void Main(string[] args)
        {
            /*int level = "1.3.4".ToString().Length;
            
            int dif = level - "1.2".Length + 1;
            Console.WriteLine(dif);
            string a = "1.3.4";
            int lastC = int.Parse(a.Substring(a.Length-dif,1));
            Console.WriteLine(lastC);
            a = a.Substring(0, a.Length - dif) + (lastC + 1).ToString();
            Console.WriteLine(a);*/
            //p = 1;
            //a = a.Substring(0, a.Length - 3) + (lastC + 1).ToString();
            //return a;
            
            
            
            DateTime t1 = DateTime.Now;
            //Console.WriteLine(DateTime.Now.ToString("HH:mm:ss tt"));
            Core core = new Core();
            core.SwInit();
            DateTime t2 = DateTime.Now;
            Console.WriteLine((t2 - t1).TotalSeconds + "s.");
            Console.WriteLine("You can edit excel now");
            Console.ReadLine();
        }
    }
}