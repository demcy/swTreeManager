using System;
using CORE;


namespace ConsoleApp
{
    internal class Program
    {
        public static void Main(string[] args)
        {
            DateTime t1 = DateTime.Now;
            Console.WriteLine(DateTime.Now.ToString("HH:mm:ss tt"));
            Core core = new Core();
            core.SwInit();
            DateTime t2 = DateTime.Now;
            Console.WriteLine((t2 - t1).TotalSeconds);
            Console.WriteLine("You can edit excel now");
            Console.ReadLine();
        }
    }
}