using System;
using CORE;


namespace ConsoleApp
{
    internal class Program
    {
        public static void Main(string[] args)
        {
            Core core = new Core();
            core.swInit();
            Console.WriteLine("FINISH");
            Console.ReadLine();
        }
    }
}