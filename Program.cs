using System;
using SW;

namespace ConsoleApp
{
    internal class Program
    {
        public static void Main(string[] args)
        {
            SwTools swTools = new SwTools();
            Console.Write(swTools.swConnect());
            Console.ReadLine();
        }
    }
}