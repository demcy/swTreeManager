using System;
using System.IO;
using System.Windows.Forms;

namespace ConsoleApp_Revision
{
    internal class Program
    {
        public static void Main(string[] args)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog(); 
            openFileDialog1.ShowDialog();  
        }
    }
}