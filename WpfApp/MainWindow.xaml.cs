using System;
using System.Diagnostics;
using System.IO;
using System.Windows;
using Microsoft.Win32;
using System.Windows.Forms;
using DrawRevision;
using MessageBox = System.Windows.Forms.MessageBox;
using OpenFileDialog = System.Windows.Forms.OpenFileDialog;


namespace WpfApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private string[] solidFiles;
        private string drawingPath;
        
        private void SolidFiles_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = @"S:\Programs\Macros\SW TreeManager\19-816-04 - Additional Items";
            openFileDialog.Multiselect = true;
            openFileDialog.Filter = "SolidWorks Drawings (*.slddrw)|*.slddrw";
            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                solidFiles = openFileDialog.FileNames;
            }
        }

        private void GetDirectory_Click(object sender, RoutedEventArgs e)
        {
            var folderBrowserDialog = new FolderBrowserDialog();
            folderBrowserDialog.SelectedPath = @"S:\Programs\Macros\SW TreeManager\19-816-04 - Additional Items";
            folderBrowserDialog.ShowDialog();
            drawingPath = folderBrowserDialog.SelectedPath + @"\";
        }

        private void Ready_Click(object sender, RoutedEventArgs e)
        {
            if(solidFiles == null)
            {
                MessageBox.Show("Choose SolidWorks drawings first!");
            }
            else if (drawingPath == null)
            {
                MessageBox.Show("Choose PDF drawings directory first!");
            }
            else
            {
                var revision = new Revision();
                revision.GetFiles(solidFiles, drawingPath);
            }
        }
    }
}