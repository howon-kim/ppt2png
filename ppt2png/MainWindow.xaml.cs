using System;
using System.IO;
using System.Linq;
using System.Windows;
using Microsoft.Office.Interop.PowerPoint;

namespace ppt2png
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            AllowDrop = true;
            DragEnter += MainWindow_DragEnter;
            Drop += MainWindow_Drop;
        }

        private void MainWindow_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effects = DragDropEffects.Copy;
            else
                e.Effects = DragDropEffects.None;
        }

        private void MainWindow_Drop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            string pptFile = files.FirstOrDefault(f => f.ToLower().EndsWith(".ppt") || f.ToLower().EndsWith(".pptx"));

            if (pptFile != null)
            {
                ConvertPowerPointToPNG(pptFile);
            }
            else
            {
                MessageBox.Show("Please drop a PowerPoint file.", "Invalid File", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ConvertPowerPointToPNG(string pptFile)
        {
            Microsoft.Office.Interop.PowerPoint.Application pptApp = new Microsoft.Office.Interop.PowerPoint.Application();
            Presentation presentation = pptApp.Presentations.Open(pptFile, WithWindow: Microsoft.Office.Core.MsoTriState.msoFalse);

            string outputDirectory = Path.GetDirectoryName(pptFile);
            string pptFileNameWithoutExtension = Path.GetFileNameWithoutExtension(pptFile);

            foreach (Slide slide in presentation.Slides)
            {
                // Get the current slide's size
                float width = slide.Master.Width;
                float height = slide.Master.Height;

                // Create a larger image
                float scaleFactor = 8.0f;
                int newWidth = (int)(width * scaleFactor);
                int newHeight = (int)(height * scaleFactor);

                // Construct the output filename with the PowerPoint file's name
                string pngFileName = $"{pptFileNameWithoutExtension}_Slide_{slide.SlideIndex}.png";
                string pngFile = Path.Combine(outputDirectory, pngFileName);

                // Export the slide as PNG with the new size
                slide.Export(pngFile, "PNG", newWidth, newHeight);
            }

            presentation.Close();
            pptApp.Quit();

            MessageBox.Show("Conversion complete. PNG files saved to:\n" + outputDirectory, "Conversion Complete", MessageBoxButton.OK, MessageBoxImage.Information);
        }


    }
}
