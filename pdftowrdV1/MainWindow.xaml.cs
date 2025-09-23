using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Win32;
using PdfSharpCore.Pdf;
using PdfSharpCore.Pdf.IO;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection.PortableExecutable;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Windows;
using System.Windows.Forms;
using Windows.Globalization;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;
using Windows.Storage.Streams;

namespace PdfToWordArabicOCR
{
    public partial class MainWindow : Window
    {
        string pdfPath = "";
        string outputFolder = "";

        public MainWindow() => InitializeComponent();

        private void btnSelectPdf_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog { Filter = "PDF Files (*.pdf)|*.pdf" };
            if (dlg.ShowDialog() == true)
                pdfPath = dlg.FileName;
        }

        //private void btnSelectOutput_Click(object sender, RoutedEventArgs e)
        //{
        //    var dlg = new Microsoft.Win32.OpenFileDialog();
        //    bool? result = dlg.ShowDialog();

        //    if (result == true)
        //    {
        //        pdfPath = dlg.FileName;
        //    }
        //}

        private void btnSelectOutput_Click(object sender, RoutedEventArgs e)
        {
            using var dlg = new FolderBrowserDialog();
            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                outputFolder = dlg.SelectedPath;
            }
        }

        private async void btnStart_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(pdfPath) || string.IsNullOrEmpty(outputFolder))
            {
                System.Windows.MessageBox.Show("يرجى اختيار ملف PDF ومجلد الحفظ أولاً.");
                return;
            }

            txtStatus.Text = "جارٍ التحويل...";
            progressBar.Value = 0;

            var images = ConvertPdfToImages(pdfPath, outputFolder);
            var paragraphs = new List<string>();

            for (int i = 0; i < images.Count; i++)
            {
                string imgPath = images[i];
                string text = await ExtractTextFromImage(imgPath);
                paragraphs.Add(text);
                progressBar.Value = ((i + 1) * 100.0) / images.Count;
            }

            string wordPath = Path.Combine(outputFolder, "الناتج.docx");
            CreateWordFile(paragraphs, wordPath);

            txtStatus.Text = "تم التحويل بنجاح!";
            Process.Start("explorer.exe", outputFolder);
        }

        List<string> ConvertPdfToImages(string pdfPath, string outputFolder)
        {
            var images = new List<string>();
            using var document = PdfReader.Open(pdfPath, PdfDocumentOpenMode.ReadOnly);

            for (int i = 0; i < document.PageCount; i++)
            {
                int width = 1200;
                int height = 1600;
                using var bmp = new Bitmap(width, height);
                using var gfx = Graphics.FromImage(bmp);
                gfx.Clear(System.Drawing.Color.White);
                gfx.DrawString($"صفحة {i + 1} (رسم تجريبي)", new System.Drawing.Font("Arial", 24), System.Drawing.Brushes.Black, new System.Drawing.PointF(100, 100));

                string imgPath = Path.Combine(outputFolder, $"page_{i + 1}.png");
                bmp.Save(imgPath, ImageFormat.Png);
                images.Add(imgPath);
            }

            return images;
        }

        async System.Threading.Tasks.Task<string> ExtractTextFromImage(string imagePath)
        {
            var stream = new InMemoryRandomAccessStream();
            var bytes = File.ReadAllBytes(imagePath);
            var buffer = WindowsRuntimeBufferExtensions.AsBuffer(bytes);
            await stream.WriteAsync(buffer);
            stream.Seek(0);

            var decoder = await Windows.Graphics.Imaging.BitmapDecoder.CreateAsync(stream);
            var bitmap = await decoder.GetSoftwareBitmapAsync();

            var ocrEngine = Windows.Media.Ocr.OcrEngine.TryCreateFromLanguage(new Language("ar"));
            var result = await ocrEngine.RecognizeAsync(bitmap);

            return result.Text;
        }

        void CreateWordFile(List<string> paragraphs, string outputPath)
        {
            using var doc = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document);
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            foreach (var para in paragraphs)
            {
                var paragraph = new Paragraph(new Run(new Text(para)));
                mainPart.Document.Body.Append(paragraph);
            }

            doc.Save();
        }
    }
}