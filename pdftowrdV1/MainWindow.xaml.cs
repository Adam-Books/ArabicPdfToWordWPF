using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
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

        private void btnSelectOutput_Click(object sender, RoutedEventArgs e)
        {
            using var dlg = new FolderBrowserDialog();
            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                outputFolder = dlg.SelectedPath;
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

            var images = await ConvertPdfToImagesAsync(pdfPath, outputFolder);
            if (images.Count == 0)
            {
                txtStatus.Text = "حدث خطأ أثناء تحويل PDF إلى صور.";
                return;
            }

            var paragraphs = new List<string>();
            var ocrEngine = OcrEngine.TryCreateFromLanguage(new Language("ar"));

            for (int i = 0; i < images.Count; i++)
            {
                string imgPath = images[i];
                string text = await ExtractTextFromImage(imgPath, ocrEngine);
                paragraphs.Add(text);
                progressBar.Value = ((i + 1) * 100.0) / images.Count;
            }

            string wordPath = Path.Combine(outputFolder, "الناتج.docx");
            CreateWordFile(paragraphs, wordPath);

            txtStatus.Text = "تم التحويل بنجاح!";
            Process.Start("explorer.exe", outputFolder);
        }

        async Task<List<string>> ConvertPdfToImagesAsync(string pdfPath, string outputFolder)
        {
            Directory.CreateDirectory(outputFolder);
            string gsPath = @"C:\Program Files\gs\gs10.06.0\bin\gswin64c.exe"; // عدّل حسب تثبيت Ghostscript
            if (!File.Exists(gsPath))
            {
                System.Windows.MessageBox.Show("لم يتم العثور على Ghostscript في المسار المحدد.");
                return new List<string>();
            }

            string outputPattern = Path.Combine(outputFolder, "page_%d.png");
            string args = $"-dNOPAUSE -dBATCH -sDEVICE=pngalpha -r600 -sOutputFile=\"{outputPattern}\" \"{pdfPath}\"";

            return await Task.Run(() =>
            {
                var psi = new ProcessStartInfo
                {
                    FileName = gsPath,
                    Arguments = args,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };

                using var proc = Process.Start(psi);
                if (proc != null)
                {
                    string output = proc.StandardOutput.ReadToEnd();
                    string error = proc.StandardError.ReadToEnd();
                    proc.WaitForExit();
                    if (proc.ExitCode != 0)
                        System.Windows.Application.Current.Dispatcher.Invoke(() =>
                            System.Windows.MessageBox.Show($"Ghostscript Error:\n{error}"));
                }

                return Directory.GetFiles(outputFolder, "page_*.png")
                    .OrderBy(f => int.Parse(Regex.Match(Path.GetFileNameWithoutExtension(f), @"\d+").Value))
                    .ToList();
            });
        }

        async Task<string> ExtractTextFromImage(string imagePath, OcrEngine ocrEngine)
        {
            var stream = new InMemoryRandomAccessStream();
            var bytes = File.ReadAllBytes(imagePath);
            await stream.WriteAsync(bytes.AsBuffer());
            stream.Seek(0);

            var decoder = await BitmapDecoder.CreateAsync(stream);
            var bitmap = await decoder.GetSoftwareBitmapAsync();

            var result = await ocrEngine.RecognizeAsync(bitmap);

            var paragraphs = new List<string>();
            var currentParagraph = new List<string>();
            double lastBottom = -1;

            foreach (var line in result.Lines)
            {
                if (line.Words == null || line.Words.Count == 0)
                    continue;

                string lineText = FixArabicLine(line.Text);
                var rect = line.Words[0].BoundingRect;

                bool isNewParagraph = false;

                // 1. مسافة بين الأسطر كبيرة
                if (lastBottom >= 0 && rect.Top - lastBottom > 25)
                    isNewParagraph = true;

                // 2. السطر يبدأ بمسافة (غالبًا مسافة بادئة في الكتاب الأصلي)
                if (lineText.StartsWith(" "))
                    isNewParagraph = true;

                // 3. السطر قصير جدًا (غالبًا عنوان أو بداية فقرة)
                if (lineText.Length < 15)
                    isNewParagraph = true;

                if (isNewParagraph && currentParagraph.Count > 0)
                {
                    paragraphs.Add(string.Join(" ", currentParagraph));
                    currentParagraph.Clear();
                }

                currentParagraph.Add(lineText);
                lastBottom = rect.Bottom;
            }

            if (currentParagraph.Count > 0)
                paragraphs.Add(string.Join(" ", currentParagraph));

            return string.Join(Environment.NewLine + Environment.NewLine, paragraphs);
        }

        string FixArabicLine(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return text;

            string normalized = text.Trim()
                                    .Replace(",", "،")
                                    .Replace(";", "؛")
                                    .Replace("?", "؟");

            var tokens = Regex.Split(normalized, @"\s+")
                              .Where(t => !string.IsNullOrWhiteSpace(t))
                              .ToArray();

            Array.Reverse(tokens);
            return "\u200F" + string.Join(" ", tokens);
        }

        void CreateWordFile(List<string> paragraphs, string outputPath)
        {
            using var doc = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document);
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            foreach (var para in paragraphs)
            {
                var paragraphNode = new Paragraph(
                    new ParagraphProperties(
                        new BiDi() { Val = OnOffValue.FromBoolean(true) },
                        new Justification() { Val = JustificationValues.Right },
                        new RightToLeftText() { Val = OnOffValue.FromBoolean(true) },

                        // ✨ بداية الفقرة بمسافة إضافية
                        new Indentation()
                        {
                            FirstLine = "720" // نصف بوصة تقريباً
                        }
                    ),
                    new Run(
                        new RunProperties(
                            new RunFonts()
                            {
                                Ascii = "adwa-assalaf",
                                HighAnsi = "adwa-assalaf",
                                ComplexScript = "adwa-assalaf"
                            },
                            new FontSize() { Val = "32" },
                            new RightToLeftText() { Val = OnOffValue.FromBoolean(true) }
                        ),
                        new Text(para) { Space = SpaceProcessingModeValues.Preserve }
                    )
                );

                mainPart.Document.Body.Append(paragraphNode);
            }

            doc.Save();
        }
    }
}
