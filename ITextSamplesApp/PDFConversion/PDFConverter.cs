using iText.Html2pdf;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using iText.Pdfoptimizer;
using iText.Pdfoptimizer.Handlers.Imagequality.Processors;
using iText.Pdfoptimizer.Handlers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iText.Licensing.Base;

namespace ITextSamples.PDFConversion
{
    internal class PDFConverter
    {


        private string InputHtmlFilesFolder = @"<BASEPATH>\ApryseSamples\ApryseSamples\InHtmlFiles";
        private const string BasePath = @"<BASEPATH>";

        private string LicenseFilePath = BasePath + @"\ApryseSamples\<LICENSEFILE>.json";
        private string OutputPath = BasePath + @"\ApryseSamples\ApryseSamples\ConvertedOutput\";
        private string MergedOutputPath = BasePath + @"\ApryseSamples\ApryseSamples\MergedOutput\";
        private string OptimizedOutputPath = BasePath + @"\ApryseSamples\ApryseSamples\OptimizedOutput";

        public PDFConverter()
        {
            
        }

        public void DoWork()
        {
            Console.WriteLine("Started PDFConverter.DoWork()");

            // Clean all files in output folder
            CleanOutputFolders();


            Console.WriteLine("Converting Html Documents to Pdf");
            ConvertHtmlDocumentsToPdf();
            Console.WriteLine("Successfully converted Html Documents to Pdf");



            Console.WriteLine("Merging Pdfs");
            MergeCreatedPdfs();
            Console.WriteLine("Completed Merging Pdfs");

            Console.WriteLine("Optimizing PDF");
            OptimizePdfs();
            Console.WriteLine("Completed Optimizing PDFs");

            Console.WriteLine("Completed PDFConverter.DoWork()");
        }

        private void CleanOutputFolders()
        {
            string[] filesToDelete = Directory.GetFiles(OutputPath);
            filesToDelete.ToList().ForEach(file => File.Delete(file));

            filesToDelete = Directory.GetFiles(MergedOutputPath);
            filesToDelete.ToList().ForEach(file => File.Delete(file));

            filesToDelete = Directory.GetFiles(OptimizedOutputPath);
            filesToDelete.ToList().ForEach(file => File.Delete(file));

        }

        private void ConvertHtmlDocumentsToPdf()
        {

            try
            {

                string[] files = Directory.GetFiles(InputHtmlFilesFolder);
                foreach (string file in files)
                {
                    Console.WriteLine($"Converting Html file {file}");
                    string outputFile = $"{OutputPath}{Guid.NewGuid().ToString()}.pdf";

                    string htmlFileContent = File.ReadAllText(file);
                    using ( FileStream inputStream = File.Open(file,FileMode.Open))
                    {
                        using (FileStream outputStream = File.Open(outputFile, FileMode.Create))
                        {
                            ConverterProperties converterProperties = new ConverterProperties();
                            HtmlConverter.ConvertToPdf(inputStream, outputStream, converterProperties);
                        }
                    }
                    Console.WriteLine($"Successfully generated PDF file {outputFile}");
                    Console.WriteLine($"***********");
                    Console.WriteLine($"");
                }
            }
            catch (Exception exception)
            {
                Console.WriteLine($"Exception occured in ConvertHtmlDocumentsToPdf : {exception.Message}, {exception.StackTrace}");
            }
        }

        private void MergeCreatedPdfs()
        {
            try
            {
                if (Directory.GetFiles(OutputPath, "*.pdf").Length == 0)
                {
                    Console.WriteLine("Not files to merge");
                    return;
                }

                string mergedOutputPdf = $"{MergedOutputPath}{Guid.NewGuid().ToString()}.pdf";
                string[] files = Directory.GetFiles(OutputPath, "*.pdf");

                // Create output pdf
                using (iText.Kernel.Pdf.PdfDocument mergedPdfDocument = new iText.Kernel.Pdf.PdfDocument(new PdfWriter(new FileStream(mergedOutputPdf, FileMode.Create, FileAccess.Write))))
                {
                    // create ouput document
                    using (iText.Layout.Document mergedDocument = new iText.Layout.Document(mergedPdfDocument))
                    {

                        // Read all input pdf files
                        foreach (var file in files)
                        {
                            int totalPages = 1;
                            // Open input pdf file
                            using (iText.Kernel.Pdf.PdfDocument sourcePdfDocument = new iText.Kernel.Pdf.PdfDocument(new PdfReader(new FileStream(file, FileMode.Open, FileAccess.Read))))
                            {
                                
                                // copy all pages from source to target
                                sourcePdfDocument.CopyPagesTo(1, sourcePdfDocument.GetNumberOfPages(), mergedPdfDocument);
                                totalPages += sourcePdfDocument.GetNumberOfPages();

                                // Add a new blank page
                                mergedPdfDocument.AddNewPage();
                            }
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                Console.WriteLine($"Exception occured in MergeCreatedPdfs : {exception.Message}, {exception.StackTrace}");
            }
        }

        private void OptimizePdfs()
        {
            try
            {

                string[] files = Directory.GetFiles(MergedOutputPath, "*.pdf");

                if (files.Length <= 0)
                {
                    Console.WriteLine("No files to optimize");
                    return;
                }


                LicenseKey.LoadLicenseFile(new FileInfo(LicenseFilePath));

                PdfOptimizer pdfOptimizer = new PdfOptimizer();
                pdfOptimizer.AddOptimizationHandler(new FontDuplicationOptimizer());
                pdfOptimizer.AddOptimizationHandler(new CompressionOptimizer());

                ImageQualityOptimizer jpeg_optimizer = new ImageQualityOptimizer();
                jpeg_optimizer.SetJpegProcessor(new JpegCompressor(.5f));
                pdfOptimizer.AddOptimizationHandler(jpeg_optimizer);

                FileInfo inputFileInfo = new FileInfo(files[0]);

                string optimizedPdfFileName = $"{OptimizedOutputPath}\\{inputFileInfo.Name}";
                string finalPdfFileName = $"{OptimizedOutputPath}\\FINAL-{inputFileInfo.Name}";

                pdfOptimizer.Optimize(new FileInfo(files[0]), new FileInfo(optimizedPdfFileName));

                Console.WriteLine("Optimized PDF.");

                Console.WriteLine("Removing blank pages");
                using (iText.Kernel.Pdf.PdfDocument sourcePdfDocument = new iText.Kernel.Pdf.PdfDocument(new PdfReader(new FileStream(optimizedPdfFileName, FileMode.Open, FileAccess.Read))))
                {
                    // Create output pdf
                    using (iText.Kernel.Pdf.PdfDocument finalPdfDocument = new iText.Kernel.Pdf.PdfDocument(new PdfWriter(new FileStream(finalPdfFileName, FileMode.Create, FileAccess.Write))))
                    {
                        // create ouput document
                        using (iText.Layout.Document finalDocument = new iText.Layout.Document(finalPdfDocument))
                        {
                            List<int> listOfPages = new List<int>();
                            int pageCounts = sourcePdfDocument.GetNumberOfPages();
                            for (int i = 1; i <= pageCounts; i++)
                            {
                                PdfPage page = sourcePdfDocument.GetPage(i);
                                byte[] bytes = page.GetContentBytes();
                                Console.WriteLine($"Page number = {i}, bytes count = {bytes.Count()}");
                                if (bytes.Length > 0)
                                {
                                    listOfPages.Add(i);
                                }
                            }
                            sourcePdfDocument.CopyPagesTo(listOfPages, finalPdfDocument);
                        }
                    }
                }
                Console.WriteLine("Completed removing blank pages");
            }
            catch (Exception exception)
            {
                Console.WriteLine($"Exception occured in OptimizePdfs : {exception.Message}, {exception.StackTrace}");
            }
        }
    }
}
