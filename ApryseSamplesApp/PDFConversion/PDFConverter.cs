using Microsoft.VisualBasic.FileIO;
using pdftron;
using pdftron.PDF;
using pdftron.SDF;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static pdftron.PDF.Image;

namespace ApryseSDKSamples.PDFConversion
{
    internal class PDFConverter
    {
        private string ApryseSDKKey = "<SDKKEY>";

        private const string BasePath = @"<BASEPATH>";

        private string InputOfficeFilesFolder = BasePath + @"\ApryseSamples\ApryseSamples\InOfficeFiles";
        private string InputTextFilesFolder = BasePath + @"\ApryseSamples\ApryseSamples\InTextFiles";
        private string InputImgFilesFolder = BasePath + @"\ApryseSamples\ApryseSamples\InImageFiles";
        private string InputHtmlFilesFolder = BasePath + @"\ApryseSamples\ApryseSamples\InHtmlFiles";

        private string OutputPath = BasePath + @"\ApryseSamples\ApryseSamples\ConvertedOutput\";
        private string MergedOutputPath = BasePath + @"\ApryseSamples\ApryseSamples\MergedOutput\";

        public PDFConverter()
        {
            
        }

        public void DoWork()
        {
            Console.WriteLine("Started PDFConverter.DoWork()");

            Console.WriteLine("Initializing PDFNet");
            PDFNet.Initialize(ApryseSDKKey);
            Console.WriteLine("Successfully Initialized PDFNet");

            // Clean all files in output folder
            CleanOutputFolders();


            //Console.WriteLine("Converting Text files to Pdf");
            //ConvertTextFilesToPdf();
            //Console.WriteLine("Successfully converted Text files to Pdf");


            //Console.WriteLine("Converting SVG files to Pdf");
            //ConvertSVGFilesToPdf();
            //Console.WriteLine("Successfully converted SVG files to Pdf");


            //Console.WriteLine("Converting Image files to Pdf");
            //ConvertImageFilesToPdf();
            //Console.WriteLine("Successfully converted Image files to Pdf");



            //Console.WriteLine("Converting Office Documents to Pdf");
            //ConvertOfficeDocumentsToPdf();
            //Console.WriteLine("Successfully converted Office Documents to Pdf");


            Console.WriteLine("Converting Html Documents to Pdf");
            ConvertHtmlDocumentsToPdf();
            Console.WriteLine("Successfully converted Html Documents to Pdf");

            Console.WriteLine("Merging Pdfs");
            MergeCreatedPdfs();
            Console.WriteLine("Completed Merging Pdfs");


            Console.WriteLine("Completed PDFConverter.DoWork()");
        }

        private void CleanOutputFolders()
        {
            string[] filesToDelete = Directory.GetFiles(OutputPath);
            filesToDelete.ToList().ForEach(file => File.Delete(file));

            filesToDelete = Directory.GetFiles(MergedOutputPath);
            filesToDelete.ToList().ForEach(file => File.Delete(file));

        }

        private void ConvertOfficeDocumentsToPdf()
        {

            try
            {

                string supportedExtensions = "*.doc, *.docx, *.xls, *.xlsx, *.ppt, *.pptx";

                IEnumerable<string> files = Directory.GetFiles(InputOfficeFilesFolder, "*.*", System.IO.SearchOption.AllDirectories)
                    .Where(f => supportedExtensions.Contains(Path.GetExtension(f).ToLower()));

                foreach (string file in files)
                {
                    Console.WriteLine($"Converting Office document file {file}");
                    string outputFile = $"{OutputPath}{Guid.NewGuid().ToString()}.pdf";
                    using (PDFDoc doc = new PDFDoc())
                    {
                        doc.InitSecurityHandler();

                        OfficeToPDFOptions options = new OfficeToPDFOptions();
                        pdftron.PDF.Convert.OfficeToPDF(doc, file, options);
                        
                        doc.Save(outputFile, SDFDoc.SaveOptions.e_linearized);
                    }

                    Console.WriteLine($"Successfully generated PDF file {outputFile}");
                    Console.WriteLine($"***********");
                    Console.WriteLine($"");
                }

                files = Directory.GetFiles(InputOfficeFilesFolder, "*.rtf");

                foreach (string file in files)
                {
                    Console.WriteLine($"Converting RTF file {file}");
                    string outputFile = $"{OutputPath}{Guid.NewGuid().ToString()}.pdf";
                    using (PDFDoc doc = new PDFDoc())
                    {
                        doc.InitSecurityHandler();
                        pdftron.PDF.Convert.ToPdf(doc, file);
                        doc.Save(outputFile, SDFDoc.SaveOptions.e_remove_unused);
                    }

                    Console.WriteLine($"Successfully generated PDF file {outputFile}");
                    Console.WriteLine($"***********");
                    Console.WriteLine($"");
                }
            }
            catch( Exception exception)
            {
                Console.WriteLine($"Exception occured in ConvertOfficeDocumentsToPdf : {exception.Message}, {exception.StackTrace}");
            }
        }

        private void ConvertTextFilesToPdf()
        {

            try
            {

                ObjSet set = new ObjSet();
                Obj options = set.CreateDict();

                // Put options
                options.PutNumber("FontSize", 15);
                options.PutBool("UseSourceCodeFormatting", true);
                options.PutNumber("PageWidth", 12);
                options.PutNumber("PageHeight", 6);


                string[] files = Directory.GetFiles(InputTextFilesFolder);
                foreach (string file in files)
                {
                    Console.WriteLine($"Converting Text file {file}");
                    string outputFile = $"{OutputPath}{Guid.NewGuid().ToString()}.pdf";
                    using (PDFDoc doc = new PDFDoc())
                    {
                        doc.InitSecurityHandler();
                        pdftron.PDF.Convert.FromText(doc, file, options);
                        doc.Save(outputFile, SDFDoc.SaveOptions.e_remove_unused);
                    }

                    Console.WriteLine($"Successfully generated PDF file {outputFile}");
                    Console.WriteLine($"***********");
                    Console.WriteLine($"");
                }
            }
            catch (Exception exception)
            {
                Console.WriteLine($"Exception occured in ConvertTextFilesToPdf : {exception.Message}, {exception.StackTrace}");
            }
        }

        private void ConvertSVGFilesToPdf()
        {

            try
            {

                string[] files = Directory.GetFiles(InputImgFilesFolder, "*.svg");
                foreach (string file in files)
                {
                    Console.WriteLine($"Converting SVG file {file}");
                    string outputFile = $"{OutputPath}{Guid.NewGuid().ToString()}.pdf";
                    using (PDFDoc doc = new PDFDoc())
                    {
                        doc.InitSecurityHandler();
                        pdftron.PDF.Convert.FromSVG(doc, file, null);
                        doc.Save(outputFile, SDFDoc.SaveOptions.e_remove_unused);
                    }

                    Console.WriteLine($"Successfully generated PDF file {outputFile}");
                    Console.WriteLine($"***********");
                    Console.WriteLine($"");
                }
            }
            catch (Exception exception)
            {
                Console.WriteLine($"Exception occured in ConvertSVGFilesToPdf : {exception.Message}, {exception.StackTrace}");
            }
        }

        private void ConvertImageFilesToPdf()
        {

            try
            {
                string supportedExtensions = "*.jpg,*.gif,*.png,*.bmp,*.jpe,*.jpeg,*.tif,*.tiff";

                IEnumerable<string> files = Directory.GetFiles(InputImgFilesFolder, "*.*", System.IO.SearchOption.AllDirectories)
                    .Where(f => supportedExtensions.Contains(Path.GetExtension(f).ToLower()));

                foreach (string file in files)
                {
                    Console.WriteLine($"Converting Image file {file}");
                    string outputFile = $"{OutputPath}{Guid.NewGuid().ToString()}.pdf";
                    using (PDFDoc doc = new PDFDoc())
                    {
                        doc.InitSecurityHandler();
                        pdftron.PDF.Convert.ToPdf(doc, file);
                        doc.Save(outputFile, SDFDoc.SaveOptions.e_remove_unused);
                    }

                    Console.WriteLine($"Successfully generated PDF file {outputFile}");
                    Console.WriteLine($"***********");
                    Console.WriteLine($"");
                }
            }
            catch (Exception exception)
            {
                Console.WriteLine($"Exception occured in ConvertImageFilesToPdf : {exception.Message}, {exception.StackTrace}");
            }
        }

        private void ConvertHtmlDocumentsToPdf()
        {

            try
            {
                if (!HTML2PDF.IsModuleAvailable())
                {
                    Console.WriteLine("HTML2PDF Module not loaded !");
                }

                string[] files = Directory.GetFiles(InputHtmlFilesFolder);
                foreach (string file in files)
                {
                    Console.WriteLine($"Converting Html file {file}");
                    string outputFile = $"{OutputPath}{Guid.NewGuid().ToString()}.pdf";

                    string htmlFileContent = File.ReadAllText(file);
                    using (PDFDoc doc = new PDFDoc())
                    {
                        doc.InitSecurityHandler();

                        HTML2PDF converter = new HTML2PDF();
                        converter.SetPaperSize(PrinterMode.PaperSize.e_11x17);
                        converter.InsertFromHtmlString(htmlFileContent);

                        if (converter.Convert(doc))
                        {


                            doc.Save(outputFile, SDFDoc.SaveOptions.e_linearized);
                        }
                        else
                        {
                            Console.WriteLine("Conversion failed. HTTP Code: {0}\n{1}", converter.GetHTTPErrorCode(), converter.GetLog());
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
                if ( Directory.GetFiles(OutputPath, "*.pdf").Length == 0)
                {
                    Console.WriteLine("Not files to merge");
                    return;
                }

                using (PDFDoc outputDoc = new PDFDoc())
                {
                    outputDoc.InitSecurityHandler();
                    int outputDocPageCtr = 1;

                    string[] files = Directory.GetFiles(OutputPath, "*.pdf");
                    foreach (string file in files)
                    {
                        using (PDFDoc inDoc = new PDFDoc(file))
                        {
                            int inDocPageCount = inDoc.GetPageCount();
                            outputDoc.InsertPages(outputDocPageCtr, inDoc, 1, inDocPageCount, PDFDoc.InsertFlag.e_none);
                            outputDocPageCtr += inDocPageCount;
                        }
                    }

                    outputDoc.Save($"{MergedOutputPath}{Guid.NewGuid().ToString()}.pdf", SDFDoc.SaveOptions.e_remove_unused);

                }
            }
            catch (Exception exception)
            {
                Console.WriteLine($"Exception occured in MergeCreatedPdfs : {exception.Message}, {exception.StackTrace}");
            }
        }
    }
}
