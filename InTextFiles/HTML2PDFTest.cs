//---------------------------------------------------------------------------------------
// Copyright (c) 2001-2023 by Apryse Software Inc. All Rights Reserved.
// Consult legal.txt regarding legal and license information.
//---------------------------------------------------------------------------------------

using System;
using System.IO;

using pdftron;
using pdftron.Common;
using pdftron.SDF;
using pdftron.PDF;

namespace HTML2PDFTestCS
{
	//---------------------------------------------------------------------------------------
	// The following sample illustrates how to convert HTML pages to PDF format using
	// the HTML2PDF class.
	// 
	// 'pdftron.PDF.HTML2PDF' is an optional PDFNet Add-On utility class that can be 
	// used to convert HTML web pages into PDF documents by using an external module (html2pdf).
	//
	// html2pdf modules can be downloaded from http://www.pdftron.com/pdfnet/downloads.html.
	//
	// Users can convert HTML pages to PDF using the following operations:
	// - Simple one line static method to convert a single web page to PDF. 
	// - Convert HTML pages from URL or string, plus optional table of contents, in user defined order. 
	// - Optionally configure settings for proxy, images, java script, and more for each HTML page. 
	// - Optionally configure the PDF output, including page size, margins, orientation, and more. 
	// - Optionally add table of contents, including setting the depth and appearance.
	//---------------------------------------------------------------------------------------
	class HTML2PDFSample
	{
		private static pdftron.PDFNetLoader pdfNetLoader = pdftron.PDFNetLoader.Instance();
		static HTML2PDFSample() {}
		
		static void Main(string[] args)
		{
			string output_path = "../../TestFiles/Output/html2pdf_example";
			string host = "https://docs.apryse.com";
			string page0 = "/";
			string page1 = "/all-products/";
			string page2 = "/documentation/web/faq";

			// The first step in every application using PDFNet is to initialize the 
			// library and set the path to common PDF resources. The library is usually 
			// initialized only once, but calling Initialize() multiple times is also fine.
			//PDFNet.Initialize(PDFTronLicense.Key);
			PDFNet.Initialize("demo:1694739843838:7c08853e03000000008c3c8e97210e69cc825f36a851bfb2804332cc1f");
			// For HTML2PDF we need to locate the html2pdf module. If placed with the 
			// PDFNet library, or in the current working directory, it will be loaded
			// automatically. Otherwise, it must be set manually using HTML2PDF.SetModulePath().
			HTML2PDF.SetModulePath("../../../Lib");
			if (!HTML2PDF.IsModuleAvailable())
			{
				Console.WriteLine();
				Console.WriteLine("Unable to run HTML2PDFTest: Apryse SDK HTML2PDF module not available.");
				Console.WriteLine("---------------------------------------------------------------");
				Console.WriteLine("The HTML2PDF module is an optional add-on, available for download");
				Console.WriteLine("at http://www.pdftron.com/. If you have already downloaded this");
				Console.WriteLine("module, ensure that the SDK is able to find the required files");
				Console.WriteLine("using the HTML2PDF.SetModulePath() function.");
				Console.WriteLine();
				return;
			}
			
			//--------------------------------------------------------------------------------
			// Example 1) Simple conversion of a web page to a PDF doc. 

			try
			{
				using (PDFDoc doc = new PDFDoc())
				{

					// now convert a web page, sending generated PDF pages to doc
					HTML2PDF.Convert(doc, host + page0);
					doc.Save(output_path + "_01.pdf", SDFDoc.SaveOptions.e_linearized);
				}
			}
			catch (PDFNetException e)
			{
				Console.WriteLine(e.Message);
			}

			//--------------------------------------------------------------------------------
			// Example 2) Modify the settings of the generated PDF pages and attach to an
			// existing PDF document. 

			try
			{
				// open the existing PDF, and initialize the security handler
				using (PDFDoc doc = new PDFDoc("../../TestFiles/numbered.pdf"))
				{
					doc.InitSecurityHandler();

					// create the HTML2PDF converter object and modify the output of the PDF pages
					HTML2PDF converter = new HTML2PDF();
					converter.SetPaperSize(PrinterMode.PaperSize.e_11x17);

					// insert the web page to convert
					converter.InsertFromURL(host + page0);

					// convert the web page, appending generated PDF pages to doc
					converter.Convert(doc);
					doc.Save(output_path + "_02.pdf", SDFDoc.SaveOptions.e_linearized);
				}
			}
			catch (PDFNetException e)
			{
				Console.WriteLine(e.Message);
			}

			//--------------------------------------------------------------------------------
			// Example 3) Convert multiple web pages

			try
			{
				using (PDFDoc doc = new PDFDoc())
				{
					// convert page 0 into pdf
					HTML2PDF converter = new HTML2PDF();
					string header = "<div style='width:15%;margin-left:0.5cm;text-align:left;font-size:10px;color:#0000FF'><span class='date'></span></div><div style='width:70%;direction:rtl;white-space:nowrap;overflow:hidden;text-overflow:clip;text-align:center;font-size:10px;color:#0000FF'><span>PDFTRON HEADER EXAMPLE</span></div><div style='width:15%;margin-right:0.5cm;text-align:right;font-size:10px;color:#0000FF'><span class='pageNumber'></span> of <span class='totalPages'></span></div>";
					string footer = "<div style='width:15%;margin-left:0.5cm;text-align:left;font-size:7px;color:#FF00FF'><span class='date'></span></div><div style='width:70%;direction:rtl;white-space:nowrap;overflow:hidden;text-overflow:clip;text-align:center;font-size:7px;color:#FF00FF'><span>PDFTRON FOOTER EXAMPLE</span></div><div style='width:15%;margin-right:0.5cm;text-align:right;font-size:7px;color:#FF00FF'><span class='pageNumber'></span> of <span class='totalPages'></span></div>";
					converter.SetHeader(header);
					converter.SetFooter(footer);
					converter.SetMargins("1cm", "2cm", ".5cm", "1.5cm");

					HTML2PDF.WebPageSettings settings = new HTML2PDF.WebPageSettings();
					settings.SetZoom(0.5);
					converter.InsertFromURL(host + page0, settings);
					converter.Convert(doc);

					// convert page 1 with the same settings, appending generated PDF pages to doc
					converter.InsertFromURL(host + page1, settings);
					converter.Convert(doc);

					// convert page 2 with different settings, appending generated PDF pages to doc
					HTML2PDF another_converter = new HTML2PDF();
					another_converter.SetLandscape(true);
					HTML2PDF.WebPageSettings another_settings = new HTML2PDF.WebPageSettings();
					another_settings.SetPrintBackground(false);
					another_converter.InsertFromURL(host + page2, another_settings);
					another_converter.Convert(doc);

					doc.Save(output_path + "_03.pdf", SDFDoc.SaveOptions.e_linearized);
				}
			}
			catch (PDFNetException e)
			{
				Console.WriteLine(e.Message);
			}

			//--------------------------------------------------------------------------------
			// Example 4) Convert HTML string to PDF. 

			try
			{
				using (PDFDoc doc = new PDFDoc())
				{

					HTML2PDF converter = new HTML2PDF();
				
					// Our HTML data
					string html = "<html><body><h1>Heading</h1><p>Paragraph.</p></body></html>";
					
					// Add html data
					converter.InsertFromHtmlString(html);
					// Note, InsertFromHtmlString can be mixed with the other Insert methods.

					converter.Convert(doc);
					doc.Save(output_path + "_04.pdf", SDFDoc.SaveOptions.e_linearized);
				}
			}
			catch (PDFNetException e)
			{
				Console.WriteLine(e.Message);
			}

			//--------------------------------------------------------------------------------
			// Example 5) Set the location of the log file to be used during conversion.

			try
			{
				using (PDFDoc doc = new PDFDoc())
				{
					HTML2PDF converter = new HTML2PDF();
					converter.SetLogFilePath("../../TestFiles/Output/html2pdf.log");
					converter.InsertFromURL(host + page0);
					converter.Convert(doc);
					doc.Save(output_path + "_05.pdf", SDFDoc.SaveOptions.e_linearized);
				}
			}
			catch (PDFNetException e)
			{
				Console.WriteLine(e.Message);
			}

			PDFNet.Terminate();
		}
	}
}
