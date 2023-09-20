namespace ITextSamples
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Started ITextSamples App");
            Console.WriteLine("**************************");

            PDFConversion.PDFConverter instance = new PDFConversion.PDFConverter();

            instance.DoWork();

            Console.WriteLine("**************************");
            Console.WriteLine("Completed ITextSamples App");
        }
    }
}