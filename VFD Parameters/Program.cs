using IronPdf;
   


namespace VFD_Parameters
{
    internal class Program
    {
        private const string PdfFilePath = @"R:\Quotes\Mseries\released\744983 LML Laundry TI.pdf";
           
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {

            string modelSize = "";
            string blastDirection = "";
            string gasType = "";
            string jobName = "";
            string shopOrderNumber = "";
            string jobQuantity = "";
            string date = DateTime.Now.ToString(@"MM\/dd\/yyyy");
            string createdBy = "KRP";
            string horsePower = "";
            string voltage = "";
            string phase = "";
            string FLA = "";
            string RPM = "";
            string CFM = "";
            string TESP = "";
            string maxCFM = "";
            string manifoldPressureGEO = "";
            string manifoldPressureMax = "";


            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            //ApplicationConfiguration.Initialize();
            //Application.Run(new Form1());


            // Extracting image and text content from PDF Document
            using PdfDocument PDF = PdfDocument.FromFile(PdfFilePath);
            // Get all text to put in a search index
            string AllText = PDF.ExtractAllText();
           
            //Search text to grab needed data
            modelSize = AllText.Substring(AllText.LastIndexOf("Model:") + 4);

            Console.WriteLine(modelSize);
            Console.ReadKey();


        }

       

    }
}