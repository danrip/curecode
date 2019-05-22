using System;
using System.IO;
using Simplexcel;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;
using ZXing;
using ZXing.Common;



namespace CureCode
{
    class Program
    {
        static void Main(string[] args)
        {
            var code = CreateQrCodeImage("Hello word", 100, 100);
            File.WriteAllBytes("curecode.png", code);

            var excel = CreateExcel();
            File.WriteAllBytes("excel.xlsx", excel);

            Console.WriteLine("Done, press Enter to exit");
            Console.ReadLine();
        }

        private static byte[] CreateQrCodeImage(string content, int height, int width)
        {
            byte[] result;

            var barcodeWriter = new BarcodeWriterPixelData()
            {
                Format = BarcodeFormat.QR_CODE,
                Options = new EncodingOptions() 
                { 
                    Height = height,
                    Width = width
                }
            };

            var pixelData = barcodeWriter.Write(content);
            using(var image = Image.LoadPixelData<Rgba32>(pixelData.Pixels, pixelData.Width, pixelData.Height))
            using(var ms = new MemoryStream())
            {
                image.SaveAsPng(ms);
                result = ms.ToArray();
            }

            return result;
        }

        private static byte[] CreateExcel()
        {
            var workBook = new Workbook();
            for (int i = 0; i < 100; i++)
            {
                if(i%25 > 0)
                {
                    Console.WriteLine("Adding code to page");
                }
                else
                {
                    workBook.Add(new Worksheet("page : " + i));
                }
            }
        
            
            using(var ms = new MemoryStream())
            {
                workBook.Save(ms);
                return ms.ToArray();
            }
        }

    }
}
