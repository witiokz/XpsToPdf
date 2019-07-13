using PdfSharp;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using System;
using System.IO;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Xps.Packaging;

namespace xps_to_pdf
{
    class Program
    {
        static readonly string DestinationExtension = "pdf";

        [STAThread]
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Incorrect arguments count. Example: <source directory with xps files> <destination of pdf files> <pageSize>");
                return;
            }

            var source = args[0];
            var destination = args[1];

            if (!Directory.Exists(source) || !Directory.Exists(destination))
            {
                Console.WriteLine("Source and/or destination directory do not exist");
                return;
            }

            var pageSize = args.Length >= 3 ? GetPageSize(args[2]) : PageSize.A4;

            var sourceFiles = Directory.EnumerateFiles(source, "*.xps");

            foreach (var sourceFile in sourceFiles)
            {
                var destFile = Path.Combine(destination, $"{Path.GetFileNameWithoutExtension(sourceFile)}.{DestinationExtension}");
                XpsToBmp(sourceFile, destFile, pageSize);
            }
        }

        static public void XpsToBmp(string xpsFile, string source, PageSize pageSize)
        {
            var xps = new XpsDocument(xpsFile, FileAccess.Read);
            var sequence = xps.GetFixedDocumentSequence();

            using (var doc = new PdfDocument())
            {
                for (var pageCount = 0; pageCount < sequence.DocumentPaginator.PageCount; ++pageCount)
                {
                    DocumentPage page = sequence.DocumentPaginator.GetPage(pageCount);
                    var toBitmap = new RenderTargetBitmap((int)page.Size.Width, (int)page.Size.Height, 96, 96, PixelFormats.Default);
                    toBitmap.Render(page.Visual);

                    var bmpEncoder = new BmpBitmapEncoder();
                    bmpEncoder.Frames.Add(BitmapFrame.Create(toBitmap));

                    using (var ms = new MemoryStream())
                    {
                        bmpEncoder.Save(ms);

                        doc.Pages.Add(new PdfPage { Size = pageSize, Orientation = GetOrientation((int)page.Size.Width) });

                        var xgr = XGraphics.FromPdfPage(doc.Pages[pageCount]);
                        var img = XImage.FromStream(ms);

                        xgr.DrawImage(img, 0, 0);
                    }
                }

                doc.Save(source);
            }
        }

        static PageSize GetPageSize(string pageSize)
        {
            return (PageSize)Enum.Parse(typeof(PageSize), pageSize);
        }

        static PageOrientation GetOrientation(int width)
        {
            return width > 794 ? PageOrientation.Landscape : PageOrientation.Portrait;
        }
    }
}
