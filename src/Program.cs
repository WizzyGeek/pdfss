using UglyToad.PdfPig;
using PDFtoImage;
using System.Drawing;
using SkiaSharp;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.Util;
using System.ComponentModel.Design;

namespace pdfss;

class Program
{

    static void PrintHelp() {
        Console.WriteLine("Usage:\n\tpdfss <FILE> [-o/--output <FILE>]\n\n\t<FILE>\t\tThe PDF file to process\n\t-o/--output\tThe destination path of the generated .xlsx file");
    }

    static int Main(string[] args)
    {
        if (args.Length == 0) {
            Console.WriteLine("Required atleast a path");
            PrintHelp();
            return 0;
        }

        string path;
        string output;

        if (args.Length > 1) {
            int i;
            for(i = 0; i < args.Length - 1; i++) {
                if (args[i] == "-o" || args[i] == "--output") {
                    output = args[i + 1];
                    goto OutputFound;
                }
            }

            path = args[0];
            output = Path.GetFileNameWithoutExtension(path) + ".xlsx";
            Console.WriteLine("[!] WARNING: Multiple unknown arguments supplied!");
            goto OutputAndPathFound;

            OutputFound:
            for(int j = 0; j < args.Length; j++) {
                if (j == i || j == i + 1) continue;
                path = args[j];
                goto OutputAndPathFound;
            }

            Console.WriteLine("Required parameter <FILE> not found!");
            PrintHelp();
            return 0;

            OutputAndPathFound:;
        }
        else {
            path = args[0];
            output = Path.GetFileNameWithoutExtension(path) + ".xlsx";
        }


        if (!File.Exists(path)) {
            Console.WriteLine("The provided file path does not exist");
            return 0;
        }

        const float paddingX = 15; // Padding in points not pixels!
        const float paddingY = 15;
        const float magXL = 1.5F;

        using var fdoc = File.OpenRead(path);
        Console.WriteLine($"[*] Reading PDF file at {fdoc.Name}");

        using PdfDocument pdf = PdfDocument.Open(path);
        using IWorkbook workbook = new XSSFWorkbook();

        ISheet sheet1 = workbook.CreateSheet("sheet1");
        IDrawing patriarch = sheet1.CreateDrawingPatriarch();

        var pages = pdf.GetPages();

        int row = 1;

        IRow r = sheet1.CreateRow(0);
        r.CreateCell(0).SetCellValue("Comment");
        r.CreateCell(1).SetCellValue("Annotation");

        IFont font = workbook.CreateFont();
        font.IsBold = true;

        for (int i = 0; i <= 1; i++) {
            r.GetCell(i).RichStringCellValue.ApplyFont(font);
            r.GetCell(i).CellStyle.Alignment = HorizontalAlignment.Center;
            r.GetCell(i).CellStyle.VerticalAlignment = VerticalAlignment.Center;
            sheet1.SetColumnWidth(i, 25 * 256);
        }

        r.HeightInPoints = (float)Units.ToPoints(Units.EMU_PER_CENTIMETER);

        foreach (var page in pages) {
            double maxy = page.Height;
            foreach (var ann in page.ExperimentalAccess.GetAnnotations()) {
                double[] lims = [Double.PositiveInfinity, Double.NegativeInfinity, Double.PositiveInfinity, Double.NegativeInfinity];
                // min x - 0
                // max x - 1
                // min y - 2
                // max y - 3

                foreach (var q in ann.QuadPoints) {
                    foreach (var p in q.Points) {
                        if (p.X > lims[1]) lims[1] = p.X;
                        if (p.X < lims[0]) lims[0] = p.X;
                        if (p.Y > lims[3]) lims[3] = p.Y;
                        if (p.Y < lims[2]) lims[2] = p.Y;
                    }
                };

                // Do the affine transform, 2 is max y 3 is min y
                lims[2] = maxy - lims[2];
                lims[3] = maxy - lims[3];

                var bbox = new RectangleF(
                    float.Max((float)lims[0] - paddingX, 0),
                    float.Max((float)lims[3] - paddingY, 0),
                    float.Min((float)(lims[1] - lims[0]) + 2 * paddingX, (float)page.Width),
                    float.Min((float)(lims[2] - lims[3]) + 2 * paddingY, (float)maxy)
                );

                if (Double.IsFinite(bbox.Width) && Double.IsFinite(bbox.Height) && ann.Content != "") {
                    Console.WriteLine($"[+] Processing annotation {ann.Name} on page {page.Number}");
                    #pragma warning disable CA1416 // Validate platform compatibility

                    SKBitmap bmap = Conversion.ToImage(
                        fdoc,
                        leaveOpen: true,
                        page: page.Number - 1,
                        options: new(
                            DpiRelativeToBounds: true,
                            Bounds: bbox,
                            Dpi: 300,
                            WithAnnotations: true,
                            WithFormFill: true
                        )
                    );
                    #pragma warning restore CA1416 // Validate platform compatibility
                    IRow ri = sheet1.CreateRow(row);
                    ri.HeightInPoints = bbox.Height * magXL;
                    SKData dat = bmap.Encode(SKEncodedImageFormat.Png, 90);
                    XSSFClientAnchor anchor = new(0, 0, Units.ToEMU(bbox.Width * magXL), Units.ToEMU(bbox.Height * magXL), 1, row, 1, row++);
                    anchor.AnchorType = AnchorType.MoveDontResize;
                    XSSFPicture img = (XSSFPicture)patriarch.CreatePicture(anchor, workbook.AddPicture(dat.ToArray(), PictureType.PNG));
                    img.LineStyle = LineStyle.Solid;
                    img.SetLineStyleColor(0, 0, 0);
                    img.LineWidth = 1;
                    ri.CreateCell(0).SetCellValue(ann.Content);
                }
            }
        }

        using FileStream sw = File.Create(output);
        workbook.Write(sw, false);
        Console.WriteLine($"[*] Saved output at {sw.Name}");

        return 0;
    }
}
