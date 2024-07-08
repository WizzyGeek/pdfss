using UglyToad.PdfPig;
using PDFtoImage;
using System.Drawing;

namespace pdfss;

class Program
{
    static int Main(string[] args)
    {
        if (args.Length == 0) {
            Console.Write("Required a path\nUsage:\n\tpdfss <FILE>\n");
            return 0;
        }

        if (args.Length > 1) {
            Console.WriteLine("WARNING: Multiple arguments provided, only first considered");
        }

        string path = args[0];

        if (!File.Exists(path)) {
            Console.WriteLine("The provided file path does not exist");
            return 0;
        }

        float paddingX = 15; // Padding in points not pixels!
        float paddingY = 15;

        using PdfDocument pdf = PdfDocument.Open(path);
        using var fdoc = File.OpenRead(path);

        var pages = pdf.GetPages();

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
                    string name = $"img_{page.Number}_{ann.Name}.png";
                    Console.WriteLine($"[+] Saving image {name}");
                    #pragma warning disable CA1416 // Validate platform compatibility
                    Conversion.SavePng(
                        name,
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
                }
            }
        }
        return 0;
    }
}
