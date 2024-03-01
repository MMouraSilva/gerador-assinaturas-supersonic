using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Html.Converters;
using Aspose.Html.Rendering.Image;
using Aspose.Html.Saving;
using Aspose.Html;
using Aspose.Cells;
using System.Drawing;
using System.IO;
using System.Drawing.Drawing2D;

// TODO: Apply clean code concepts, convert to OOP

namespace Assinatura
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Getting HTML file
            string fileName = "Assinatura.html";
            // HTML file stored in an array
            string[] arrLine = File.ReadAllLines(fileName);

            // Get all sheets
            Workbook wb = new Workbook("C:\\Users\\Matheus Moura\\Pictures\\ASSINATURAS\\SOURCE\\source.xlsx");

            WorksheetCollection collection = wb.Worksheets;

            // Get worksheet by index
            Worksheet worksheet = collection[0];

            //Console.WriteLine("Worksheet: " + worksheet.Name);

            // Get number of rows and columns
            int rows = worksheet.Cells.MaxDataRow;
            int cols = worksheet.Cells.MaxDataColumn;

            string[] nome = new string[rows + 1];
            string[] setor = new string[rows + 1];
            string[] mobile = new string[rows + 1];
            string[] phone = new string[rows + 1];
            string[] qrcode = new string[rows + 1];
            string[] email = new string[rows + 1];

            // Run through each row
            for (int i = 1; i <= rows; i++) {
                // Run through each column in the selected row
                for (int j = 0; j <= cols; j++) {
                    // Saving the cell's value into the arrays
                    switch (j) {
                        case 0:
                            nome[i] = worksheet.Cells[i, j].Value.ToString();
                            break;
                        case 2:
                            setor[i] = worksheet.Cells[i, j].Value.ToString();
                            break;
                        case 3:
                            email[i] = worksheet.Cells[i, j].Value.ToString();
                            break;
                        case 4:
                            phone[i] = worksheet.Cells[i, j].Value.ToString();
                            break;
                        case 5:
                            mobile[i] = worksheet.Cells[i, j].Value.ToString();
                            break;
                        case 6:
                            qrcode[i] = worksheet.Cells[i, j].Value.ToString();
                            break;
                        default:
                            break;
                    }
                }
            }

            for (int i = 1; i <= rows; i++) {

                arrLine[239] = nome[i];
                arrLine[242] = setor[i];

                if (mobile[i] != "-") {
                    arrLine[249] = "Mobile: ";
                    arrLine[250] = mobile[i];
                } else {
                    arrLine[249] = "";
                    arrLine[250] = "";
                }

                if (phone[i] != "-") {
                    arrLine[247] = phone[i];
                    arrLine[246] = "Phone: ";

                    if (mobile[i] != "-") {
                        arrLine[248] = " | ";
                    } else {
                        arrLine[248] = "";
                    }
                } else {
                    arrLine[247] = "";
                    arrLine[246] = "";
                    arrLine[248] = "";
                }

                arrLine[253] = email[i];

                if (qrcode[i] != "-") {
                    arrLine[226] = "<img src=\"https://api.qrserver.com/v1/create-qr-code/?data=https://wa.me/55";
                    arrLine[227] = qrcode[i];
                    arrLine[229] = "<img src=\"https://drive.google.com/uc?export=view&amp;id=19VQpU5AOz7TcT-pXWzX_E1xDOEKix6xb\" alt=\"\" class=\"wppico\"/>";
                } else {
                    arrLine[226] = "<img src=\"https://api.qrserver.com/v1/create-qr-code/?data=https://www.supersonic.com.br/";
                    arrLine[227] = "";
                    arrLine[229] = "";
                }

                // Writing data into the HTML file
                File.WriteAllLines(fileName, arrLine);

                // Start Aspose convert HTML to PNG
                var document = new HTMLDocument("Assinatura.html");
                var options = new Aspose.Html.Saving.ImageSaveOptions(ImageFormat.Png);
                Converter.ConvertHTML(document, options, "output.png");
                // End Aspose convert HTML to PNG

                // Creating a Rectangle object and setting their dimensions
                Rectangle cropRect = new Rectangle(0, 80, 1594, 452);
                // Setting the image source that will be cropped
                using (Bitmap src = Image.FromFile("output.png") as Bitmap) {
                    using (Bitmap target = new Bitmap(cropRect.Width, cropRect.Height)) {
                        // Cropping the image
                        using (Graphics g = Graphics.FromImage(target)) {
                            g.DrawImage(src, new Rectangle(0, 0, target.Width, target.Height),
                                cropRect,
                                GraphicsUnit.Pixel);
                        }

                        // Reduce image size
                        Bitmap r = new Bitmap(510, 145);
                        Graphics gr = Graphics.FromImage((System.Drawing.Image)r);
                        gr.InterpolationMode = InterpolationMode.HighQualityBicubic;
                        gr.DrawImage(target, 0, 0, 510, 145);
                        gr.Dispose();

                        // Saving the Cropped image in two different resolutions
                        string saveGmail = "C:\\Users\\Matheus Moura\\Pictures\\ASSINATURAS\\" + arrLine[239] + " gmail.png";
                        string saveOutlook = "C:\\Users\\Matheus Moura\\Pictures\\ASSINATURAS\\" + arrLine[239] + " outlook.png";

                        r.Save(saveGmail);
                        target.Save(saveOutlook);
                    }
                }
                Console.WriteLine("Assinatura de " + nome[i] + " foi gerada com sucesso!");
            }

            Console.WriteLine("Todas as assinaturas foram geradas com sucesso!");
            Console.ReadKey();
        }
    }
}
