using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace GeneradorBingo
{
    public static class Exportar
    {
        public static void CrearArchivo(string[][] carton, int n)
        {

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Crear un nuevo archivo Excel
            using (var package = new ExcelPackage(new FileInfo("C:\\Users\\elrey\\OneDrive\\Escritorio\\cartonesBingo.xlsx")))
            {
                var workbook = package.Workbook;
                // Obtener una hoja
                var sourceWorksheet = workbook.Worksheets["CARTON BASE"];
                var newWorksheet = workbook.Worksheets.Add($"CARTON{n:D3}");

                // Copia el contenido y formato de la hoja completa
                foreach (var cell in sourceWorksheet.Cells)
                {
                    var targetCell = newWorksheet.Cells[cell.Start.Row, cell.Start.Column];
                    targetCell.Value = cell.Text;

                    // Copia el estilo de la celda (incluyendo colores)
                    targetCell.Style.Numberformat.Format = cell.Style.Numberformat.Format;
                    targetCell.Style.Font.Bold = cell.Style.Font.Bold;
                    targetCell.Style.Font.Italic = cell.Style.Font.Italic;
                    targetCell.Style.Font.Size = cell.Style.Font.Size;
                    targetCell.Style.HorizontalAlignment = cell.Style.HorizontalAlignment;
                    targetCell.Style.VerticalAlignment = cell.Style.VerticalAlignment;

                    // Copia los bordes de la celda (si existen)
                    if (cell.Style.Border.Left.Style != ExcelBorderStyle.None)
                    {
                        targetCell.Style.Border.Left.Style = cell.Style.Border.Left.Style;
                        targetCell.Style.Border.Left.Color.SetColor(ColorTranslator.FromHtml(cell.Style.Border.Left.Color.Rgb));
                    }
                    if (cell.Style.Border.Right.Style != ExcelBorderStyle.None)
                    {
                        targetCell.Style.Border.Right.Style = cell.Style.Border.Right.Style;
                        targetCell.Style.Border.Right.Color.SetColor(ColorTranslator.FromHtml(cell.Style.Border.Right.Color.Rgb));
                    }
                    if (cell.Style.Border.Top.Style != ExcelBorderStyle.None)
                    {
                        targetCell.Style.Border.Top.Style = cell.Style.Border.Top.Style;
                        targetCell.Style.Border.Top.Color.SetColor(ColorTranslator.FromHtml(cell.Style.Border.Top.Color.Rgb));
                    }
                    if (cell.Style.Border.Bottom.Style != ExcelBorderStyle.None)
                    {
                        targetCell.Style.Border.Bottom.Style = cell.Style.Border.Bottom.Style;
                        targetCell.Style.Border.Bottom.Color.SetColor(ColorTranslator.FromHtml(cell.Style.Border.Bottom.Color.Rgb));
                    }

                    targetCell.AutoFitColumns();
                }

                var texto1 = new List<string>() { "M", "N", "O", "P", "Q" };
                var num1 = new List<int>() { 3, 4, 5, 6, 7 };
                var texto2 = new List<string>() { "U", "V", "W", "X", "Y" };
                var num2 = new List<int>() {  11, 12, 13 ,14, 15};
                var texto3 = new List<string>() { "Q", "R", "S", "T", "U" };
                var num3 = new List<int> { 19, 20, 21, 22,23 };


                for (int i = 0; i < 5; i++)
                {
                    for (int j = 0; j < 5; j++)
                    {
                        if(carton[i][j] != "libre")
                        {
                            var te = $"{carton[i][j]:D2}".ToString();

                            newWorksheet.Cells[$"{texto1[j]}{num1[i]}"].Value = te;
                            newWorksheet.Cells[$"{texto1[j]}{num1[i]}"].Style.Numberformat.Format = "@";
                            newWorksheet.Cells[$"{texto1[j]}{num3[i]}"].Value = te; ;
                            newWorksheet.Cells[$"{texto1[j]}{num3[i]}"].Style.Numberformat.Format = "@";
                            newWorksheet.Cells[$"{texto3[j]}{num2[i]}"].Value = te; ;
                            newWorksheet.Cells[$"{texto3[j]}{num2[i]}"].Style.Numberformat.Format = "@";
                            newWorksheet.Cells[$"{texto2[j]}{num1[i]}"].Value = te; ;
                            newWorksheet.Cells[$"{texto2[j]}{num1[i]}"].Style.Numberformat.Format = "@";
                            newWorksheet.Cells[$"{texto2[j]}{num3[i]}"].Value = te; ;
                            newWorksheet.Cells[$"{texto2[j]}{num3[i]}"].Style.Numberformat.Format = "@";
                        }

                    }

                }
                newWorksheet.Cells[$"Q8"].LoadFromText($"Cartón{n:D3}");
                newWorksheet.Cells[$"Q8"].AutoFitColumns();

                newWorksheet.Cells[$"U16"].LoadFromText($"Cartón{n:D3}");
                newWorksheet.Cells[$"U16"].AutoFitColumns();

                newWorksheet.Cells[$"Y8"].LoadFromText($"Cartón{n:D3}");
                newWorksheet.Cells[$"Y8"].AutoFitColumns();

                newWorksheet.Cells[$"Q24"].LoadFromText($"Cartón{n:D3}");
                newWorksheet.Cells[$"Q24"].AutoFitColumns();

                newWorksheet.Cells[$"Y24"].LoadFromText($"Cartón{n:D3}");
                newWorksheet.Cells[$"Y24"].AutoFitColumns();

                // Guarda el archivo Excel con la hoja copiada
                package.Save();
            }

        }


        

    }
}
