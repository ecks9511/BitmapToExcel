
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Microsoft.Office.Core;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ImageToExcelCSharp
{
    class Program
    {
        static void Main()
        {

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            ExcelPackage ExcelPkg = new ExcelPackage();
            var pictureWorksheet = ExcelPkg.Workbook.Worksheets.Add("Picture");



            //Get file location from user
            Console.WriteLine("Please enter the location of your bitmap image:");
            String fileLocation = Console.ReadLine();

            //Open bitmap file
            Bitmap image = new Bitmap(fileLocation);

            //Rotate because im lazy and its printing wrong
            image.RotateFlip(RotateFlipType.Rotate90FlipNone);
            image.RotateFlip(RotateFlipType.RotateNoneFlipX);

            //Create a 2d array based on image height/width
            //Get all of the pixel colors in the image pixel by pixel, line by line
            Color[,] allImageColors = new Color[image.Width, image.Height];


            //Go down all rows
            for (int currentRow = 0; currentRow < image.Height; currentRow++)
            {
                //Going across all columns
                for (int currentCol = 0; currentCol < image.Width; currentCol++)
                {
                    allImageColors[currentCol, currentRow] = image.GetPixel(currentCol, currentRow);
                }
            }


            //Go down all rows
            for (int currentRow = 1; currentRow <= image.Height; currentRow++)
            {
                //Go accross all columns
                for (int currentCol = 1; currentCol <= image.Width; currentCol++)
                {
                    //Get current pixel color
                    Color currentColor = allImageColors[currentCol - 1, currentRow - 1];

                    //Set background color of current cell to pixel color
                    pictureWorksheet.Cells[currentCol, currentRow].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    pictureWorksheet.Cells[currentCol, currentRow].Style.Fill.BackgroundColor.SetColor(currentColor);

                }
            }

            for(int currentCol = 1; currentCol <= image.Width; currentCol++)
                pictureWorksheet.Column(currentCol).Width = 16;

            for (int currentRow = 1; currentRow <= image.Height; currentRow++)
                pictureWorksheet.Row(currentRow).Height = 32;

            //Save workbook
            pictureWorksheet.Protection.IsProtected = false;
            pictureWorksheet.Protection.AllowSelectLockedCells = false;
            ExcelPkg.SaveAs(new FileInfo(@"C:\Users\Jeremy\Desktop\picture.xlsx"));


        }
    }
}
