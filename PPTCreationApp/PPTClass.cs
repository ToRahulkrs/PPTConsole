using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Syncfusion.Presentation;
using System.Drawing;



namespace PPTCreationApp
{
    class PPTClass
    {
        string fileName = null;
        PPTClass(string _fileName)
        {
            try
            {
                fileName = _fileName;
            }
            catch (Exception)
            {

                throw;
            }
        }
        public string GeneratePath()
        {
            try
            {
                return Path.Combine(Directory.GetParent(System.IO.Directory.GetCurrentDirectory()).Parent.Parent.Parent.FullName, fileName);
            }
            catch (Exception)
            {

                throw;
            }
        }
        public void GeneratePPT()
        {
            try
            {
                string sPath = GeneratePath();
                if (File.Exists(sPath))
                {
                    File.Delete(sPath);
                }
                using (var powerpointDoc = Presentation.Create())
                {
                    ISlide slide = powerpointDoc.Slides.Add(SlideLayoutType.Blank);
                    powerpointDoc.Save(sPath);
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        public void GeerateTablePPT()
        {
            try
            {
                using (var powerpointDoc = Presentation.Create())
                {
                    ISlide slide = powerpointDoc.Slides.Add(SlideLayoutType.Blank);
                    ITable table = slide.Shapes.AddTable(2, 2, 100, 120, 300, 200);
                    int rowIndex = 0, colIndex;
                    foreach (IRow rows in table.Rows)
                    {
                        colIndex = 0;

                        foreach (ICell cell in rows.Cells)
                        {

                            cell.TextBody.AddParagraph("(" + rowIndex.ToString() + " , " + colIndex.ToString() + ")");

                            colIndex++;

                        }

                        rowIndex++;

                    }
                    powerpointDoc.Save(GeneratePath());
                }
        
         
            }
            catch (Exception)
            {
                throw;
            }
        }
        public void ReadPPT()
        {
            try
            {
                using (var powerpointDoc = Presentation.Open(GeneratePath()))
                {
                    foreach (ISlide slide in powerpointDoc.Slides)
                    {
                        foreach (IShape shape in slide.Shapes)
                        {
                            if (shape.SlideItemType == SlideItemType.AutoShape)
                                Console.WriteLine(shape.TextBody.Text + "\n");
                        }
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        public void WritePPT(string sText)
        {
            try
            {
                using (var powerpointDoc = Presentation.Open(GeneratePath()))
                {
                    ISlide slide = powerpointDoc.Slides[0];
                    IShape shape = slide.AddTextBox(10, 10, 500, 100);
                    shape.TextBody.AddParagraph(sText);
                    powerpointDoc.Save(GeneratePath());
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        public void GenratepptxUsingExistingPpt(string sText)
        {
            try
            {
                IPresentation pptxDoc = Presentation.Open("C:\\Project-Result\\PPTx-Project\\PPTCreationApp\\DataFile\\DemoPPT.pptx");
                ILayoutSlide layoutSlide = pptxDoc.Masters[1].LayoutSlides.Add(SlideLayoutType.Blank, "CustomLayout");
                IShape shape = layoutSlide.Shapes.AddShape(AutoShapeType.Diamond, 30, 20, 400, 300);
                layoutSlide.Background.Fill.SolidFill.Color = ColorObject.FromArgb(78, 89, 90);
                ISlide slide = pptxDoc.Slides.Add(layoutSlide);
                ITable table = slide.Shapes.AddTable(1, 3, 100, 120, 600, 50);

                Syncfusion.Drawing.Image image = Syncfusion.Drawing.Image.FromFile("C:\\Project-Result\\PPTx-Project\\PPTCreationApp\\DataFile\\Ninja.jpg");

                ICell cell1 = table.Rows[0].Cells[0] as ICell;
                cell1.Fill.FillType = FillType.Picture;
                cell1.Fill.PictureFill.ImageBytes = image.ImageData;

                ICell cell2 = table.Rows[0].Cells[1] as ICell;
                cell2.Fill.FillType = FillType.Solid;
                cell2.Fill.SolidFill.Color.SystemColor = Color.BlueViolet;
                cell2.TextBody.Text = sText.ToString();

                ICell cell3 = table.Rows[0].Cells[2] as ICell;
                Stream txtStream = File.Open("C:\\Project-Result\\PPTx-Project\\PPTCreationApp\\DataFile\\Ornament-CodeFlow.pdf", FileMode.Open);
                //Stream imageStream = File.Open("D:\\PPTxExamplesConsoleApps\\PPTConsole\\PPTCreationApp\\Images\\TextFile.png", FileMode.Open);
                byte[] file = File.ReadAllBytes("C:\\Project-Result\\PPTx-Project\\PPTCreationApp\\DataFile\\logopdf.png");
                Stream imageStream = new MemoryStream(file);
                cell3.Fill.FillType = FillType.Picture;
                cell3.Fill.PictureFill.ImageBytes = file;



                IOleObject oleObject = slide.Shapes.AddOleObject(imageStream, "Excel.Sheet.12", txtStream);


                //Set size and position of the OLE object
                oleObject.Left = 80;
                oleObject.Top = 30;
                oleObject.Width = 40;
                oleObject.Height = 30;



                int rowIndex = 0, colIndex;

                //Iterate row-wise cells and add text to it
                foreach (IRow rows in table.Rows)
                {
                    colIndex = 0;

                    foreach (ICell cell in rows.Cells)
                    {

                        //cell.TextBody.AddParagraph("(" + rowIndex.ToString() + " , " + colIndex.ToString() + ")");

                        //colIndex++;

                    }

                    rowIndex++;

                }
                pptxDoc.Save("C:\\Project-Result\\PPTx-Project\\PPTCreationApp\\DataFile\\DemoPPT.pptx");
                pptxDoc.Close();
            }
            catch (Exception ex)
            {

                throw;
            }

        }
        static void Main(string[] args)
        {
            PPTClass objPPTClass = new PPTClass("Rahul.pptx");
            //objPPTClass.GeneratePPT();
            string sText = "Lorem Ipsum is simply dummy text of the printing and typesetting industry";
            //string sText = Console.ReadLine();
            //objPPTClass.GeerateTablePPT();
            //objPPTClass.WritePPT(sText);
            //objPPTClass.ReadPPT();
            objPPTClass.GenratepptxUsingExistingPpt(sText);

            Console.ReadLine();
        }
    }
    
}
