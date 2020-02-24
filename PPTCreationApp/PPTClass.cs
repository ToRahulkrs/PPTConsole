using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Syncfusion.Presentation;



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
        static void Main(string[] args)
        {
            PPTClass objPPTClass = new PPTClass("Rahul.pptx");
            objPPTClass.GeneratePPT();
            //string sText = "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with desktop publishing software like Aldus PageMaker including versions of Lorem Ipsum.";
            string sText = Console.ReadLine();
            objPPTClass.GeerateTablePPT();
            objPPTClass.WritePPT(sText);
            objPPTClass.ReadPPT();
         
            Console.ReadLine();
        }
    }
    
}
