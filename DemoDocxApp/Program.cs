using System.Diagnostics;
//using System.Drawing;
using System.Linq;
using Xceed.Words.NET;
using Xceed.Document.NET;
using Font = Xceed.Document.NET.Font;

namespace DemoDocxApp
{
    class Program
    {
        static void Main(string[] args)
        {
            string fileName = @"C:\temp\exempleWord.docx";
            var doc = DocX.Create(fileName);

            doc.PageLayout.Orientation = Orientation.Landscape;

            //Formatting Text Paragraph  
            Formatting textParagraphFormat = new Formatting
            {
                //font family  
                FontFamily = new Font("Times New Roman"),
                //font size  
                Size = 10D,
                //Spaces between characters  
                //Spacing = 1
            };

            /*
            Table t = document.AddTable(rows,columns);
            t.AutoFit = AutoFit.ColumnWidth;
            for (int x = 0; x < t.ColumnCount; x++)
            {
            //t.SetColumnWidth(x, columnsizes[x]);
            for (int y = 0; y < t.RowCount; y++)
            {
            t.Rows[y].Cells[x].Width = columnsizes[x];
            }
            }

              */


            //Create Table with 2 rows and 4 columns. 

            Image img = doc.AddImage(@"C:\Users\Владимир\source\repos\DemoDocxApp\Image.PNG");
            Picture pic = img.CreatePicture();
            Table t = doc.AddTable(2, 3);
            Table t2 = doc.AddTable(2, 3);
            //t.AutoFit = AutoFit.Fixed;
            t.SetWidthsPercentage(new float[] { 45F,10F, 45F },doc.PageWidth);
            t.Alignment = Alignment.center;
            t.Design = TableDesign.TableGrid;
            //Fill  by adding text.  
            t.Rows[0].Cells[0].Paragraphs.First().Append("Было").Bold();
            t.Rows[0].Cells[1].Paragraphs.First().Append("Статус").Bold();
            t.Rows[0].Cells[2].Paragraphs.First().Append("Стало").Bold();
            t.Rows[0].Cells[0].FillColor = System.Drawing.Color.LightGray;
            t.Rows[0].Cells[1].FillColor = System.Drawing.Color.LightGray;
            t.Rows[0].Cells[2].FillColor = System.Drawing.Color.LightGray;

            t.Rows[1].Cells[0].Paragraphs.First().Append("Отчет должен строится на основе данных, получаемых из следующих систем и подсистем/программных модулей:\n1.ИСУП = MS Project\n2.IBM Лотус, используемые подсистемы:\na.ПМ Поручения\nb.ПМ Согласование ПГ\nc.ПМ Согласования ТЗ\nd.ПМ Договоры\n"
                , textParagraphFormat).AppendPicture(pic);
            t.Rows[1].Cells[1].Paragraphs.First().Append("Изменено");
            t.Rows[1].Cells[1].FillColor = System.Drawing.Color.Yellow;
            t.Rows[1].Cells[2].Paragraphs.First().Append("Отчет должен строится на основе данных, получаемых из следующих систем и подсистем/программных модулей:\n1.ИСУП = MS Project\n2.IBM Лотус, используемые подсистемы:\na.ПМ Поручения\nb.ПМ Согласование ПГ"
                , textParagraphFormat).AppendPicture(pic);
            t.InsertRow();
            t.Rows[2].Cells[0].Paragraphs.First().AppendPicture(pic);
            t.Rows[2].Cells[1].Paragraphs.First().Append("Удалено");
            t.Rows[2].Cells[1].FillColor = System.Drawing.Color.Red;
            
            t.InsertRow();
            t.Rows[3].Cells[2].Paragraphs.First().Append("Отчет должен поддерживать фильтрацию (отображение) целевых задач по следующим условиям, с учётом пересчета кол-ва задач, описанного выше:\n1.По профильной системе\n2.По категории\n3.По дате / периоду\n4.По статусу задачи(завершенна, выполняется, планируется)")
                .InsertTableAfterSelf(t);
            t.Rows[3].Cells[2].InsertParagraph(doc.InsertParagraph(""));

            t.Rows[3].Cells[1].Paragraphs.First().Append("Добавлено");
            t.Rows[3].Cells[1].FillColor = System.Drawing.Color.LightGreen;
            doc.InsertTable(t);
            doc.Save();
            Process.Start("WINWORD.EXE", fileName);

        }
    }
}
