using System.Diagnostics;
//using System.Drawing;
using System.Linq;
using Xceed.Words.NET;
using Xceed.Document.NET;
using Font = Xceed.Document.NET.Font;
using System.IO;
using System.Text;
using Spire.Doc;
using Table = Xceed.Document.NET.Table;

namespace DemoDocxApp
{
    class Program
    {
        static void FillRow(Row row, string[] val,System.Drawing.Color color,bool header = false)
        {
            int i = 0;
            foreach (Cell cell in row.Cells)
            {
                if (header)
                {
                    cell.Paragraphs.First().Append(val[i]).Bold();
                    cell.FillColor = color;
                    i++;
                }
                else
                {
                    cell.Paragraphs.First().Append(val[i]);
                    if (i==1) cell.FillColor = color;
                    i++;
                }

            }
        }
        static Stream RtfDocx(string rtf)
        {
            byte[] byteArray = Encoding.ASCII.GetBytes(rtf);
            MemoryStream ms = new MemoryStream(byteArray);
            MemoryStream msout = new MemoryStream();
            StreamReader sReader = new StreamReader(ms);
            var document = new Spire.Doc.Document();
            //document.LoadFromFile(@"test-doc.rtf");
            ms.Seek(0, SeekOrigin.Begin);
            document.LoadRtf(ms, Encoding.UTF8);
            document.SaveToStream(msout, FileFormat.Docx);
            msout.Seek(0, SeekOrigin.Begin);
            //StreamReader reader = new StreamReader(msout);
            //string text = reader.ReadToEnd();
            return msout;


        }
        static void Main(string[] args)
        {
            string path = "test-doc.rtf";
            string readText = File.ReadAllText(path);
            var docin = DocX.Load(RtfDocx(readText));
            var tins = docin.Tables[0];

            string fileName = @"exempleWord.docx";
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


            //Create Table with 2 rows and 3 columns. 
            var header = new string[] { "Было", "Статус", "Стало" };
            Table t = doc.AddTable(2, header.Length);

            Image img = doc.AddImage(@"C:\Users\Владимир\source\repos\DemoDocxApp\Image.PNG");
            Picture pic = img.CreatePicture();
            
            
            //t.AutoFit = AutoFit.Fixed;
            t.SetWidthsPercentage(new float[] { 45F,10F, 45F },doc.PageWidth);
            t.Alignment = Alignment.center;
            t.Design = TableDesign.TableGrid;
            //Fill  by adding text.  
            //t.Rows[0].Cells[0].Paragraphs.First().Append("Было").Bold();
            //t.Rows[0].Cells[1].Paragraphs.First().Append("Статус").Bold();
            //t.Rows[0].Cells[2].Paragraphs.First().Append("Стало").Bold();
            //t.Rows[0].Cells[0].FillColor = System.Drawing.Color.LightGray;
            //t.Rows[0].Cells[1].FillColor = System.Drawing.Color.LightGray;
            //t.Rows[0].Cells[2].FillColor = System.Drawing.Color.LightGray;
            //foreach (Cell cell in t.Rows[0].Cells)
            //{ 
            //    cell.FillColor = System.Drawing.Color.LightGray;
            //}
            
            FillRow(t.Rows[0], header, System.Drawing.Color.LightGray,true);

            var r1 = new string[] {
                "Отчет должен строится на основе данных, получаемых из следующих систем и подсистем/программных модулей:\n1.ИСУП = MS Project\n2.IBM Лотус, используемые подсистемы:\na.ПМ Поручения\nb.ПМ Согласование ПГ\nc.ПМ Согласования ТЗ\nd.ПМ Договоры\n",
                "Изменено",
                "Отчет должен строится на основе данных, получаемых из следующих систем и подсистем/программных модулей:\n1.ИСУП = MS Project\n2.IBM Лотус, используемые подсистемы:\na.ПМ Поручения\nb.ПМ Согласование ПГ"
            };
            t.InsertRow();
            FillRow(t.Rows[1], r1, System.Drawing.Color.Yellow);
            t.InsertRow();
            t.Rows[2].Cells[0].Paragraphs.First().AppendPicture(pic);
            t.Rows[2].Cells[1].Paragraphs.First().Append("Удалено");
            t.Rows[2].Cells[1].FillColor = System.Drawing.Color.Red;
            
            t.InsertRow();

            //FillRow(t.Rows[3], r3, System.Drawing.Color.LightGreen);
            t.Rows[3].Cells[2].Paragraphs.First().Append("Отчет должен поддерживать фильтрацию (отображение) целевых задач по следующим условиям, с учётом пересчета кол-ва задач, описанного выше:\n1.По профильной системе\n2.По категории\n3.По дате / периоду\n4.По статусу задачи(завершенна, выполняется, планируется)")
                .InsertTableAfterSelf(tins);
            t.Rows[3].Cells[2].InsertParagraph(doc.InsertParagraph(""));
            

            t.Rows[3].Cells[1].Paragraphs.First().Append("Добавлено");
            t.Rows[3].Cells[1].FillColor = System.Drawing.Color.LightGreen;
            doc.InsertTable(t);
            //doc.InsertDocument(docin);
            doc.Save();
            Process.Start("WINWORD.EXE", fileName);

        }
    }
}
