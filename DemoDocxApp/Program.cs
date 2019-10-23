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
using Item = Cradle.Data.Item;
using Cradle.Data;
using Cradle.Lists;
using Cradle.Definitions;
using Cradle.ProjectSchema;
using Cradle.Server;
using List = Cradle.Lists.List;
using System;
using System.Runtime.InteropServices;
using DCSoft.RTF;

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
        static Stream BRtfDocx(byte[] byteArray)
        {
            
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
        private static RTFDomElement FindImage(RTFDomElementList list)
        {
            RTFDomElement outel = null;
            var elems = list.OfType<RTFDomImage>();
            foreach (RTFDomElement el in list)
            {
                var type = el.GetType();
                if (type.Equals(typeof(RTFDomImage)))
                { return el; }
                else
                {
                    outel = FindImage(el.Elements);
                    if (outel.GetType().Equals(typeof(RTFDomImage)))
                    {
                        return outel;
                    }
                }
            }
            return outel;
        }
        static Stream BRtfImage(byte[] byteArray)
        {

            MemoryStream msrtf = new MemoryStream(byteArray);
            MemoryStream msout = null;
            /*
            var document = new Spire.Doc.Document();
            //document.LoadFromFile(@"test-doc.rtf");
            ms.Seek(0, SeekOrigin.Begin);
            document.LoadRtf(ms, Encoding.UTF8);
            document.SaveToStream(msout, FileFormat.Docx2010);
            msout.Seek(0, SeekOrigin.Begin);
            //StreamReader reader = new StreamReader(msout);
            //string text = reader.ReadToEnd();
            */
            var rtf = new RTFDomDocument();
            rtf.Load(msrtf);
            var elems = rtf.Elements;
            var el = FindImage(elems);
            if (el != null)
            {
                var image = (RTFDomImage)el;
                image.DesiredWidth = 200;
                msout = new MemoryStream(image.Data); 
            }
            return msout;
        }
        static void OutItems(Table gtable, Item it, ref int row)
        {

            int myrow = row;
            //sl.InsertRow();
            //row++;
            Console.Write(".");

            {
                LinkedItem linked_item = null;
                List linked_items = null;
                Navigation nav = null;
                int num_linked_items;
                nav = new Navigation();

                nav.Type = "Включает";

                nav.Direction = CAPI_XREF_DIR.ONLY_DOWN;
                // Get linked items
                if (it.GetLinkedItems(CAPI_LINKS.ALL, nav, CAPI_INFO.NOTE, CAPI_SUBTYPE.NULL, "", null, default(int), default(QueryStereotypeTest), default(QueryModelviewTest), null, out linked_items))
                {
                    num_linked_items = linked_items.Length;
                    var loopTo = num_linked_items - 1;
                    for (int i = 0; i <= loopTo; i++)
                    {
                        if (linked_items.GetElement(i, out linked_item))
                        {
                            var l_it = linked_item.Item;
                            Frame frame = null;
                            string frame_text = null;
                            BinaryContents frame_tbl = null;
                            BinaryContents frame_pic = null;
                            DocX docin = null;
                            Table tins = null;
                            Picture pins = null;
                            Paragraph par = null;
                            l_it.Open(false);
                            if (!l_it.GetFrame("TEXT", out frame))
                                return ;

                            // Get TEXT frame contents
                            if (!frame.GetContents(out frame_text))
                                return ;
                            if (l_it.NoteType != "Титульный лист" )
                            {
                                if (!l_it.GetFrame("Таблица", out frame))
                                return;

                                // Get  frame contents
                                if (!frame.GetContents(out frame_tbl))
                                return;
                                if(frame_tbl.Size != 0)
                                {
                                    byte[] memdocx = new byte[frame_tbl.Size];
                                    IntPtr pnt = frame_tbl.Data;

                                    Marshal.Copy(pnt, memdocx, 0, memdocx.Length);
                                    docin = DocX.Load(BRtfDocx(memdocx));
                                    tins = docin.Tables[0];
                                }
                                if (!l_it.GetFrame("Рисунок", out frame))
                                    return;

                                // Get  frame contents
                                if (!frame.GetContents(out frame_pic))
                                    return;
                                if (frame_pic.Size != 0)
                                {
                                    byte[] memdocx = new byte[frame_pic.Size];
                                    IntPtr pnt = frame_pic.Data;

                                    Marshal.Copy(pnt, memdocx, 0, memdocx.Length);
                                    Stream stream = BRtfImage(memdocx);
                                    stream.Position = 0;
                                   
                                    Image img = doc.AddImage(stream);
                                    pic = img.CreatePicture();
                                    double  scale =(double)pic.Width / (double)pic.Height;
                                    pic.WidthInches = 3;
                                    pic.HeightInches = pic.WidthInches / scale;


                                }


                            }

                            l_it.Close();

                            gtable.InsertRow();
                            row++;
                            gtable.Rows[row].Cells[2].Paragraphs.First().Append(l_it.Identity);
                            gtable.Rows[row].Cells[2].InsertParagraph().Append(frame_text);
                            if (frame_tbl != null)
                            { if (frame_tbl.Size != 0)
                                {
                                    gtable.Rows[row].Cells[2].InsertParagraph().InsertTableAfterSelf(tins);
                                    gtable.Rows[row].Cells[2].InsertParagraph(docin.InsertParagraph(""));
                                }
                            }
                            if (frame_pic != null)
                            {
                                if (frame_pic.Size != 0)
                                {
                                    //Image img = doc.AddImage(bRtfDocx(memdocx));
                                    //Picture pic = img.CreatePicture();
                                    gtable.Rows[row].Cells[0].Paragraphs.First().InsertPicture(pic);
                                    //gtable.Rows[row].Cells[2].InsertParagraph(par);
                                    //gtable.Rows[row].Cells[2].InsertParagraph(doc.InsertParagraph(""));
                                }
                            }
                            //sl.InsertRow();


                        }

                    }
                    linked_items.Dispose();
                }

                nav.Type = "Содержит";

                nav.Direction = CAPI_XREF_DIR.ONLY_DOWN;
                // Get linked items
                if (it.GetLinkedItems(CAPI_LINKS.ALL, nav, CAPI_INFO.NOTE, CAPI_SUBTYPE.NULL, "", null, default(int), default(QueryStereotypeTest), default(QueryModelviewTest), null, out linked_items))
                {
                    num_linked_items = linked_items.Length;
                    var loopTo = num_linked_items - 1;
                    for (int i = 0; i <= loopTo; i++)
                    {
                        if (linked_items.GetElement(i, out linked_item))
                        {
                            var l_it = linked_item.Item;
                            gtable.InsertRow();
                            row++;
                            OutItems(gtable, l_it, ref row);

                        }

                    }
                    linked_items.Dispose();
                }

                gtable.Rows[myrow].Cells[2].Paragraphs.First().Append (it.Identity);
                gtable.Rows[myrow].Cells[2].InsertParagraph().Append(it.Name);
                //sl.InsertRow();

                return;
            }
   
        }
        static DocX doc;
        static Picture pic ;
        static void Main(string[] args)
        {
            //string path = "test-doc.rtf";
            //string readText = File.ReadAllText(path);
            //var docin = DocX.Load(RtfDocx(readText));
            //var tins = docin.Tables[0];

            string fileName = @"exempleWord.docx";
            doc = DocX.Create(fileName);

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

            //var r1 = new string[] {
            //    "Отчет должен строится на основе данных, получаемых из следующих систем и подсистем/программных модулей:\n1.ИСУП = MS Project\n2.IBM Лотус, используемые подсистемы:\na.ПМ Поручения\nb.ПМ Согласование ПГ\nc.ПМ Согласования ТЗ\nd.ПМ Договоры\n",
            //    "Изменено",
            //    "Отчет должен строится на основе данных, получаемых из следующих систем и подсистем/программных модулей:\n1.ИСУП = MS Project\n2.IBM Лотус, используемые подсистемы:\na.ПМ Поручения\nb.ПМ Согласование ПГ"
            //};
            //t.InsertRow();
            //FillRow(t.Rows[1], r1, System.Drawing.Color.Yellow);
            //t.InsertRow();
            //t.Rows[2].Cells[0].Paragraphs.First().AppendPicture(pic);
            //t.Rows[2].Cells[1].Paragraphs.First().Append("Удалено");
            //t.Rows[2].Cells[1].FillColor = System.Drawing.Color.Red;

            //t.InsertRow();

            ////FillRow(t.Rows[3], r3, System.Drawing.Color.LightGreen);
            //t.Rows[3].Cells[2].Paragraphs.First().Append("Отчет должен поддерживать фильтрацию (отображение) целевых задач по следующим условиям, с учётом пересчета кол-ва задач, описанного выше:\n1.По профильной системе\n2.По категории\n3.По дате / периоду\n4.По статусу задачи(завершенна, выполняется, планируется)")
            //    .InsertTableAfterSelf(tins);
            //t.Rows[3].Cells[2].InsertParagraph(doc.InsertParagraph(""));


            //t.Rows[3].Cells[1].Paragraphs.First().Append("Добавлено");
            //t.Rows[3].Cells[1].FillColor = System.Drawing.Color.LightGreen;
            Globals.Load_CradleAPI();
            Globals.GetArgs();
            //&TBL2&ПД&BL0&БЛ1&&Рзд ПД-37&ТЛ-1
            var bl1 = Globals.Args[3];
            var bl2 = Globals.Args[4];
            var phase = Globals.Args[2];
            var proj = new Project();
            var ldap = new LDAPInformation();
            if (!proj.Connect(Globals.CRADLE_CDS_HOST, Globals.CRADLE_PROJECT_CODE, Globals.CRADLE_USERNAME, Globals.CRADLE_PASSWORD, true,
                Cradle.Server.Connection.API_LICENCE, ldap, false))
            { return; };
            var project = proj;

            proj.SetBaselineMode(CAPI_BASELINE_MODE.SPECIFIED, bl2);
            Console.WriteLine("Начало обработки проекта " + proj.Title);

            int row = 1;

            Item root = new Item(CAPI_INFO.NOTE, "Раздел ТЗ " + phase);
            root.Identity = Globals.Args[6];
            root.Baseline = bl2;
            root.Version = "01";
            root.Draft = " ";
            if (root.Open(false))
            {
                root.Close();
                OutItems(t, root, ref row);
            }
            doc.InsertTable(t);
            //doc.InsertDocument(docin);
            doc.Save();
            Process.Start("WINWORD.EXE", fileName);

        }
    }
}
