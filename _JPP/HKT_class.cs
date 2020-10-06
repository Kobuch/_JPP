using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

namespace _JPP
{

    public class HKT_class
    {
        public void KHT_schemat()
        {

            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Tabelka tabelka = new Tabelka();
            List<Texty> napisycad = new List<Texty>();
            


            // Start a transaction
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Get the current document editor
                Editor acDocEd = Application.DocumentManager.MdiActiveDocument.Editor;
              //  Application.ShowAlertDialog("Wskaż rogi tabekli wg opisu \n 1 - gornylewy róg wartości " +
               //                              "\n 2 - dolnyprawy róg wartości");

                tabelka.punkt1 = acDocEd.GetPoint("\n Wskaż punkt 1 - gornylewy róg wartości").Value;
                tabelka.punkt2 = acDocEd.GetPoint("\n Wskaż punkt 2 - dolnyprawy róg wartości").Value;
                string[,] listaatrybutow = new string[2, 30];

                //wybor oknem automatycznym

                PromptSelectionResult acSSPrompt;

                acSSPrompt = acDocEd.SelectCrossingWindow(tabelka.punkt1, tabelka.punkt2);
                if (acSSPrompt.Status == PromptStatus.OK)
                {
                    SelectionSet acSSet = acSSPrompt.Value;

                    // Step through the objects in the selection set
                    foreach (SelectedObject acSSObj in acSSet)
                    {
                        // Check to make sure a valid SelectedObject object was returned
                        if (acSSObj != null)
                        {
                            // Open the selected object for write
                            Entity acEnt = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead) as Entity;

                            if (acEnt != null)

                            {
                                if (acEnt.GetType().Name == "Line") tabelka.dodajjezelilinia(acEnt as Line);

                                if (acEnt.GetType().Name == "Polyline") tabelka.dodajjezelipolilinia(acEnt as Autodesk.AutoCAD.DatabaseServices.Polyline);


                                if (acEnt.GetType().Name == "DBText") tabelka.dodajtextdolisty(acEnt as DBText);

                                if (acEnt.GetType().Name == "MText") tabelka.dodajMtextdolisty(acEnt as MText);

                                if (acEnt.GetType().Name == "BlockReference")
                                {
                                    tabelka.dodajliniejezeliblok(acEnt as BlockReference);

                                    AttributeCollection attCol = ((BlockReference)acEnt).AttributeCollection;
                                    foreach (ObjectId attId in attCol)
                                    {
                                        AttributeReference attRef = (AttributeReference)acTrans.GetObject(attId, OpenMode.ForRead);
                                        tabelka.dodajBlockreferencedolisty(attRef);
                                    }

                                    BlockReference blockref = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead) as BlockReference;

                                    BlockTableRecord blockTablrec = acTrans.GetObject(blockref.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;

                                    foreach (ObjectId asObjId in blockTablrec)
                                    {
                                        Entity e = (Entity)acTrans.GetObject(asObjId, OpenMode.ForRead);

                                        if (e.GetType().Name == "Line") tabelka.dodajjezelilinia_wbloku(e as Line, blockref.Position);
                                        if (e.GetType().Name == "Polyline") tabelka.dodajjezelipolilinia_wbloku(e as Autodesk.AutoCAD.DatabaseServices.Polyline, blockref.Position);

                                    }
                                }
                            }
                        }
                    }
                    acTrans.Commit();

                    //porzadkowanie wartosci
                    tabelka.porzadkujpionowelinie();
                    tabelka.porzadkujpoziome();
                    tabelka.dodajkolumnyiwierszedonapisow();
                    //  tabelka.aktualizacjakolumn();

                    Application.ShowAlertDialog("Number of objects selected: " +
                                                 acSSet.Count.ToString() + "\n Kolumn: " + tabelka.ilekolumn.ToString() +
                                                  "\n Wierszy: " + tabelka.ilewierszy.ToString());


                }
            }
            //napisycad1 = tabelka.textycad;
            //tabelka1 = tabelka;
          



            ////TODO tutaj dorobić prodecyrę wszystkich i wywołanie różnych schematów

        }







        public void rysuj_schemat_rifu_80(Point3d X0Y0, List<string> teksty)
        {

            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {

                Editor acDocEd = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;

                ObjectIdCollection acObjIdColl = new ObjectIdCollection();

                //zasil
                // Create a lightweight polyline
                Autodesk.AutoCAD.DatabaseServices.Polyline acPoly = new Autodesk.AutoCAD.DatabaseServices.Polyline();
                acPoly.SetDatabaseDefaults();
                acPoly.AddVertexAt(0, new Point2d(X0Y0.X+0 , X0Y0.Y+0), 0, 0, 0);
                acPoly.AddVertexAt(1, new Point2d(X0Y0.X + 0, X0Y0.Y + 17800), 0, 0, 0);
                acPoly.AddVertexAt(2, new Point2d(X0Y0.X - 300, X0Y0.Y + 17800), 0, 0, 0);


                //odu
                Autodesk.AutoCAD.DatabaseServices.Polyline acPoly1 = new Autodesk.AutoCAD.DatabaseServices.Polyline();
                acPoly1.SetDatabaseDefaults();
                acPoly1.AddVertexAt(0, new Point2d(X0Y0.X -300, X0Y0.Y + 17700), 0, 0, 0);
                acPoly1.AddVertexAt(1, new Point2d(X0Y0.X - 300, X0Y0.Y + 17700 + 560), 0, 0, 0);
                acPoly1.AddVertexAt(2, new Point2d(X0Y0.X - 300-960, X0Y0.Y + 17700), 0, 0, 0);
                acPoly1.AddVertexAt(2, new Point2d(X0Y0.X - 300 - 960, X0Y0.Y + 17700 + 560), 0, 0, 0);
                acPoly1.Closed = true;


                Ellipse acEllipse = new Ellipse(new Point3d(X0Y0.X - 300 - 960, X0Y0.Y + 17700 + 280, 0), 40* Vector3d.ZAxis, 160 * Vector3d.YAxis, 0.25, 0, 360 * Math.Atan(1.0) / 45.0);

                                  

                //antena
                Autodesk.AutoCAD.DatabaseServices.Arc acArc = new Autodesk.AutoCAD.DatabaseServices.Arc(
                    new Point3d(X0Y0.X - 300 - 960 - 1050, X0Y0.Y + 17800 + 180, 0), 1010, 3.1415*1.5, 3.1415/2);


                Autodesk.AutoCAD.DatabaseServices.Line acline = new Line(acArc.StartPoint, acArc.EndPoint);

                //kabel eth

                Autodesk.AutoCAD.DatabaseServices.Polyline acPoly2 = new Autodesk.AutoCAD.DatabaseServices.Polyline();
                acPoly2.SetDatabaseDefaults();
                acPoly2.AddVertexAt(0, new Point2d(X0Y0.X + 850, X0Y0.Y + 0), 0, 0, 0);
                acPoly2.AddVertexAt(1, new Point2d(X0Y0.X + 850, X0Y0.Y + 17800+250), 0, 0, 0);
                acPoly2.AddVertexAt(2, new Point2d(X0Y0.X - 300, X0Y0.Y + 17800+250), 0, 0, 0);
               
                //uziemienie
                Autodesk.AutoCAD.DatabaseServices.Line acline2 = new Line(new Point3d(X0Y0.X - 400, X0Y0.Y + 17700,0), new Point3d(X0Y0.X - 400, X0Y0.Y + 17700-230, 0));
                Autodesk.AutoCAD.DatabaseServices.Line acline3 = new Line(new Point3d(X0Y0.X - 400-77, X0Y0.Y + 17700-230, 0), new Point3d(X0Y0.X - 400+77, X0Y0.Y + 17700 - 230, 0));
                Autodesk.AutoCAD.DatabaseServices.Line acline4 = new Line(new Point3d(X0Y0.X - 400 - 56, X0Y0.Y + 17700 - 230-40, 0), new Point3d(X0Y0.X - 400 + 56, X0Y0.Y + 17700 - 230-40, 0));
                Autodesk.AutoCAD.DatabaseServices.Line acline5 = new Line(new Point3d(X0Y0.X - 400 - 20, X0Y0.Y + 17700 - 230-80, 0), new Point3d(X0Y0.X - 400 + 20, X0Y0.Y + 17700 - 230-80, 0));



                MText acMText = new MText();
                acMText.SetDatabaseDefaults();
                acMText.Rotation = Math.PI / 2;
                acMText.SetAttachmentMovingLocation(AttachmentPoint.MiddleCenter);
                acMText.Location = new Point3d(X0Y0.X , X0Y0.Y + 8800,0);
                acMText.ColorIndex = 7;
                acMText.Contents = teksty[0];
                acMText.TextHeight = 250;



                    // Add the new object to the block table record and the transaction

                    Autodesk.AutoCAD.DatabaseServices.Wipeout wipeout = new Autodesk.AutoCAD.DatabaseServices.Wipeout();

                var space = (BlockTableRecord)acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite);
                space.AppendEntity(acPoly);
                space.AppendEntity(acPoly1);
                space.AppendEntity(acArc);
                space.AppendEntity(acline);
                space.AppendEntity(acPoly2);

                space.AppendEntity(acline2);
                space.AppendEntity(acline3);
                space.AppendEntity(acline4);
                space.AppendEntity(acline5);

                space.AppendEntity(acEllipse);
                space.AppendEntity(acMText);




                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.AddNewlyCreatedDBObject(acPoly1, true);
                acTrans.AddNewlyCreatedDBObject(acArc, true);
                acTrans.AddNewlyCreatedDBObject(acline, true);
                acTrans.AddNewlyCreatedDBObject(acPoly2, true);

                acTrans.AddNewlyCreatedDBObject(acline2, true);
                acTrans.AddNewlyCreatedDBObject(acline3, true);
                acTrans.AddNewlyCreatedDBObject(acline4, true);
                acTrans.AddNewlyCreatedDBObject(acline5, true);

                acTrans.AddNewlyCreatedDBObject(acEllipse, true);

                acTrans.AddNewlyCreatedDBObject(acMText, true);







                acTrans.Commit();

                Hatch_object(acEllipse.ObjectId);


            }

        


         }

        public static void HatchPolyLine(ObjectId plineId)
        {
            try
            {
                if (plineId.IsNull)
                    throw new ArgumentNullException("plineId");

                if (plineId.ObjectClass != RXObject.GetClass(typeof(Polyline)))
                    throw new Autodesk.AutoCAD.Runtime.Exception(ErrorStatus.IllegalEntityType);

                var ids = new ObjectIdCollection();
                ids.Add(plineId);

                using (var tr = plineId.Database.TransactionManager.StartTransaction())
                {
                    var pline = (Polyline)tr.GetObject(plineId, OpenMode.ForRead);
                    if (!(pline.Closed || pline.GetPoint2dAt(0).IsEqualTo(pline.GetPoint2dAt(pline.NumberOfVertices - 1))))
                        throw new InvalidOperationException("Opened polyline.");

                    var owner = (BlockTableRecord)tr.GetObject(pline.OwnerId, OpenMode.ForWrite);
                    var hatch = new Hatch();
                    hatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
                    owner.AppendEntity(hatch);
                    tr.AddNewlyCreatedDBObject(hatch, true);
                    hatch.Associative = true;
                    hatch.AppendLoop(HatchLoopTypes.Default, ids);
                    hatch.EvaluateHatch(true);
                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                var ed = Application.DocumentManager.MdiActiveDocument.Editor;
                ed.WriteMessage($"{ex.Message}\n{ex.StackTrace}");
            }
        }

        public static void Hatch_object(ObjectId objId)
        {
            try
            {
                if (objId.IsNull)
                    throw new ArgumentNullException("objId");

                
                var ids = new ObjectIdCollection();
                ids.Add(objId);

                using (var tr = objId.Database.TransactionManager.StartTransaction())
                {

                    var pline = (Ellipse)tr.GetObject(objId, OpenMode.ForRead);

                    var owner = (BlockTableRecord)tr.GetObject(pline.OwnerId, OpenMode.ForWrite);
                    var hatch = new Hatch();
                    hatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
                    owner.AppendEntity(hatch);
                    tr.AddNewlyCreatedDBObject(hatch, true);
                    hatch.Associative = true;
                    hatch.AppendLoop(HatchLoopTypes.Default, ids);
                    hatch.EvaluateHatch(true);
                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                var ed = Application.DocumentManager.MdiActiveDocument.Editor;
                ed.WriteMessage($"{ex.Message}\n{ex.StackTrace}");
            }
        }







        public class Tabelka
        {
            public Point3d punkt1 { get; set; }
            public Point3d punkt2 { get; set; }
            public Point3d punkt3 { get; set; }
            public Point3d punkt4 { get; set; }
            public Point3d punkt5 { get; set; }


            public int[] nrkolumny = new int[30];

            public kolumna[] zesatwkolumn = new kolumna[30];
            public kolumna[] zesatwwierszy = new kolumna[20];

            private List<int> liniepionowe = new List<int>();
            private List<int> liniepoziome = new List<int>();

            public List<Texty> textycad = new List<Texty>();
            public List<Textydocad> textydocad = new List<Textydocad>();

            public int ilekolumn = 0;
            public int ilewierszy = 0;

            public string[,] napisy_z_excel = new string[20, 30];

            public Tabelka()
            {
                punkt1 = new Point3d();
                punkt2 = new Point3d();
                punkt3 = new Point3d();
                punkt4 = new Point3d();
                punkt5 = new Point3d();
                for (int k = 0; k < 30; k++)
                {
                    zesatwkolumn[k] = new kolumna();
                }
                for (int k = 0; k < 20; k++)
                {
                    zesatwwierszy[k] = new kolumna();
                }

            }


            public void porzadkujpoziome()
            //sortuje liste watosci Y na liście  linii poziomych i jeżeli sie nie powtarzaja to po kolei kworzy woersze od linii poprezdniej do bierzącej.
            //tolerancja to dopuczalny nimimalna odległośc międzyliniami pionowymi mniej niz toleracja sa traktowane jako jedna linia 
            {
                int toleracja = 5;
                liniepoziome.Sort();
                liniepoziome.Reverse();
                int k = 1;



                liniepoziome.RemoveAll(item => item > Convert.ToInt32(punkt2.Y));
                liniepoziome.RemoveAll(item => item < Convert.ToInt32(punkt1.Y));


                if (liniepoziome.Count < 1) return;
                zesatwwierszy[0].Y0 = liniepoziome[0];
                zesatwwierszy[0].Y1 = liniepoziome[0];
                foreach (int pojedyncza in liniepoziome)
                {
                    if ((k < 20) && (pojedyncza < zesatwwierszy[k - 1].Y1 - toleracja))
                    {

                        zesatwwierszy[k].Y0 = zesatwwierszy[k - 1].Y1;
                        zesatwwierszy[k].Y1 = pojedyncza;
                        k = k + 1;
                    }
                    ilewierszy = k - 1;
                }
            }


            public void dodajjezelipolilinia(Polyline plina)
            {

                double deltaX = Math.Abs(plina.StartPoint.X - plina.EndPoint.X);
                double deltaY = Math.Abs(plina.StartPoint.Y - plina.EndPoint.Y);

                //sprawdz czy linia jest pionowa, porównanie delt
                if (deltaX < deltaY)
                {
                    //jezeli pionowa dodaje do pionowych
                    liniepionowe.Add(Convert.ToInt32(plina.StartPoint.X));
                }
                else
                {
                    // jezeli pozioma dodaje do poziomych
                    liniepoziome.Add(Convert.ToInt32(plina.StartPoint.Y));
                }
            }

            public void dodajjezelipolilinia_wbloku(Polyline plina, Point3d punkt)
            {
                liniepionowe.Add(Convert.ToInt32(plina.Bounds.Value.MinPoint.X + punkt.X));
                liniepionowe.Add(Convert.ToInt32(plina.Bounds.Value.MaxPoint.X + punkt.X));

                liniepoziome.Add(Convert.ToInt32(plina.Bounds.Value.MinPoint.Y + punkt.Y));
                liniepoziome.Add(Convert.ToInt32(plina.Bounds.Value.MaxPoint.Y + punkt.Y));

            }



            public void dodajjezelilinia(Line line)
            {
                double deltaX = Math.Abs(line.StartPoint.X - line.EndPoint.X);
                double deltaY = Math.Abs(line.StartPoint.Y - line.EndPoint.Y);

                //sprawdz czy linia jest pionowa, porównanie delt
                if (deltaX < deltaY)
                {
                    //jezeli pionowa dodaje do pionowych
                    liniepionowe.Add(Convert.ToInt32(line.StartPoint.X));
                }
                else
                {
                    // jezeli pozioma dodaje do poziomych
                    liniepoziome.Add(Convert.ToInt32(line.StartPoint.Y));
                }
            }
            public void dodajjezelilinia_wbloku(Line line, Point3d punkt)
            {
                double deltaX = Math.Abs(line.StartPoint.X - line.EndPoint.X);
                double deltaY = Math.Abs(line.StartPoint.Y - line.EndPoint.Y);

                //sprawdz czy linia jest pionowa, porównanie delt
                if (deltaX < deltaY)
                {
                    //jezeli pionowa dodaje do pionowych
                    liniepionowe.Add(Convert.ToInt32(line.StartPoint.X + punkt.X));
                }
                else
                {
                    // jezeli pozioma dodaje do poziomych
                    liniepoziome.Add(Convert.ToInt32(line.StartPoint.Y + punkt.Y));
                }
            }





            public void dodajliniejezeliblok(BlockReference block)
            {
                liniepionowe.Add(Convert.ToInt32(block.Bounds.Value.MinPoint.X));
                liniepionowe.Add(Convert.ToInt32(block.Bounds.Value.MaxPoint.X));

                liniepoziome.Add(Convert.ToInt32(block.Bounds.Value.MinPoint.Y));
                liniepoziome.Add(Convert.ToInt32(block.Bounds.Value.MaxPoint.Y));
            }

            public void porzadkujpionowelinie()
            //sortuje liste watosci X na liście i jeżeli sie nie powtarzaja to po kolei kworzy kolumny od linii poprezdniej do bierzącej.
            //tolerancja to dopuczalny nimimalna odległośc międzyliniami pionowymi mniej niz toleracja sa traktowane jako jedna linia 
            {
                int toleracja = 5;
                liniepionowe.Sort();
                int k = 1;





                if (liniepionowe.Count < 1) return;

                zesatwkolumn[0].X0 = liniepionowe[0];
                zesatwkolumn[0].X1 = liniepionowe[0];
                foreach (int pojedyncza in liniepionowe)
                {
                    if ((k < 30) && (pojedyncza > zesatwkolumn[k - 1].X1 + toleracja))
                    {

                        zesatwkolumn[k].X0 = zesatwkolumn[k - 1].X1;
                        zesatwkolumn[k].X1 = pojedyncza;
                        k = k + 1;
                    }
                    ilekolumn = k - 1;
                }
            }

            public void dodajtextdolisty(DBText mText)
            {
                Texty text1 = new Texty();
                text1.X0 = Convert.ToInt32(mText.Position.X);
                text1.Y0 = Convert.ToInt32(mText.Position.Y);
                text1.Text = mText.TextString;

                textycad.Add(text1);
            }



            public void dodajMtextdolisty(MText mText)
            {
                Texty text1 = new Texty();
                text1.X0 = Convert.ToInt32(mText.Location.X);
                text1.Y0 = Convert.ToInt32(mText.Location.Y);
                text1.Text = mText.Text;

                textycad.Add(text1);
            }


            public void dodajBlockreferencedolisty(AttributeReference attRef)
            {
                Texty text1 = new Texty();

                text1.X0 = Convert.ToInt32(attRef.Position.X);
                text1.Y0 = Convert.ToInt32(attRef.Position.Y);
                text1.Text = attRef.TextString;

                textycad.Add(text1);



            }






            public void dodajkolumnyiwierszedonapisow()
            //na podstawie informacji o kolumnach i wierszach sprawdza sie połozenie tekstu i na tej podtsawie dopisuje nr kolumny i wiersza.
            {
                int k;
                int j;
                foreach (Texty text2 in textycad)
                {
                    for (k = 1; k < 30; k++)
                    {
                        if ((text2.X0 >= zesatwkolumn[k].X0) && (text2.X0 < zesatwkolumn[k].X1))
                        {
                            text2.Kol = k;
                            break;
                        }
                    }

                    for (j = 1; j < 20; j++)
                    {
                        if ((text2.Y0 < zesatwwierszy[j].Y0) && (text2.Y0 > zesatwwierszy[j].Y1))
                        {
                            text2.Wier = j;
                            break;
                        }
                    }

                    napisy_z_excel[j, k] = text2.Text;


                }

            }

           


            public void generujopiswcad(string styl1, string text1, int wiersz1, int kolumna1)
            {



            }

        }
        public class Texty
        {
            public int X0 { get; set; }
            public int Y0 { get; set; }
            public int Wier { get; set; }
            public int Kol { get; set; }
            public string Text { get; set; }

            public Texty()
            {
                Text = "-";
            }

            public Texty(string text, int x0, int y0, int wier, int kol)
            {
                Text = text;
                X0 = x0;
                Y0 = y0;
                Wier = wier;
                Kol = kol;

            }
        }

        public class Textydocad
        {
            public int X0 { get; set; }
            public int Y0 { get; set; }
            public int Wiersz { get; set; }
            public int Kolumna { get; set; }


            public int WysTla { get; set; }
            public int SzerTla { get; set; }

            public string Kolor { get; set; }
            public string KolorTla { get; set; }
            public bool czytlo { get; set; }

            public string Text { get; set; }

            public Textydocad()
            {
                Text = "";
                czytlo = false;
            }

            public Textydocad(string text, int wiersz, int kolumna, string kolor, string kolortla)
            {
                Text = text;
                Wiersz = wiersz;
                Kolumna = kolumna;
                Kolor = kolor;
                KolorTla = kolortla;
                czytlo = false;
            }
            public void ustwdomyslnie()
            {
            }
        }





        public class kolumna
        {
            public int X0 { get; set; }
            public int X1 { get; set; }
            public int Y0 { get; set; }
            public int Y1 { get; set; }

            public kolumna()
            {
            }
        }
    }
}
