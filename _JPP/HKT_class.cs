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



                tabelka.punkt5 = acDocEd.GetPoint("\n Wskaż wstawienia schematu").Value;

                }
            }
            //napisycad1 = tabelka.textycad;
            //tabelka1 = tabelka;




            ////TODO tutaj dorobić prodecyrę wszystkich i wywołanie różnych schematów
            rysuj_schemat(tabelka);
        }


        public void rysuj_schemat(Tabelka tabelka)
        {
            int odu=1;

            Point3d point_tmp = tabelka.punkt5;

            for (int k=1; k<= tabelka.ilewierszy; k++)
            {

                //sprawdz ilosc ODU
                //18
                odu = 1;
                if (tabelka.napisy_z_excel[k, 18] == "2") { odu = 2; }
               
                
                //sprawdz czsetotliwosc
                if ((tabelka.napisy_z_excel[k, 6] == "80,0") || (tabelka.napisy_z_excel[k, 6] == "80"))
                    { rysuj_schemat_rifu_80(tabelka, k, point_tmp); 
                     point_tmp =new Point3d(point_tmp.X + 3650, point_tmp.Y, 0);
                        }
                            

                else
                { rysuj_schemat_rifu_normal(tabelka, k,point_tmp);
                    point_tmp = new Point3d(point_tmp.X + 3650, point_tmp.Y, 0);
                }
            }
        }




        public void rysuj_schemat_rifu_80(Tabelka tabelka,int wiersz, Point3d X0Y0)
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
                acPoly.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 19]);

                Autodesk.AutoCAD.DatabaseServices.Polyline acPolyk = new Autodesk.AutoCAD.DatabaseServices.Polyline();
                acPolyk.SetDatabaseDefaults();
                acPolyk.AddVertexAt(0, new Point2d(X0Y0.X, X0Y0.Y + 120+30), 0, 0, 0);
                acPolyk.AddVertexAt(1, new Point2d(X0Y0.X + 30, X0Y0.Y + 120+30), 0, 0, 0);
                acPolyk.AddVertexAt(2, new Point2d(X0Y0.X + 30, X0Y0.Y + 45+30), 0, 0, 0);
                acPolyk.AddVertexAt(3, new Point2d(X0Y0.X + 15, X0Y0.Y + 45+30), 0, 0, 0);
                acPolyk.AddVertexAt(4, new Point2d(X0Y0.X + 15, X0Y0.Y + 30), 0, 0, 0);
                acPolyk.AddVertexAt(5, new Point2d(X0Y0.X - 15, X0Y0.Y + 30), 0, 0, 0);
                acPolyk.AddVertexAt(6, new Point2d(X0Y0.X - 15, X0Y0.Y + 45+30), 0, 0, 0);
                acPolyk.AddVertexAt(7, new Point2d(X0Y0.X - 30, X0Y0.Y + 45+30), 0, 0, 0);
                acPolyk.AddVertexAt(8, new Point2d(X0Y0.X - 30, X0Y0.Y + 120+30), 0, 0, 0);
                acPolyk.Closed = true;

                acPolyk.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 19]);

                Autodesk.AutoCAD.DatabaseServices.Arc acArck0 = new Autodesk.AutoCAD.DatabaseServices.Arc(
                   new Point3d(X0Y0.X , X0Y0.Y + 60, 0), 60 , 3.1415,0 );
                acArck0.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 19]);



                //odu
                Autodesk.AutoCAD.DatabaseServices.Polyline acPoly1 = new Autodesk.AutoCAD.DatabaseServices.Polyline();
                acPoly1.SetDatabaseDefaults();
                acPoly1.AddVertexAt(0, new Point2d(X0Y0.X - 300, X0Y0.Y + 17700), 0, 0, 0);
                acPoly1.AddVertexAt(1, new Point2d(X0Y0.X - 300, X0Y0.Y + 17700 + 560), 0, 0, 0);
                acPoly1.AddVertexAt(2, new Point2d(X0Y0.X - 300 - 960, X0Y0.Y + 17700 + 560), 0, 0, 0);
                acPoly1.AddVertexAt(3, new Point2d(X0Y0.X - 300 - 960, X0Y0.Y + 17700 ), 0, 0, 0);
                acPoly1.Closed = true;
                acPoly1.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 17]);

                Ellipse acEllipse = new Ellipse(new Point3d(X0Y0.X - 300 - 960, X0Y0.Y + 17700 + 280, 0), 40* Vector3d.ZAxis, 160 * Vector3d.YAxis, 0.25, 0, 360 * Math.Atan(1.0) / 45.0);
                acEllipse.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 3]);


                //antena
                Autodesk.AutoCAD.DatabaseServices.Arc acArc = new Autodesk.AutoCAD.DatabaseServices.Arc(
                    new Point3d(X0Y0.X - 300 - 960 - 1050, X0Y0.Y + 17800 + 180, 0), 1010, 3.1415*1.5, 3.1415/2);
                acArc.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 3]);
                Autodesk.AutoCAD.DatabaseServices.Line acline = new Line(acArc.StartPoint, acArc.EndPoint);
                acline.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 3]);

                //kabel eth

                Autodesk.AutoCAD.DatabaseServices.Polyline acPoly2 = new Autodesk.AutoCAD.DatabaseServices.Polyline();
                acPoly2.SetDatabaseDefaults();
                acPoly2.AddVertexAt(0, new Point2d(X0Y0.X + 850, X0Y0.Y + 0), 0, 0, 0);
                acPoly2.AddVertexAt(1, new Point2d(X0Y0.X + 850, X0Y0.Y + 17800+250), 0, 0, 0);
                acPoly2.AddVertexAt(2, new Point2d(X0Y0.X - 300, X0Y0.Y + 17800+250), 0, 0, 0);
                acPoly2.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 19]);

                Autodesk.AutoCAD.DatabaseServices.Polyline acPolyk2 = new Autodesk.AutoCAD.DatabaseServices.Polyline();
                acPolyk2.SetDatabaseDefaults();
                acPolyk2.AddVertexAt(0, new Point2d(X0Y0.X+850, X0Y0.Y + 120 + 30), 0, 0, 0);
                acPolyk2.AddVertexAt(1, new Point2d(X0Y0.X + 30 + 850, X0Y0.Y + 120 + 30), 0, 0, 0);
                acPolyk2.AddVertexAt(2, new Point2d(X0Y0.X + 30 + 850, X0Y0.Y + 45 + 30), 0, 0, 0);
                acPolyk2.AddVertexAt(3, new Point2d(X0Y0.X + 15 + 850, X0Y0.Y + 45 + 30), 0, 0, 0);
                acPolyk2.AddVertexAt(4, new Point2d(X0Y0.X + 15 + 850, X0Y0.Y + 30), 0, 0, 0);
                acPolyk2.AddVertexAt(5, new Point2d(X0Y0.X - 15 + 850, X0Y0.Y + 30), 0, 0, 0);
                acPolyk2.AddVertexAt(6, new Point2d(X0Y0.X - 15 + 850, X0Y0.Y + 45 + 30), 0, 0, 0);
                acPolyk2.AddVertexAt(7, new Point2d(X0Y0.X - 30 + 850, X0Y0.Y + 45 + 30), 0, 0, 0);
                acPolyk2.AddVertexAt(8, new Point2d(X0Y0.X - 30 + 850, X0Y0.Y + 120 + 30), 0, 0, 0);
                acPolyk2.Closed = true;

                acPolyk2.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 19]);

                Autodesk.AutoCAD.DatabaseServices.Arc acArck2 = new Autodesk.AutoCAD.DatabaseServices.Arc(
               new Point3d(X0Y0.X+850, X0Y0.Y + 60, 0), 60, 3.1415, 0);
                acArck2.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 19]);




                //kabel lWL

                Line acline6 = new Line(new Point3d(X0Y0.X - 850, X0Y0.Y + 0,0), new Point3d(X0Y0.X - 850, X0Y0.Y +17700,0));
                acline6.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 22]);

                Autodesk.AutoCAD.DatabaseServices.Polyline acPolyk3 = new Autodesk.AutoCAD.DatabaseServices.Polyline();
                acPolyk3.SetDatabaseDefaults();
                acPolyk3.AddVertexAt(0, new Point2d(X0Y0.X - 850, X0Y0.Y + 120 + 30), 0, 0, 0);
                acPolyk3.AddVertexAt(1, new Point2d(X0Y0.X + 30 - 850, X0Y0.Y + 120 + 30), 0, 0, 0);
                acPolyk3.AddVertexAt(2, new Point2d(X0Y0.X + 30 - 850, X0Y0.Y + 45 + 30), 0, 0, 0);
                acPolyk3.AddVertexAt(3, new Point2d(X0Y0.X + 15 - 850, X0Y0.Y + 45 + 30), 0, 0, 0);
                acPolyk3.AddVertexAt(4, new Point2d(X0Y0.X + 15 - 850, X0Y0.Y + 30), 0, 0, 0);
                acPolyk3.AddVertexAt(5, new Point2d(X0Y0.X - 15 - 850, X0Y0.Y + 30), 0, 0, 0);
                acPolyk3.AddVertexAt(6, new Point2d(X0Y0.X - 15 - 850, X0Y0.Y + 45 + 30), 0, 0, 0);
                acPolyk3.AddVertexAt(7, new Point2d(X0Y0.X - 30 - 850, X0Y0.Y + 45 + 30), 0, 0, 0);
                acPolyk3.AddVertexAt(8, new Point2d(X0Y0.X - 30 - 850, X0Y0.Y + 120 + 30), 0, 0, 0);
                acPolyk3.Closed = true;

                acPolyk3.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 22]);

                Autodesk.AutoCAD.DatabaseServices.Arc acArck3 = new Autodesk.AutoCAD.DatabaseServices.Arc(
                        new Point3d(X0Y0.X - 850, X0Y0.Y + 60, 0), 60, 3.1415, 0);
                acArck3.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 22]);

                //uziemienie
                Autodesk.AutoCAD.DatabaseServices.Line acline2 = new Line(new Point3d(X0Y0.X - 400, X0Y0.Y + 17700,0), new Point3d(X0Y0.X - 400, X0Y0.Y + 17700-230, 0));
                Autodesk.AutoCAD.DatabaseServices.Line acline3 = new Line(new Point3d(X0Y0.X - 400-77, X0Y0.Y + 17700-230, 0), new Point3d(X0Y0.X - 400+77, X0Y0.Y + 17700 - 230, 0));
                Autodesk.AutoCAD.DatabaseServices.Line acline4 = new Line(new Point3d(X0Y0.X - 400 - 56, X0Y0.Y + 17700 - 230-40, 0), new Point3d(X0Y0.X - 400 + 56, X0Y0.Y + 17700 - 230-40, 0));
                Autodesk.AutoCAD.DatabaseServices.Line acline5 = new Line(new Point3d(X0Y0.X - 400 - 20, X0Y0.Y + 17700 - 230-80, 0), new Point3d(X0Y0.X - 400 + 20, X0Y0.Y + 17700 - 230-80, 0));
                
                acline2.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 17]);
                acline3.Layer = acline2.Layer;
                acline4.Layer = acline2.Layer;
                acline5.Layer = acline2.Layer;




                  //kabel dc
                MText acMText = new MText();
                acMText.SetDatabaseDefaults();
                acMText.Rotation = Math.PI / 2;
                acMText.Attachment = AttachmentPoint.MiddleLeft;
                acMText.Location = new Point3d(X0Y0.X , X0Y0.Y + 8800,0);
                

                acMText.Contents = "1xDC" + "\nL=" + tabelka.napisy_z_excel[wiersz, 21] + " m";
                acMText.TextHeight = 250;
                acMText.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 19]);
                
                //kabel eth

                MText acMText3 = new MText();
                acMText3.SetDatabaseDefaults();
                acMText3.Rotation = Math.PI / 2;
                acMText3.Attachment = AttachmentPoint.MiddleLeft;
                acMText3.Location = new Point3d(X0Y0.X +850, X0Y0.Y + 8800, 0);
                

                acMText3.Contents = "1xETH" + "\nL=" + tabelka.napisy_z_excel[wiersz, 21] + " m";
                acMText3.TextHeight = 250;
                acMText3.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 19]);

                //kabel lwl
                MText acMText6 = new MText();
                acMText6.SetDatabaseDefaults();
                acMText6.Rotation = Math.PI / 2;
                acMText6.Attachment = AttachmentPoint.MiddleLeft;
                acMText6.Location = new Point3d(X0Y0.X - 850, X0Y0.Y + 8800, 0);
                

                acMText6.Contents = tabelka.napisy_z_excel[wiersz, 23] + "x" + tabelka.napisy_z_excel[wiersz, 22] + "\n" + "L=" + tabelka.napisy_z_excel[wiersz, 24] + " m";
                acMText6.TextHeight = 250;
                acMText6.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 22]);

                //odu
                MText acMText2 = new MText();
                acMText2.SetDatabaseDefaults();
                acMText2.Rotation = 0;
                acMText2.Attachment = AttachmentPoint.BottomLeft;
          
                acMText2.Location = new Point3d(X0Y0.X-1020, X0Y0.Y+ 17830, 0);
                acMText2.Contents = "ODU";
                acMText2.TextHeight = 250;
                acMText2.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 17]);

                //rifu

                MText acMText4 = new MText();
                acMText4.SetDatabaseDefaults();
                acMText4.Rotation = Math.PI/2;
                acMText4.Attachment = AttachmentPoint.MiddleCenter;

                acMText4.Location = new Point3d(X0Y0.X - 1890, X0Y0.Y + 17800 + 180, 0);
                
                acMText4.Contents = "Rifu %%c" + tabelka.napisy_z_excel[wiersz, 3];
                acMText4.TextHeight = 250;
                acMText4.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 3]);

                //rifu opis

                MText acMText5 = new MText();
                acMText5.SetDatabaseDefaults();
                acMText5.Rotation = 0;
                acMText5.Attachment = AttachmentPoint.BottomCenter;

                acMText5.Location = new Point3d(X0Y0.X - 700, X0Y0.Y + 19700, 0);
                acMText5.Contents = tabelka.napisy_z_excel[wiersz, 1]+", "+ tabelka.napisy_z_excel[wiersz, 8];
                acMText5.TextHeight = 250;
                acMText5.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 3]);




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
                space.AppendEntity(acline6);

                space.AppendEntity(acEllipse);
                space.AppendEntity(acMText);
                space.AppendEntity(acMText2);
                space.AppendEntity(acMText3);
                space.AppendEntity(acMText4);
                space.AppendEntity(acMText5);
                space.AppendEntity(acMText6);
                space.AppendEntity(acPolyk);
                space.AppendEntity(acPolyk2); 
                space.AppendEntity(acPolyk3);
                space.AppendEntity(acArck0);
                space.AppendEntity(acArck2);
                space.AppendEntity(acArck3);


                acTrans.AddNewlyCreatedDBObject(acPoly, true);
                acTrans.AddNewlyCreatedDBObject(acPolyk, true);

                acTrans.AddNewlyCreatedDBObject(acPoly1, true);
                acTrans.AddNewlyCreatedDBObject(acArc, true);
                acTrans.AddNewlyCreatedDBObject(acline, true);
                acTrans.AddNewlyCreatedDBObject(acPoly2, true);
                acTrans.AddNewlyCreatedDBObject(acPolyk2, true);
                acTrans.AddNewlyCreatedDBObject(acPolyk3, true);



                acTrans.AddNewlyCreatedDBObject(acline2, true);
                acTrans.AddNewlyCreatedDBObject(acline3, true);
                acTrans.AddNewlyCreatedDBObject(acline4, true);
                acTrans.AddNewlyCreatedDBObject(acline5, true);
                acTrans.AddNewlyCreatedDBObject(acline6, true);

                acTrans.AddNewlyCreatedDBObject(acEllipse, true);

                acTrans.AddNewlyCreatedDBObject(acMText, true);
                acTrans.AddNewlyCreatedDBObject(acMText2, true);
                acTrans.AddNewlyCreatedDBObject(acMText3, true);
                acTrans.AddNewlyCreatedDBObject(acMText4, true);
                acTrans.AddNewlyCreatedDBObject(acMText5, true);
                acTrans.AddNewlyCreatedDBObject(acMText6, true);

                acTrans.AddNewlyCreatedDBObject(acArck0, true);
                acTrans.AddNewlyCreatedDBObject(acArck2, true);
                acTrans.AddNewlyCreatedDBObject(acArck3, true);


                acTrans.Commit();

                Hatch_object(acEllipse.ObjectId, acEllipse.Layer);
                HatchPolyLine(acPolyk.ObjectId, acPolyk.Layer);

                HatchPolyLine(acPolyk2.ObjectId, acPolyk2.Layer);
                HatchPolyLine(acPolyk3.ObjectId, acPolyk3.Layer);
            }


         }

        public void rysuj_schemat_rifu_normal(Tabelka tabelka, int wiersz, Point3d X0Y0)
        {

            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {

                Editor acDocEd = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;

                ObjectIdCollection acObjIdColl = new ObjectIdCollection();
                var space = (BlockTableRecord)acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite);


                // RG8 1   
                Line acline1 = new Line(new Point3d(X0Y0.X , X0Y0.Y + 0, 0), new Point3d(X0Y0.X , X0Y0.Y + 14800, 0));
                acline1.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 22]);

                space.AppendEntity(acline1);
                acTrans.AddNewlyCreatedDBObject(acline1, true);

                MText acMText1 = new MText();
                acMText1.SetDatabaseDefaults();
                acMText1.Rotation = Math.PI / 2;
                acMText1.Attachment = AttachmentPoint.MiddleLeft;
                acMText1.Location = new Point3d(X0Y0.X , X0Y0.Y + 8800, 0);
                acMText1.Contents = tabelka.napisy_z_excel[wiersz, 19]  + "-Kabel\nL=" + tabelka.napisy_z_excel[wiersz, 21] + " m";
                acMText1.TextHeight = 250;
                acMText1.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 19]);

                space.AppendEntity(acMText1);
                acTrans.AddNewlyCreatedDBObject(acMText1, true);


                Autodesk.AutoCAD.DatabaseServices.Polyline acPolyk10 = new Autodesk.AutoCAD.DatabaseServices.Polyline();
                acPolyk10.SetDatabaseDefaults();
                acPolyk10.AddVertexAt(0, new Point2d(X0Y0.X, X0Y0.Y + 120 + 30), 0, 0, 0);
                acPolyk10.AddVertexAt(1, new Point2d(X0Y0.X + 30, X0Y0.Y + 120 + 30), 0, 0, 0);
                acPolyk10.AddVertexAt(2, new Point2d(X0Y0.X + 30, X0Y0.Y + 45 + 30), 0, 0, 0);
                acPolyk10.AddVertexAt(3, new Point2d(X0Y0.X + 15, X0Y0.Y + 45 + 30), 0, 0, 0);
                acPolyk10.AddVertexAt(4, new Point2d(X0Y0.X + 15, X0Y0.Y + 30), 0, 0, 0);
                acPolyk10.AddVertexAt(5, new Point2d(X0Y0.X - 15, X0Y0.Y + 30), 0, 0, 0);
                acPolyk10.AddVertexAt(6, new Point2d(X0Y0.X - 15, X0Y0.Y + 45 + 30), 0, 0, 0);
                acPolyk10.AddVertexAt(7, new Point2d(X0Y0.X - 30, X0Y0.Y + 45 + 30), 0, 0, 0);
                acPolyk10.AddVertexAt(8, new Point2d(X0Y0.X - 30, X0Y0.Y + 120 + 30), 0, 0, 0);
                acPolyk10.Closed = true;
                acPolyk10.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 19]);

                space.AppendEntity(acPolyk10);
                acTrans.AddNewlyCreatedDBObject(acPolyk10, true);


                Autodesk.AutoCAD.DatabaseServices.Arc acArck10 = new Autodesk.AutoCAD.DatabaseServices.Arc(
                   new Point3d(X0Y0.X, X0Y0.Y + 60, 0), 60, 3.1415, 0);
                acArck10.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 19]);

                space.AppendEntity(acArck10);
                acTrans.AddNewlyCreatedDBObject(acArck10, true);


                Autodesk.AutoCAD.DatabaseServices.Polyline acPolykg10 = new Autodesk.AutoCAD.DatabaseServices.Polyline();
                acPolykg10.SetDatabaseDefaults();
                acPolykg10.AddVertexAt(0, new Point2d(X0Y0.X, X0Y0.Y + 14770 ), 0, 0, 0);
                acPolykg10.AddVertexAt(1, new Point2d(X0Y0.X + 15, X0Y0.Y + 14770), 0, 0, 0);
                acPolykg10.AddVertexAt(2, new Point2d(X0Y0.X + 15, X0Y0.Y + 14770 - 45 ), 0, 0, 0);
                acPolykg10.AddVertexAt(3, new Point2d(X0Y0.X + 30, X0Y0.Y + +14770-45), 0, 0, 0);
                acPolykg10.AddVertexAt(4, new Point2d(X0Y0.X + 30, X0Y0.Y + 14770-120), 0, 0, 0);
                acPolykg10.AddVertexAt(5, new Point2d(X0Y0.X - 30, X0Y0.Y + 14770-120), 0, 0, 0);
                acPolykg10.AddVertexAt(6, new Point2d(X0Y0.X - 30, X0Y0.Y + 14770 -45), 0, 0, 0);
                acPolykg10.AddVertexAt(7, new Point2d(X0Y0.X - 15, X0Y0.Y + +14770-45), 0, 0, 0);
                acPolykg10.AddVertexAt(8, new Point2d(X0Y0.X - 15, X0Y0.Y + 14770), 0, 0, 0);
                acPolykg10.Closed = true;

                acPolykg10.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 19]);

                space.AppendEntity(acPolykg10);
                acTrans.AddNewlyCreatedDBObject(acPolykg10, true);

                Autodesk.AutoCAD.DatabaseServices.Arc acArckg10 = new Autodesk.AutoCAD.DatabaseServices.Arc(
                   new Point3d(X0Y0.X, X0Y0.Y + 14800 - 60, 0), 60, 0, 3.1415);
                acArckg10.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 19]);

                space.AppendEntity(acArckg10);
                acTrans.AddNewlyCreatedDBObject(acArckg10, true);

                // //odu
                Autodesk.AutoCAD.DatabaseServices.Polyline acPolyodu1 = new Autodesk.AutoCAD.DatabaseServices.Polyline();
                acPolyodu1.SetDatabaseDefaults();
                acPolyodu1.AddVertexAt(0, new Point2d(X0Y0.X - 480, X0Y0.Y + 14800), 0, 0, 0);
                acPolyodu1.AddVertexAt(1, new Point2d(X0Y0.X - 480, X0Y0.Y + 14800 + 560), 0, 0, 0);
                acPolyodu1.AddVertexAt(2, new Point2d(X0Y0.X + 480, X0Y0.Y + 14800 + 560), 0, 0, 0);
                acPolyodu1.AddVertexAt(3, new Point2d(X0Y0.X + 480, X0Y0.Y + 14800), 0, 0, 0);
                acPolyodu1.Closed = true;
                acPolyodu1.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 17]);

                space.AppendEntity(acPolyodu1);
                acTrans.AddNewlyCreatedDBObject(acPolyodu1, true);

                Ellipse acEllipseodu1 = new Ellipse(new Point3d(X0Y0.X , X0Y0.Y + 14800 + 560 + 40, 0), 40 * Vector3d.ZAxis, 160 * Vector3d.XAxis, 0.25, 0, 360 * Math.Atan(1.0) / 45.0);
                acEllipseodu1.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 3]);
                
                space.AppendEntity(acEllipseodu1);
                acTrans.AddNewlyCreatedDBObject(acEllipseodu1, true);

                //odu
                MText acMTextodu1 = new MText();
                acMTextodu1.SetDatabaseDefaults();
                acMTextodu1.Rotation = 0;
                acMTextodu1.Attachment = AttachmentPoint.BottomCenter;

                acMTextodu1.Location = new Point3d(X0Y0.X , X0Y0.Y + 14800+170, 0);
                acMTextodu1.Contents = "ODU";
                acMTextodu1.TextHeight = 250;
                acMTextodu1.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 17]);

                space.AppendEntity(acMTextodu1);
                acTrans.AddNewlyCreatedDBObject(acMTextodu1, true);

                //holaiter
                Autodesk.AutoCAD.DatabaseServices.Polyline acPolyhol1 = new Autodesk.AutoCAD.DatabaseServices.Polyline();
                acPolyhol1.SetDatabaseDefaults();
                acPolyhol1.AddVertexAt(0, new Point2d(X0Y0.X , X0Y0.Y + 14800 + 560), 0, 0, 0);
                acPolyhol1.AddVertexAt(1, new Point2d(X0Y0.X , X0Y0.Y + 14800 + 3220), 0, 0, 0);
                acPolyhol1.AddVertexAt(2, new Point2d(X0Y0.X + 1470, X0Y0.Y + 14800 + 3220), 0, 0, 0);

                acPolyhol1.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 13]);
                space.AppendEntity(acPolyhol1);
                acTrans.AddNewlyCreatedDBObject(acPolyhol1, true);

                Ellipse acEllipserifu = new Ellipse(new Point3d(X0Y0.X + 1440, X0Y0.Y + 14800 + 3220-40, 0), 40 * Vector3d.ZAxis, 160 * Vector3d.YAxis, 0.25, 0, 360 * Math.Atan(1.0) / 45.0);
                acEllipserifu.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 3]);
                
                space.AppendEntity(acEllipserifu);
                acTrans.AddNewlyCreatedDBObject(acEllipserifu, true);
                

                //holl
                MText acMTexthol1 = new MText();
                acMTexthol1.SetDatabaseDefaults();
                acMTexthol1.Rotation = Math.PI / 2;
                acMTexthol1.Attachment = AttachmentPoint.MiddleLeft;
                acMTexthol1.Location = new Point3d(X0Y0.X, X0Y0.Y + 14800+1230, 0);
                acMTexthol1.Contents = "Hohlleiter\nL=" + tabelka.napisy_z_excel[wiersz, 14] + " m";
                acMTexthol1.TextHeight = 250;
                acMTexthol1.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 14]);

                space.AppendEntity(acMTexthol1);
                acTrans.AddNewlyCreatedDBObject(acMTexthol1, true);

                //antena
                Autodesk.AutoCAD.DatabaseServices.Arc acArcant1 = new Autodesk.AutoCAD.DatabaseServices.Arc(
                    new Point3d(X0Y0.X + 1470 + 1050, X0Y0.Y + 14800 + 3220-40, 0), 1010, 3.1415 / 2, 3.1415 * 1.5);
                acArcant1.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 3]);

                space.AppendEntity(acArcant1);
                acTrans.AddNewlyCreatedDBObject(acArcant1, true);

                Autodesk.AutoCAD.DatabaseServices.Line aclineAnt1 = new Line(acArcant1.StartPoint, acArcant1.EndPoint);
                aclineAnt1.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 3]);

                space.AppendEntity(aclineAnt1);
                acTrans.AddNewlyCreatedDBObject(aclineAnt1, true);

                 //rifu

                 MText acMText4 = new MText();
                 acMText4.SetDatabaseDefaults();
                 acMText4.Rotation = Math.PI / 2;
                 acMText4.Attachment = AttachmentPoint.MiddleCenter;
                 acMText4.Location = new Point3d(X0Y0.X +2000, X0Y0.Y + 14800 + 3220 - 40, 0);
                 acMText4.Contents = "Rifu %%c" + tabelka.napisy_z_excel[wiersz, 3];
                 acMText4.TextHeight = 250;
                 acMText4.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 3]);

                space.AppendEntity(acMText4);
                acTrans.AddNewlyCreatedDBObject(acMText4, true);

                //rifu opis

                MText acMText5 = new MText();
                 acMText5.SetDatabaseDefaults();
                 acMText5.Rotation = 0;
                 acMText5.Attachment = AttachmentPoint.BottomCenter;

                 acMText5.Location = new Point3d(X0Y0.X + 1300, X0Y0.Y + 19700, 0);
                 acMText5.Contents = tabelka.napisy_z_excel[wiersz, 1] + ", " + tabelka.napisy_z_excel[wiersz, 8];
                 acMText5.TextHeight = 250;
                 acMText5.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 3]);

                space.AppendEntity(acMText5);
                acTrans.AddNewlyCreatedDBObject(acMText5, true);





                Autodesk.AutoCAD.DatabaseServices.Polyline acPolyk20 = new Autodesk.AutoCAD.DatabaseServices.Polyline();
                Autodesk.AutoCAD.DatabaseServices.Polyline acPolykg20 = new Autodesk.AutoCAD.DatabaseServices.Polyline();
                Ellipse acEllipseodu2 = new Ellipse();

                if (tabelka.napisy_z_excel[wiersz, 18]=="2")
                {
                    //rysuj drugie odu 


                    // RG8 1   
                    Line acline2 = new Line(new Point3d(X0Y0.X+1000, X0Y0.Y + 0, 0), new Point3d(X0Y0.X+1000, X0Y0.Y + 14800, 0));
                    acline2.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 22]);

                    space.AppendEntity(acline2);
                    acTrans.AddNewlyCreatedDBObject(acline2, true);

                    MText acMText2 = new MText();
                    acMText2.SetDatabaseDefaults();
                    acMText2.Rotation = Math.PI / 2;
                    acMText2.Attachment = AttachmentPoint.MiddleLeft;
                    acMText2.Location = new Point3d(X0Y0.X+1000, X0Y0.Y + 8800, 0);
                    acMText2.Contents = tabelka.napisy_z_excel[wiersz, 19] + "-Kabel\nL=" + tabelka.napisy_z_excel[wiersz, 21] + " m";
                    acMText2.TextHeight = 250;
                    acMText2.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 19]);

                    space.AppendEntity(acMText2);
                    acTrans.AddNewlyCreatedDBObject(acMText2, true);


                    
                    acPolyk20.SetDatabaseDefaults();
                    acPolyk20.AddVertexAt(0, new Point2d(X0Y0.X+1000, X0Y0.Y + 120 + 30), 0, 0, 0);
                    acPolyk20.AddVertexAt(1, new Point2d(X0Y0.X + 30+1000, X0Y0.Y + 120 + 30), 0, 0, 0);
                    acPolyk20.AddVertexAt(2, new Point2d(X0Y0.X + 30+1000, X0Y0.Y + 45 + 30), 0, 0, 0);
                    acPolyk20.AddVertexAt(3, new Point2d(X0Y0.X + 15+1000, X0Y0.Y + 45 + 30), 0, 0, 0);
                    acPolyk20.AddVertexAt(4, new Point2d(X0Y0.X + 15+1000, X0Y0.Y + 30), 0, 0, 0);
                    acPolyk20.AddVertexAt(5, new Point2d(X0Y0.X - 15+1000, X0Y0.Y + 30), 0, 0, 0);
                    acPolyk20.AddVertexAt(6, new Point2d(X0Y0.X - 15+1000, X0Y0.Y + 45 + 30), 0, 0, 0);
                    acPolyk20.AddVertexAt(7, new Point2d(X0Y0.X - 30+1000, X0Y0.Y + 45 + 30), 0, 0, 0);
                    acPolyk20.AddVertexAt(8, new Point2d(X0Y0.X - 30+1000, X0Y0.Y + 120 + 30), 0, 0, 0);
                    acPolyk20.Closed = true;
                    acPolyk20.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 19]);

                    space.AppendEntity(acPolyk20);
                    acTrans.AddNewlyCreatedDBObject(acPolyk20, true);


                    Autodesk.AutoCAD.DatabaseServices.Arc acArck20 = new Autodesk.AutoCAD.DatabaseServices.Arc(
                       new Point3d(X0Y0.X+1000, X0Y0.Y + 60, 0), 60, 3.1415, 0);
                    acArck20.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 19]);

                    space.AppendEntity(acArck20);
                    acTrans.AddNewlyCreatedDBObject(acArck20, true);


                    
                    acPolykg20.SetDatabaseDefaults();
                    acPolykg20.AddVertexAt(0, new Point2d(X0Y0.X + 1000, X0Y0.Y + 14770), 0, 0, 0);
                    acPolykg20.AddVertexAt(1, new Point2d(X0Y0.X + 15 + 1000, X0Y0.Y + 14770), 0, 0, 0);
                    acPolykg20.AddVertexAt(2, new Point2d(X0Y0.X + 15 + 1000, X0Y0.Y + 14770 - 45), 0, 0, 0);
                    acPolykg20.AddVertexAt(3, new Point2d(X0Y0.X + 30 + 1000, X0Y0.Y + +14770 - 45), 0, 0, 0);
                    acPolykg20.AddVertexAt(4, new Point2d(X0Y0.X + 30 + 1000, X0Y0.Y + 14770 - 120), 0, 0, 0);
                    acPolykg20.AddVertexAt(5, new Point2d(X0Y0.X - 30 + 1000, X0Y0.Y + 14770 - 120), 0, 0, 0);
                    acPolykg20.AddVertexAt(6, new Point2d(X0Y0.X - 30 + 1000, X0Y0.Y + 14770 - 45), 0, 0, 0);
                    acPolykg20.AddVertexAt(7, new Point2d(X0Y0.X - 15 + 1000, X0Y0.Y + +14770 - 45), 0, 0, 0);
                    acPolykg20.AddVertexAt(8, new Point2d(X0Y0.X - 15 + 1000, X0Y0.Y + 14770), 0, 0, 0);
                    acPolykg20.Closed = true;

                    acPolykg20.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 19]);

                    space.AppendEntity(acPolykg20);
                    acTrans.AddNewlyCreatedDBObject(acPolykg20, true);

                    Autodesk.AutoCAD.DatabaseServices.Arc acArckg20 = new Autodesk.AutoCAD.DatabaseServices.Arc(
                       new Point3d(X0Y0.X+1000, X0Y0.Y + 14800 - 60, 0), 60, 0, 3.1415);
                    acArckg20.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 19]);

                    space.AppendEntity(acArckg20);
                    acTrans.AddNewlyCreatedDBObject(acArckg20, true);

                    // //odu
                    Autodesk.AutoCAD.DatabaseServices.Polyline acPolyodu2 = new Autodesk.AutoCAD.DatabaseServices.Polyline();
                    acPolyodu2.SetDatabaseDefaults();
                    acPolyodu2.AddVertexAt(0, new Point2d(X0Y0.X - 480+1000, X0Y0.Y + 14800), 0, 0, 0);
                    acPolyodu2.AddVertexAt(1, new Point2d(X0Y0.X - 480+1000, X0Y0.Y + 14800 + 560), 0, 0, 0);
                    acPolyodu2.AddVertexAt(2, new Point2d(X0Y0.X + 480+1000, X0Y0.Y + 14800 + 560), 0, 0, 0);
                    acPolyodu2.AddVertexAt(3, new Point2d(X0Y0.X + 480+1000, X0Y0.Y + 14800), 0, 0, 0);
                    acPolyodu2.Closed = true;
                    acPolyodu2.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 17]);

                    space.AppendEntity(acPolyodu2);
                    acTrans.AddNewlyCreatedDBObject(acPolyodu2, true);

                    acEllipseodu2 = new Ellipse(new Point3d(X0Y0.X+1000, X0Y0.Y + 14800 + 560 + 40, 0), 40 * Vector3d.ZAxis, 160 * Vector3d.XAxis, 0.25, 0, 360 * Math.Atan(1.0) / 45.0);
                    acEllipseodu2.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 3]);

                    space.AppendEntity(acEllipseodu2);
                    acTrans.AddNewlyCreatedDBObject(acEllipseodu2, true);

                    //odu
                    MText acMTextodu2 = new MText();
                    acMTextodu2.SetDatabaseDefaults();
                    acMTextodu2.Rotation = 0;
                    acMTextodu2.Attachment = AttachmentPoint.BottomCenter;

                    acMTextodu2.Location = new Point3d(X0Y0.X+1000, X0Y0.Y + 14800 + 170, 0);
                    acMTextodu2.Contents = "ODU";
                    acMTextodu2.TextHeight = 250;
                    acMTextodu2.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 17]);

                    space.AppendEntity(acMTextodu2);
                    acTrans.AddNewlyCreatedDBObject(acMTextodu2, true);

                    //holaiter
                    Autodesk.AutoCAD.DatabaseServices.Polyline acPolyhol2 = new Autodesk.AutoCAD.DatabaseServices.Polyline();
                    acPolyhol2.SetDatabaseDefaults();
                    acPolyhol2.AddVertexAt(0, new Point2d(X0Y0.X+1000, X0Y0.Y + 14800 + 560), 0, 0, 0);
                    acPolyhol2.AddVertexAt(1, new Point2d(X0Y0.X+1000, X0Y0.Y + 14800 + 3120), 0, 0, 0);
                    acPolyhol2.AddVertexAt(2, new Point2d(X0Y0.X+ 1470, X0Y0.Y + 14800 + 3120), 0, 0, 0);

                    acPolyhol2.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 13]);
                    space.AppendEntity(acPolyhol2);
                    acTrans.AddNewlyCreatedDBObject(acPolyhol2, true);

                    //holl
                    MText acMTexthol2 = new MText();
                    acMTexthol2.SetDatabaseDefaults();
                    acMTexthol2.Rotation = Math.PI / 2;
                    acMTexthol2.Attachment = AttachmentPoint.MiddleLeft;
                    acMTexthol2.Location = new Point3d(X0Y0.X+1000, X0Y0.Y + 14800 + 1230, 0);
                    acMTexthol2.Contents = "Hohlleiter\nL=" + tabelka.napisy_z_excel[wiersz, 14] + " m";
                    acMTexthol2.TextHeight = 250;
                    acMTexthol2.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 14]);

                    space.AppendEntity(acMTexthol2);
                    acTrans.AddNewlyCreatedDBObject(acMTexthol2, true);
                }





                acTrans.Commit();

                 Hatch_object(acEllipseodu1.ObjectId, acEllipseodu1.Layer);

                HatchPolyLine(acPolyk10.ObjectId, acPolyk10.Layer);

                HatchPolyLine(acPolykg10.ObjectId, acPolyk10.Layer);

                Hatch_object(acEllipserifu.ObjectId, acEllipserifu.Layer);
              

                if (tabelka.napisy_z_excel[wiersz, 18] == "2")
                {
                    Hatch_object(acEllipseodu2.ObjectId, acEllipseodu2.Layer);
                    HatchPolyLine(acPolyk20.ObjectId, acPolyk20.Layer);
                    HatchPolyLine(acPolykg20.ObjectId, acPolykg20.Layer);
                }

                
            }


        }

        private string zmiana_warstwy_tabelka_na_schemat(string wartwa_in)
        {
            string wartwa_wyjsc;

            switch (wartwa_in)
            {
            case "00_AntennenTabelle_Text_NEU":
                    wartwa_wyjsc = "10_30_RiFuAntenne_Option";
                        break;

            case "00_AntennenTabelle":

                    wartwa_wyjsc = "20_30_Kabel_Rifu";
                    break;
            default:
                    wartwa_wyjsc = "0";
                    break;
            }
            return wartwa_wyjsc;
        }

        public static void HatchPolyLine(ObjectId plineId, string layer1)
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
                    hatch.Layer = layer1;
                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                var ed = Application.DocumentManager.MdiActiveDocument.Editor;
                ed.WriteMessage($"{ex.Message}\n{ex.StackTrace}");
            }
        }

        public static void Hatch_object(ObjectId objId, string layer1)
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
                    hatch.Layer = layer1;
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
            public string[,] napisy_z_excel_kolor = new string[20, 30];
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
                text1.Kolor = mText.Layer;

                textycad.Add(text1);
            }



            public void dodajMtextdolisty(MText mText)
            {
                Texty text1 = new Texty();
                text1.X0 = Convert.ToInt32(mText.Location.X);
                text1.Y0 = Convert.ToInt32(mText.Location.Y);
                text1.Text = mText.Text;
                text1.Kolor = mText.Layer;

                textycad.Add(text1);
            }


            public void dodajBlockreferencedolisty(AttributeReference attRef)
            {
                Texty text1 = new Texty();

                text1.X0 = Convert.ToInt32(attRef.Position.X);
                text1.Y0 = Convert.ToInt32(attRef.Position.Y);
                text1.Text = attRef.TextString;
                text1.Kolor = attRef.Layer;
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
                        napisy_z_excel[j, text2.Kol] = text2.Text;
                            napisy_z_excel_kolor[j, text2.Kol] = text2.Kolor;
                            break;
                        }
                    }

                    


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
            public string Kolor { get; set; }
            public Texty()
            {
                Text = "-";
            }

            public Texty(string text, int x0, int y0, int wier, int kol, string kolor)
            {
                Text = text;
                X0 = x0;
                Y0 = y0;
                Wier = wier;
                Kol = kol;
                Kolor = kolor;

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
