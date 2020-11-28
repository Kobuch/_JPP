using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace _JPP
{

    public class HKT_class
    {
        ExcelAll excelAll;
        Tabelka tabelka = new Tabelka();
        Tabelka_plan tabelka_Plan = new Tabelka_plan();
        Obsluga_prop_cad obsluga_Prop_Cad = new Obsluga_prop_cad();

        public void KHT_odczyt_tabeli29kol()
        {

            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            tabelka = new Tabelka();
            List<Texty> napisycad = new List<Texty>();

            // Start a transaction
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Get the current document editor
                Editor acDocEd = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                //  Application.ShowAlertDialog("Wskaż rogi tabekli wg opisu \n 1 - gornylewy róg wartości " +
                //                              "\n 2 - dolnyprawy róg wartości");

                tabelka.punkt1 = acDocEd.GetPoint("\n Wskaż punkt 1 - dolny lewy róg wartości").Value;
                tabelka.punkt2 = acDocEd.GetPoint("\n Wskaż punkt 2 - górny prawy róg wartości").Value;
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

                    Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("Number of objects selected: " +
                                                 acSSet.Count.ToString() + "\n Kolumn: " + tabelka.ilekolumn.ToString() +
                                                  "\n Wierszy: " + tabelka.ilewierszy.ToString());
                }
            }
        }

        public void rysuj_schematpoj(int nr_ant)
        {




            tabelka = new Tabelka();
            Obsluga_prop_cad obsluga_Prop_Cad = new Obsluga_prop_cad();
            tabelka = obsluga_Prop_Cad.odczyt_properties();

            if (tabelka.napisy_z_excel[nr_ant, 1] == null)
            {
                Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("Brak danych dla Rifu: " + nr_ant.ToString());
                return;
            }

            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Editor acDocEd = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;



            Point3d point_tmp = acDocEd.GetPoint("\n Wskaż punkt wstawienia").Value;


            bool integrated = true;

            if ((tabelka.napisy_z_excel[nr_ant, 13] == "2") || (tabelka.napisy_z_excel[nr_ant, 13] == "1")) { integrated = false; }



            //sprawdz czsetotliwosc
            if ((tabelka.napisy_z_excel[nr_ant, 6] == "80,0") || (tabelka.napisy_z_excel[nr_ant, 6] == "80"))
            {
                rysuj_schemat_rifu_80(tabelka, nr_ant, point_tmp);
            }

            else
            {

                if (integrated)
                { rysuj_schemat_rifu_normal_odu_integrated(tabelka, nr_ant, point_tmp); }
                else
                { rysuj_schemat_rifu_normal(tabelka, nr_ant, point_tmp); }

            }
        }

        public void rysuj_schemat()
        {
            //odczut danych z properties
            tabelka = new Tabelka();
            Obsluga_prop_cad obsluga_Prop_Cad = new Obsluga_prop_cad();

            tabelka = obsluga_Prop_Cad.odczyt_properties();


            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Editor acDocEd = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;



            Point3d point_tmp = acDocEd.GetPoint("\n Wskaż punkt wstawienia schematu").Value;

            bool integrated = true;


            for (int k = 1; k <= tabelka.ilewierszy; k++)
            {
                integrated = true;
                if ((tabelka.napisy_z_excel[k, 13] == "2") || (tabelka.napisy_z_excel[k, 13] == "1")) { integrated = false; }


                //sprawdz czsetotliwosc
                if ((tabelka.napisy_z_excel[k, 6] == "80,0") || (tabelka.napisy_z_excel[k, 6] == "80"))
                {
                    rysuj_schemat_rifu_80(tabelka, k, point_tmp);
                    point_tmp = new Point3d(point_tmp.X + 3650, point_tmp.Y, 0);
                }

                else


                {
                    if (integrated)
                    { rysuj_schemat_rifu_normal_odu_integrated(tabelka, k, point_tmp); }
                    else
                    { rysuj_schemat_rifu_normal(tabelka, k, point_tmp); }

                    point_tmp = new Point3d(point_tmp.X + 3650, point_tmp.Y, 0);
                }
            }
        }


        public void rysuj_schemat_rifu_80(Tabelka tabelka, int wiersz, Point3d X0Y0)
        {

            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
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
                acPoly.AddVertexAt(0, new Point2d(X0Y0.X + 0, X0Y0.Y + 0), 0, 0, 0);
                acPoly.AddVertexAt(1, new Point2d(X0Y0.X + 0, X0Y0.Y + 17800), 0, 0, 0);
                acPoly.AddVertexAt(2, new Point2d(X0Y0.X + 300, X0Y0.Y + 17800), 0, 0, 0);
                acPoly.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 20]);

                Autodesk.AutoCAD.DatabaseServices.Polyline acPolyk = new Autodesk.AutoCAD.DatabaseServices.Polyline();
                acPolyk.SetDatabaseDefaults();
                acPolyk.AddVertexAt(0, new Point2d(X0Y0.X, X0Y0.Y + 120 + 30), 0, 0, 0);
                acPolyk.AddVertexAt(1, new Point2d(X0Y0.X + 30, X0Y0.Y + 120 + 30), 0, 0, 0);
                acPolyk.AddVertexAt(2, new Point2d(X0Y0.X + 30, X0Y0.Y + 45 + 30), 0, 0, 0);
                acPolyk.AddVertexAt(3, new Point2d(X0Y0.X + 15, X0Y0.Y + 45 + 30), 0, 0, 0);
                acPolyk.AddVertexAt(4, new Point2d(X0Y0.X + 15, X0Y0.Y + 30), 0, 0, 0);
                acPolyk.AddVertexAt(5, new Point2d(X0Y0.X - 15, X0Y0.Y + 30), 0, 0, 0);
                acPolyk.AddVertexAt(6, new Point2d(X0Y0.X - 15, X0Y0.Y + 45 + 30), 0, 0, 0);
                acPolyk.AddVertexAt(7, new Point2d(X0Y0.X - 30, X0Y0.Y + 45 + 30), 0, 0, 0);
                acPolyk.AddVertexAt(8, new Point2d(X0Y0.X - 30, X0Y0.Y + 120 + 30), 0, 0, 0);
                acPolyk.Closed = true;

                acPolyk.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 20]);

                Autodesk.AutoCAD.DatabaseServices.Arc acArck0 = new Autodesk.AutoCAD.DatabaseServices.Arc(
                   new Point3d(X0Y0.X, X0Y0.Y + 60, 0), 60, 3.1415, 0);
                acArck0.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 20]);



                //odu
                Autodesk.AutoCAD.DatabaseServices.Polyline acPoly1 = new Autodesk.AutoCAD.DatabaseServices.Polyline();
                acPoly1.SetDatabaseDefaults();
                acPoly1.AddVertexAt(0, new Point2d(X0Y0.X + 300, X0Y0.Y + 17700), 0, 0, 0);
                acPoly1.AddVertexAt(1, new Point2d(X0Y0.X + 300, X0Y0.Y + 17700 + 560), 0, 0, 0);
                acPoly1.AddVertexAt(2, new Point2d(X0Y0.X + 300 + 960, X0Y0.Y + 17700 + 560), 0, 0, 0);
                acPoly1.AddVertexAt(3, new Point2d(X0Y0.X + 300 + 960, X0Y0.Y + 17700), 0, 0, 0);
                acPoly1.Closed = true;
                acPoly1.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 18]);

                Ellipse acEllipse = new Ellipse(new Point3d(X0Y0.X + 300 + 960, X0Y0.Y + 17700 + 280, 0), 40 * Vector3d.ZAxis, 160 * Vector3d.YAxis, 0.25, 0, 360 * Math.Atan(1.0) / 45.0);
                acEllipse.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 3]);


                //antena
                Autodesk.AutoCAD.DatabaseServices.Arc acArc = new Autodesk.AutoCAD.DatabaseServices.Arc(
                    new Point3d(X0Y0.X + 300 + 960 + 1050, X0Y0.Y + 17800 + 180, 0), 1010, 3.1415 / 2, 3.1415 * 1.5);
                acArc.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 3]);
                Autodesk.AutoCAD.DatabaseServices.Line acline = new Line(acArc.StartPoint, acArc.EndPoint);
                acline.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 3]);

                //kabel eth

                Autodesk.AutoCAD.DatabaseServices.Polyline acPoly2 = new Autodesk.AutoCAD.DatabaseServices.Polyline();
                acPoly2.SetDatabaseDefaults();
                acPoly2.AddVertexAt(0, new Point2d(X0Y0.X - 850, X0Y0.Y + 0), 0, 0, 0);
                acPoly2.AddVertexAt(1, new Point2d(X0Y0.X - 850, X0Y0.Y + 17800 + 250), 0, 0, 0);
                acPoly2.AddVertexAt(2, new Point2d(X0Y0.X + 300, X0Y0.Y + 17800 + 250), 0, 0, 0);
                acPoly2.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 20]);

                Autodesk.AutoCAD.DatabaseServices.Polyline acPolyk2 = new Autodesk.AutoCAD.DatabaseServices.Polyline();
                acPolyk2.SetDatabaseDefaults();
                acPolyk2.AddVertexAt(0, new Point2d(X0Y0.X + 850, X0Y0.Y + 120 + 30), 0, 0, 0);
                acPolyk2.AddVertexAt(1, new Point2d(X0Y0.X + 30 + 850, X0Y0.Y + 120 + 30), 0, 0, 0);
                acPolyk2.AddVertexAt(2, new Point2d(X0Y0.X + 30 + 850, X0Y0.Y + 45 + 30), 0, 0, 0);
                acPolyk2.AddVertexAt(3, new Point2d(X0Y0.X + 15 + 850, X0Y0.Y + 45 + 30), 0, 0, 0);
                acPolyk2.AddVertexAt(4, new Point2d(X0Y0.X + 15 + 850, X0Y0.Y + 30), 0, 0, 0);
                acPolyk2.AddVertexAt(5, new Point2d(X0Y0.X - 15 + 850, X0Y0.Y + 30), 0, 0, 0);
                acPolyk2.AddVertexAt(6, new Point2d(X0Y0.X - 15 + 850, X0Y0.Y + 45 + 30), 0, 0, 0);
                acPolyk2.AddVertexAt(7, new Point2d(X0Y0.X - 30 + 850, X0Y0.Y + 45 + 30), 0, 0, 0);
                acPolyk2.AddVertexAt(8, new Point2d(X0Y0.X - 30 + 850, X0Y0.Y + 120 + 30), 0, 0, 0);
                acPolyk2.Closed = true;

                acPolyk2.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 20]);

                Autodesk.AutoCAD.DatabaseServices.Arc acArck2 = new Autodesk.AutoCAD.DatabaseServices.Arc(
               new Point3d(X0Y0.X + 850, X0Y0.Y + 60, 0), 60, 3.1415, 0);
                acArck2.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 20]);




                //kabel lWL

                Line acline6 = new Line(new Point3d(X0Y0.X + 850, X0Y0.Y + 0, 0), new Point3d(X0Y0.X + 850, X0Y0.Y + 17700, 0));
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
                Autodesk.AutoCAD.DatabaseServices.Line acline2 = new Line(new Point3d(X0Y0.X + 400, X0Y0.Y + 17700, 0), new Point3d(X0Y0.X + 400, X0Y0.Y + 17700 - 230, 0));
                Autodesk.AutoCAD.DatabaseServices.Line acline3 = new Line(new Point3d(X0Y0.X + 400 - 77, X0Y0.Y + 17700 - 230, 0), new Point3d(X0Y0.X + 400 + 77, X0Y0.Y + 17700 - 230, 0));
                Autodesk.AutoCAD.DatabaseServices.Line acline4 = new Line(new Point3d(X0Y0.X + 400 - 56, X0Y0.Y + 17700 - 230 - 40, 0), new Point3d(X0Y0.X + 400 + 56, X0Y0.Y + 17700 - 230 - 40, 0));
                Autodesk.AutoCAD.DatabaseServices.Line acline5 = new Line(new Point3d(X0Y0.X + 400 - 20, X0Y0.Y + 17700 - 230 - 80, 0), new Point3d(X0Y0.X + 400 + 20, X0Y0.Y + 17700 - 230 - 80, 0));

                acline2.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 18]);
                acline3.Layer = acline2.Layer;
                acline4.Layer = acline2.Layer;
                acline5.Layer = acline2.Layer;




                //kabel dc
                MText acMText = new MText();
                acMText.SetDatabaseDefaults();
                acMText.Rotation = Math.PI / 2;
                acMText.Attachment = AttachmentPoint.MiddleLeft;
                acMText.Location = new Point3d(X0Y0.X, X0Y0.Y + 8800, 0);


                acMText.Contents = "1xDC" + "\nL=" + tabelka.napisy_z_excel[wiersz, 21] + " m";
                acMText.TextHeight = 250;
                acMText.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 21]);

                //kabel eth

                MText acMText3 = new MText();
                acMText3.SetDatabaseDefaults();
                acMText3.Rotation = Math.PI / 2;
                acMText3.Attachment = AttachmentPoint.MiddleLeft;
                acMText3.Location = new Point3d(X0Y0.X - 850, X0Y0.Y + 8800, 0);


                acMText3.Contents = "1xETH" + "\nL=" + tabelka.napisy_z_excel[wiersz, 21] + " m";
                acMText3.TextHeight = 250;
                acMText3.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 21]);

                //kabel lwl
                MText acMText6 = new MText();
                acMText6.SetDatabaseDefaults();
                acMText6.Rotation = Math.PI / 2;
                acMText6.Attachment = AttachmentPoint.MiddleLeft;
                acMText6.Location = new Point3d(X0Y0.X + 850, X0Y0.Y + 8800, 0);


                acMText6.Contents = tabelka.napisy_z_excel[wiersz, 23] + "x" + tabelka.napisy_z_excel[wiersz, 22] + "\n" + "L=" + tabelka.napisy_z_excel[wiersz, 24] + " m";
                acMText6.TextHeight = 250;
                acMText6.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 22]);

                //odu
                MText acMText2 = new MText();
                acMText2.SetDatabaseDefaults();
                acMText2.Rotation = 0;
                acMText2.Attachment = AttachmentPoint.BottomLeft;

                acMText2.Location = new Point3d(X0Y0.X + 1020 - 600, X0Y0.Y + 17830, 0);
                acMText2.Contents = "ODU";
                acMText2.TextHeight = 250;
                acMText2.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 18]);

                //rifu

                MText acMText4 = new MText();
                acMText4.SetDatabaseDefaults();
                acMText4.Rotation = Math.PI / 2;
                acMText4.Attachment = AttachmentPoint.MiddleCenter;

                acMText4.Location = new Point3d(X0Y0.X + 1890, X0Y0.Y + 17800 + 180, 0);

                acMText4.Contents = "Rifu %%c" + tabelka.napisy_z_excel[wiersz, 3];
                acMText4.TextHeight = 250;
                acMText4.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 3]);

                //rifu opis

                MText acMText5 = new MText();
                acMText5.SetDatabaseDefaults();
                acMText5.Rotation = 0;
                acMText5.Attachment = AttachmentPoint.BottomCenter;

                acMText5.Location = new Point3d(X0Y0.X + 700, X0Y0.Y + 19700, 0);
                acMText5.Contents = tabelka.napisy_z_excel[wiersz, 1] + ", " + tabelka.napisy_z_excel[wiersz, 8] + "%%d";
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


        public void rysuj_schemat_rifu_normal_odu_integrated(Tabelka tabelka, int wiersz, Point3d X0Y0)
        {

            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {

                Editor acDocEd = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;

                var space = (BlockTableRecord)acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite);
                Autodesk.AutoCAD.DatabaseServices.Wipeout wipeout = new Autodesk.AutoCAD.DatabaseServices.Wipeout();

                ObjectIdCollection acObjIdColl = new ObjectIdCollection();

                //zasil
                // Create a lightweight polyline
                Autodesk.AutoCAD.DatabaseServices.Polyline acPoly = new Autodesk.AutoCAD.DatabaseServices.Polyline();
                acPoly.SetDatabaseDefaults();
                acPoly.AddVertexAt(0, new Point2d(X0Y0.X + 0, X0Y0.Y + 0), 0, 0, 0);
                acPoly.AddVertexAt(1, new Point2d(X0Y0.X + 0, X0Y0.Y + 17800), 0, 0, 0);
                acPoly.AddVertexAt(2, new Point2d(X0Y0.X + 300, X0Y0.Y + 17800), 0, 0, 0);
                acPoly.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 20]);

                space.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);

                Autodesk.AutoCAD.DatabaseServices.Polyline acPolyk = new Autodesk.AutoCAD.DatabaseServices.Polyline();
                acPolyk.SetDatabaseDefaults();
                acPolyk.AddVertexAt(0, new Point2d(X0Y0.X, X0Y0.Y + 120 + 30), 0, 0, 0);
                acPolyk.AddVertexAt(1, new Point2d(X0Y0.X + 30, X0Y0.Y + 120 + 30), 0, 0, 0);
                acPolyk.AddVertexAt(2, new Point2d(X0Y0.X + 30, X0Y0.Y + 45 + 30), 0, 0, 0);
                acPolyk.AddVertexAt(3, new Point2d(X0Y0.X + 15, X0Y0.Y + 45 + 30), 0, 0, 0);
                acPolyk.AddVertexAt(4, new Point2d(X0Y0.X + 15, X0Y0.Y + 30), 0, 0, 0);
                acPolyk.AddVertexAt(5, new Point2d(X0Y0.X - 15, X0Y0.Y + 30), 0, 0, 0);
                acPolyk.AddVertexAt(6, new Point2d(X0Y0.X - 15, X0Y0.Y + 45 + 30), 0, 0, 0);
                acPolyk.AddVertexAt(7, new Point2d(X0Y0.X - 30, X0Y0.Y + 45 + 30), 0, 0, 0);
                acPolyk.AddVertexAt(8, new Point2d(X0Y0.X - 30, X0Y0.Y + 120 + 30), 0, 0, 0);
                acPolyk.Closed = true;

                acPolyk.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 20]);

                space.AppendEntity(acPolyk);
                acTrans.AddNewlyCreatedDBObject(acPolyk, true);


                Autodesk.AutoCAD.DatabaseServices.Arc acArck0 = new Autodesk.AutoCAD.DatabaseServices.Arc(
                   new Point3d(X0Y0.X, X0Y0.Y + 60, 0), 60, 3.1415, 0);
                acArck0.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 20]);

                space.AppendEntity(acArck0);
                acTrans.AddNewlyCreatedDBObject(acArck0, true);

                Autodesk.AutoCAD.DatabaseServices.Polyline acPolyk4 = new Autodesk.AutoCAD.DatabaseServices.Polyline();
                acPolyk4.SetDatabaseDefaults();
                acPolyk4.AddVertexAt(0, new Point2d(X0Y0.X + 300 - 120 - 30, X0Y0.Y + 17800), 0, 0, 0);
                acPolyk4.AddVertexAt(1, new Point2d(X0Y0.X + 300 - 120 - 30, X0Y0.Y + 17800 + 30), 0, 0, 0);
                acPolyk4.AddVertexAt(2, new Point2d(X0Y0.X + 300 - 45 - 30, X0Y0.Y + 17800 + 30), 0, 0, 0);
                acPolyk4.AddVertexAt(3, new Point2d(X0Y0.X + 300 - 45 - 30, X0Y0.Y + 17800 + 15), 0, 0, 0);
                acPolyk4.AddVertexAt(4, new Point2d(X0Y0.X + 300 - 30, X0Y0.Y + 17800 + 15), 0, 0, 0);
                acPolyk4.AddVertexAt(5, new Point2d(X0Y0.X + 300 - 30, X0Y0.Y + 17800 - 15), 0, 0, 0);
                acPolyk4.AddVertexAt(6, new Point2d(X0Y0.X + 300 - 45 - 30, X0Y0.Y + 17800 - 15), 0, 0, 0);
                acPolyk4.AddVertexAt(7, new Point2d(X0Y0.X + 300 - 45 - 30, X0Y0.Y + 17800 - 30), 0, 0, 0);
                acPolyk4.AddVertexAt(8, new Point2d(X0Y0.X + 300 - 120 - 30, X0Y0.Y + 17800 - 30), 0, 0, 0);
                acPolyk4.Closed = true;

                acPolyk4.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 20]);
                space.AppendEntity(acPolyk4);
                acTrans.AddNewlyCreatedDBObject(acPolyk4, true);


                Autodesk.AutoCAD.DatabaseServices.Arc acArck4 = new Autodesk.AutoCAD.DatabaseServices.Arc(
                   new Point3d(X0Y0.X + 300 - 60, X0Y0.Y + 17800, 0), 60, 3.1415 * 1.5, 3.1415 * 0.5);
                acArck4.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 20]);

                space.AppendEntity(acArck4);
                acTrans.AddNewlyCreatedDBObject(acArck4, true);



                //odu
                Autodesk.AutoCAD.DatabaseServices.Polyline acPoly1 = new Autodesk.AutoCAD.DatabaseServices.Polyline();
                acPoly1.SetDatabaseDefaults();
                acPoly1.AddVertexAt(0, new Point2d(X0Y0.X + 300, X0Y0.Y + 17700), 0, 0, 0);
                acPoly1.AddVertexAt(1, new Point2d(X0Y0.X + 300, X0Y0.Y + 17700 + 560), 0, 0, 0);
                acPoly1.AddVertexAt(2, new Point2d(X0Y0.X + 300 + 960, X0Y0.Y + 17700 + 560), 0, 0, 0);
                acPoly1.AddVertexAt(3, new Point2d(X0Y0.X + 300 + 960, X0Y0.Y + 17700), 0, 0, 0);
                acPoly1.Closed = true;
                acPoly1.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 18]);

                space.AppendEntity(acPoly1);
                acTrans.AddNewlyCreatedDBObject(acPoly1, true);


                Ellipse acEllipse = new Ellipse(new Point3d(X0Y0.X + 300 + 960, X0Y0.Y + 17700 + 280, 0), 40 * Vector3d.ZAxis, 160 * Vector3d.YAxis, 0.25, 0, 360 * Math.Atan(1.0) / 45.0);
                acEllipse.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 3]);

                space.AppendEntity(acEllipse);
                acTrans.AddNewlyCreatedDBObject(acEllipse, true);

                //antena
                Autodesk.AutoCAD.DatabaseServices.Arc acArc = new Autodesk.AutoCAD.DatabaseServices.Arc(
                    new Point3d(X0Y0.X + 300 + 960 + 1050, X0Y0.Y + 17800 + 180, 0), 1010, 3.1415 / 2, 3.1415 * 1.5);
                acArc.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 3]);
                Autodesk.AutoCAD.DatabaseServices.Line acline = new Line(acArc.StartPoint, acArc.EndPoint);
                acline.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 3]);

                space.AppendEntity(acArc);
                space.AppendEntity(acline);

                acTrans.AddNewlyCreatedDBObject(acArc, true);
                acTrans.AddNewlyCreatedDBObject(acline, true);





                //uziemienie
                Autodesk.AutoCAD.DatabaseServices.Line acline2 = new Line(new Point3d(X0Y0.X + 400, X0Y0.Y + 17700, 0), new Point3d(X0Y0.X + 400, X0Y0.Y + 17700 - 230, 0));
                Autodesk.AutoCAD.DatabaseServices.Line acline3 = new Line(new Point3d(X0Y0.X + 400 - 77, X0Y0.Y + 17700 - 230, 0), new Point3d(X0Y0.X + 400 + 77, X0Y0.Y + 17700 - 230, 0));
                Autodesk.AutoCAD.DatabaseServices.Line acline4 = new Line(new Point3d(X0Y0.X + 400 - 56, X0Y0.Y + 17700 - 230 - 40, 0), new Point3d(X0Y0.X + 400 + 56, X0Y0.Y + 17700 - 230 - 40, 0));
                Autodesk.AutoCAD.DatabaseServices.Line acline5 = new Line(new Point3d(X0Y0.X + 400 - 20, X0Y0.Y + 17700 - 230 - 80, 0), new Point3d(X0Y0.X + 400 + 20, X0Y0.Y + 17700 - 230 - 80, 0));

                acline2.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 18]);
                acline3.Layer = acline2.Layer;
                acline4.Layer = acline2.Layer;
                acline5.Layer = acline2.Layer;

                space.AppendEntity(acline2);
                space.AppendEntity(acline3);
                space.AppendEntity(acline4);
                space.AppendEntity(acline5);

                acTrans.AddNewlyCreatedDBObject(acline2, true);
                acTrans.AddNewlyCreatedDBObject(acline3, true);
                acTrans.AddNewlyCreatedDBObject(acline4, true);
                acTrans.AddNewlyCreatedDBObject(acline5, true);



                //kabel dc
                MText acMText = new MText();
                acMText.SetDatabaseDefaults();
                acMText.Rotation = Math.PI / 2;
                acMText.Attachment = AttachmentPoint.MiddleLeft;
                acMText.Location = new Point3d(X0Y0.X, X0Y0.Y + 8800, 0);


                acMText.Contents = tabelka.napisy_z_excel[wiersz, 20] + "x" + tabelka.napisy_z_excel[wiersz, 19] + "\n" + "L=" + tabelka.napisy_z_excel[wiersz, 21] + " m";
                acMText.TextHeight = 250;
                acMText.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 21]);

                space.AppendEntity(acMText);
                acTrans.AddNewlyCreatedDBObject(acMText, true);

                //odu
                MText acMText2 = new MText();
                acMText2.SetDatabaseDefaults();
                acMText2.Rotation = 0;
                acMText2.Attachment = AttachmentPoint.BottomLeft;

                acMText2.Location = new Point3d(X0Y0.X + 1020 - 600, X0Y0.Y + 17830, 0);
                acMText2.Contents = "ODU";
                acMText2.TextHeight = 250;
                acMText2.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 18]);

                space.AppendEntity(acMText2);
                acTrans.AddNewlyCreatedDBObject(acMText2, true);

                //rifu

                MText acMText4 = new MText();
                acMText4.SetDatabaseDefaults();
                acMText4.Rotation = Math.PI / 2;
                acMText4.Attachment = AttachmentPoint.MiddleCenter;

                acMText4.Location = new Point3d(X0Y0.X + 1890, X0Y0.Y + 17800 + 180, 0);

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

                acMText5.Location = new Point3d(X0Y0.X + 700, X0Y0.Y + 19700, 0);
                acMText5.Contents = tabelka.napisy_z_excel[wiersz, 1] + ", " + tabelka.napisy_z_excel[wiersz, 8] + "%%d";
                acMText5.TextHeight = 250;
                acMText5.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 3]);

                space.AppendEntity(acMText5);
                acTrans.AddNewlyCreatedDBObject(acMText5, true);

                // Add the new object to the block table record and the transaction







                acTrans.Commit();

                Hatch_object(acEllipse.ObjectId, acEllipse.Layer);
                HatchPolyLine(acPolyk.ObjectId, acPolyk.Layer);

                HatchPolyLine(acPolyk4.ObjectId, acPolyk4.Layer);

            }


        }
        public void rysuj_schemat_rifu_normal(Tabelka tabelka, int wiersz, Point3d X0Y0)
        {

            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {

                Editor acDocEd = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;

                ObjectIdCollection acObjIdColl = new ObjectIdCollection();
                var space = (BlockTableRecord)acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite);


                // RG8 1   
                Line acline1 = new Line(new Point3d(X0Y0.X, X0Y0.Y + 0, 0), new Point3d(X0Y0.X, X0Y0.Y + 14800, 0));
                acline1.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 20]);

                space.AppendEntity(acline1);
                acTrans.AddNewlyCreatedDBObject(acline1, true);

                MText acMText1 = new MText();
                acMText1.SetDatabaseDefaults();
                acMText1.Rotation = Math.PI / 2;
                acMText1.Attachment = AttachmentPoint.MiddleLeft;
                acMText1.Location = new Point3d(X0Y0.X, X0Y0.Y + 8800, 0);
                acMText1.Contents = tabelka.napisy_z_excel[wiersz, 19] + "-Kabel\nL=" + tabelka.napisy_z_excel[wiersz, 21] + " m";
                acMText1.TextHeight = 250;
                acMText1.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 21]);

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
                acPolyk10.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 20]);

                space.AppendEntity(acPolyk10);
                acTrans.AddNewlyCreatedDBObject(acPolyk10, true);


                Autodesk.AutoCAD.DatabaseServices.Arc acArck10 = new Autodesk.AutoCAD.DatabaseServices.Arc(
                   new Point3d(X0Y0.X, X0Y0.Y + 60, 0), 60, 3.1415, 0);
                acArck10.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 20]);

                space.AppendEntity(acArck10);
                acTrans.AddNewlyCreatedDBObject(acArck10, true);


                Autodesk.AutoCAD.DatabaseServices.Polyline acPolykg10 = new Autodesk.AutoCAD.DatabaseServices.Polyline();
                acPolykg10.SetDatabaseDefaults();
                acPolykg10.AddVertexAt(0, new Point2d(X0Y0.X, X0Y0.Y + 14770), 0, 0, 0);
                acPolykg10.AddVertexAt(1, new Point2d(X0Y0.X + 15, X0Y0.Y + 14770), 0, 0, 0);
                acPolykg10.AddVertexAt(2, new Point2d(X0Y0.X + 15, X0Y0.Y + 14770 - 45), 0, 0, 0);
                acPolykg10.AddVertexAt(3, new Point2d(X0Y0.X + 30, X0Y0.Y + +14770 - 45), 0, 0, 0);
                acPolykg10.AddVertexAt(4, new Point2d(X0Y0.X + 30, X0Y0.Y + 14770 - 120), 0, 0, 0);
                acPolykg10.AddVertexAt(5, new Point2d(X0Y0.X - 30, X0Y0.Y + 14770 - 120), 0, 0, 0);
                acPolykg10.AddVertexAt(6, new Point2d(X0Y0.X - 30, X0Y0.Y + 14770 - 45), 0, 0, 0);
                acPolykg10.AddVertexAt(7, new Point2d(X0Y0.X - 15, X0Y0.Y + +14770 - 45), 0, 0, 0);
                acPolykg10.AddVertexAt(8, new Point2d(X0Y0.X - 15, X0Y0.Y + 14770), 0, 0, 0);
                acPolykg10.Closed = true;

                acPolykg10.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 20]);

                space.AppendEntity(acPolykg10);
                acTrans.AddNewlyCreatedDBObject(acPolykg10, true);

                Autodesk.AutoCAD.DatabaseServices.Arc acArckg10 = new Autodesk.AutoCAD.DatabaseServices.Arc(
                   new Point3d(X0Y0.X, X0Y0.Y + 14800 - 60, 0), 60, 0, 3.1415);
                acArckg10.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 20]);

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
                acPolyodu1.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 18]);

                space.AppendEntity(acPolyodu1);
                acTrans.AddNewlyCreatedDBObject(acPolyodu1, true);


                Ellipse acEllipseodu1 = new Ellipse(new Point3d(X0Y0.X, X0Y0.Y + 14800 + 560 + 40, 0), 40 * Vector3d.ZAxis, 160 * Vector3d.XAxis, 0.25, 0, 360 * Math.Atan(1.0) / 45.0);
                acEllipseodu1.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 18]);

                space.AppendEntity(acEllipseodu1);
                acTrans.AddNewlyCreatedDBObject(acEllipseodu1, true);

                //odu
                MText acMTextodu1 = new MText();
                acMTextodu1.SetDatabaseDefaults();
                acMTextodu1.Rotation = 0;
                acMTextodu1.Attachment = AttachmentPoint.BottomCenter;

                acMTextodu1.Location = new Point3d(X0Y0.X, X0Y0.Y + 14800 + 170, 0);
                acMTextodu1.Contents = "ODU";
                acMTextodu1.TextHeight = 250;
                acMTextodu1.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 18]);

                space.AppendEntity(acMTextodu1);
                acTrans.AddNewlyCreatedDBObject(acMTextodu1, true);

                //uziemienie
                Autodesk.AutoCAD.DatabaseServices.Line acline2 = new Line(new Point3d(X0Y0.X - 280, X0Y0.Y + 14800, 0), new Point3d(X0Y0.X - 280, X0Y0.Y + 14800 - 230, 0));
                Autodesk.AutoCAD.DatabaseServices.Line acline3 = new Line(new Point3d(X0Y0.X - 280 - 77, X0Y0.Y + 14800 - 230, 0), new Point3d(X0Y0.X - 280 + 77, X0Y0.Y + 14800 - 230, 0));
                Autodesk.AutoCAD.DatabaseServices.Line acline4 = new Line(new Point3d(X0Y0.X - 280 - 56, X0Y0.Y + 14800 - 230 - 40, 0), new Point3d(X0Y0.X - 280 + 56, X0Y0.Y + 14800 - 230 - 40, 0));
                Autodesk.AutoCAD.DatabaseServices.Line acline5 = new Line(new Point3d(X0Y0.X - 280 - 20, X0Y0.Y + 14800 - 230 - 80, 0), new Point3d(X0Y0.X - 280 + 20, X0Y0.Y + 14800 - 230 - 80, 0));

                acline2.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 18]);
                acline3.Layer = acline2.Layer;
                acline4.Layer = acline2.Layer;
                acline5.Layer = acline2.Layer;

                space.AppendEntity(acline2);
                space.AppendEntity(acline3);
                space.AppendEntity(acline4);
                space.AppendEntity(acline5);

                acTrans.AddNewlyCreatedDBObject(acline2, true);
                acTrans.AddNewlyCreatedDBObject(acline3, true);
                acTrans.AddNewlyCreatedDBObject(acline4, true);
                acTrans.AddNewlyCreatedDBObject(acline5, true);






                //holaiter
                Autodesk.AutoCAD.DatabaseServices.Polyline acPolyhol1 = new Autodesk.AutoCAD.DatabaseServices.Polyline();
                acPolyhol1.SetDatabaseDefaults();
                acPolyhol1.AddVertexAt(0, new Point2d(X0Y0.X, X0Y0.Y + 14800 + 560), 0, 0, 0);
                acPolyhol1.AddVertexAt(1, new Point2d(X0Y0.X, X0Y0.Y + 14800 + 3220), 0, 0, 0);
                acPolyhol1.AddVertexAt(2, new Point2d(X0Y0.X + 1470, X0Y0.Y + 14800 + 3220), 0, 0, 0);

                acPolyhol1.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 13]);
                space.AppendEntity(acPolyhol1);
                acTrans.AddNewlyCreatedDBObject(acPolyhol1, true);

                Ellipse acEllipserifu = new Ellipse(new Point3d(X0Y0.X + 1440, X0Y0.Y + 14800 + 3220 - 40, 0), 40 * Vector3d.ZAxis, 160 * Vector3d.YAxis, 0.25, 0, 360 * Math.Atan(1.0) / 45.0);
                acEllipserifu.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 3]);

                space.AppendEntity(acEllipserifu);
                acTrans.AddNewlyCreatedDBObject(acEllipserifu, true);


                //holl
                MText acMTexthol1 = new MText();
                acMTexthol1.SetDatabaseDefaults();
                acMTexthol1.Rotation = Math.PI / 2;
                acMTexthol1.Attachment = AttachmentPoint.MiddleLeft;
                acMTexthol1.Location = new Point3d(X0Y0.X, X0Y0.Y + 14800 + 1230, 0);
                acMTexthol1.Contents = "Hohlleiter\nL=" + tabelka.napisy_z_excel[wiersz, 14] + " m";
                acMTexthol1.TextHeight = 250;
                acMTexthol1.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 14]);

                space.AppendEntity(acMTexthol1);
                acTrans.AddNewlyCreatedDBObject(acMTexthol1, true);

                //antena
                Autodesk.AutoCAD.DatabaseServices.Arc acArcant1 = new Autodesk.AutoCAD.DatabaseServices.Arc(
                    new Point3d(X0Y0.X + 1470 + 1050, X0Y0.Y + 14800 + 3220 - 40, 0), 1010, 3.1415 / 2, 3.1415 * 1.5);
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
                acMText4.Location = new Point3d(X0Y0.X + 2000, X0Y0.Y + 14800 + 3220 - 40, 0);
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
                acMText5.Contents = tabelka.napisy_z_excel[wiersz, 1] + ", " + tabelka.napisy_z_excel[wiersz, 8] + "%%d";
                acMText5.TextHeight = 250;
                acMText5.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 3]);

                space.AppendEntity(acMText5);
                acTrans.AddNewlyCreatedDBObject(acMText5, true);





                Autodesk.AutoCAD.DatabaseServices.Polyline acPolyk20 = new Autodesk.AutoCAD.DatabaseServices.Polyline();
                Autodesk.AutoCAD.DatabaseServices.Polyline acPolykg20 = new Autodesk.AutoCAD.DatabaseServices.Polyline();
                Ellipse acEllipseodu2 = new Ellipse();

                if (tabelka.napisy_z_excel[wiersz, 18] == "2")
                {
                    //rysuj drugie odu 


                    // RG8 1   
                    Line acline2a = new Line(new Point3d(X0Y0.X + 1000, X0Y0.Y + 0, 0), new Point3d(X0Y0.X + 1000, X0Y0.Y + 14800, 0));
                    acline2a.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 20]);

                    space.AppendEntity(acline2a);
                    acTrans.AddNewlyCreatedDBObject(acline2a, true);

                    MText acMText2 = new MText();
                    acMText2.SetDatabaseDefaults();
                    acMText2.Rotation = Math.PI / 2;
                    acMText2.Attachment = AttachmentPoint.MiddleLeft;
                    acMText2.Location = new Point3d(X0Y0.X + 1000, X0Y0.Y + 8800, 0);
                    acMText2.Contents = tabelka.napisy_z_excel[wiersz, 19] + "-Kabel\nL=" + tabelka.napisy_z_excel[wiersz, 21] + " m";
                    acMText2.TextHeight = 250;
                    acMText2.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 21]);

                    space.AppendEntity(acMText2);
                    acTrans.AddNewlyCreatedDBObject(acMText2, true);



                    acPolyk20.SetDatabaseDefaults();
                    acPolyk20.AddVertexAt(0, new Point2d(X0Y0.X + 1000, X0Y0.Y + 120 + 30), 0, 0, 0);
                    acPolyk20.AddVertexAt(1, new Point2d(X0Y0.X + 30 + 1000, X0Y0.Y + 120 + 30), 0, 0, 0);
                    acPolyk20.AddVertexAt(2, new Point2d(X0Y0.X + 30 + 1000, X0Y0.Y + 45 + 30), 0, 0, 0);
                    acPolyk20.AddVertexAt(3, new Point2d(X0Y0.X + 15 + 1000, X0Y0.Y + 45 + 30), 0, 0, 0);
                    acPolyk20.AddVertexAt(4, new Point2d(X0Y0.X + 15 + 1000, X0Y0.Y + 30), 0, 0, 0);
                    acPolyk20.AddVertexAt(5, new Point2d(X0Y0.X - 15 + 1000, X0Y0.Y + 30), 0, 0, 0);
                    acPolyk20.AddVertexAt(6, new Point2d(X0Y0.X - 15 + 1000, X0Y0.Y + 45 + 30), 0, 0, 0);
                    acPolyk20.AddVertexAt(7, new Point2d(X0Y0.X - 30 + 1000, X0Y0.Y + 45 + 30), 0, 0, 0);
                    acPolyk20.AddVertexAt(8, new Point2d(X0Y0.X - 30 + 1000, X0Y0.Y + 120 + 30), 0, 0, 0);
                    acPolyk20.Closed = true;
                    acPolyk20.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 20]);

                    space.AppendEntity(acPolyk20);
                    acTrans.AddNewlyCreatedDBObject(acPolyk20, true);


                    Autodesk.AutoCAD.DatabaseServices.Arc acArck20 = new Autodesk.AutoCAD.DatabaseServices.Arc(
                       new Point3d(X0Y0.X + 1000, X0Y0.Y + 60, 0), 60, 3.1415, 0);
                    acArck20.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 20]);

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

                    acPolykg20.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 20]);

                    space.AppendEntity(acPolykg20);
                    acTrans.AddNewlyCreatedDBObject(acPolykg20, true);

                    Autodesk.AutoCAD.DatabaseServices.Arc acArckg20 = new Autodesk.AutoCAD.DatabaseServices.Arc(
                       new Point3d(X0Y0.X + 1000, X0Y0.Y + 14800 - 60, 0), 60, 0, 3.1415);
                    acArckg20.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 20]);

                    space.AppendEntity(acArckg20);
                    acTrans.AddNewlyCreatedDBObject(acArckg20, true);

                    // //odu
                    Autodesk.AutoCAD.DatabaseServices.Polyline acPolyodu2 = new Autodesk.AutoCAD.DatabaseServices.Polyline();
                    acPolyodu2.SetDatabaseDefaults();
                    acPolyodu2.AddVertexAt(0, new Point2d(X0Y0.X - 480 + 1000, X0Y0.Y + 14800), 0, 0, 0);
                    acPolyodu2.AddVertexAt(1, new Point2d(X0Y0.X - 480 + 1000, X0Y0.Y + 14800 + 560), 0, 0, 0);
                    acPolyodu2.AddVertexAt(2, new Point2d(X0Y0.X + 480 + 1000, X0Y0.Y + 14800 + 560), 0, 0, 0);
                    acPolyodu2.AddVertexAt(3, new Point2d(X0Y0.X + 480 + 1000, X0Y0.Y + 14800), 0, 0, 0);
                    acPolyodu2.Closed = true;
                    acPolyodu2.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 18]);

                    space.AppendEntity(acPolyodu2);
                    acTrans.AddNewlyCreatedDBObject(acPolyodu2, true);

                    acEllipseodu2 = new Ellipse(new Point3d(X0Y0.X + 1000, X0Y0.Y + 14800 + 560 + 40, 0), 40 * Vector3d.ZAxis, 160 * Vector3d.XAxis, 0.25, 0, 360 * Math.Atan(1.0) / 45.0);
                    acEllipseodu2.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 18]);

                    space.AppendEntity(acEllipseodu2);
                    acTrans.AddNewlyCreatedDBObject(acEllipseodu2, true);

                    //odu
                    MText acMTextodu2 = new MText();
                    acMTextodu2.SetDatabaseDefaults();
                    acMTextodu2.Rotation = 0;
                    acMTextodu2.Attachment = AttachmentPoint.BottomCenter;

                    acMTextodu2.Location = new Point3d(X0Y0.X + 1000, X0Y0.Y + 14800 + 170, 0);
                    acMTextodu2.Contents = "ODU";
                    acMTextodu2.TextHeight = 250;
                    acMTextodu2.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 18]);

                    space.AppendEntity(acMTextodu2);
                    acTrans.AddNewlyCreatedDBObject(acMTextodu2, true);

                    //uziemienie
                    Autodesk.AutoCAD.DatabaseServices.Line acline2b = new Line(new Point3d(X0Y0.X - 280 + 1000, X0Y0.Y + 14800, 0), new Point3d(X0Y0.X - 280 + 1000, X0Y0.Y + 14800 - 230, 0));
                    Autodesk.AutoCAD.DatabaseServices.Line acline3b = new Line(new Point3d(X0Y0.X - 280 + 1000 - 77, X0Y0.Y + 14800 - 230, 0), new Point3d(X0Y0.X - 280 + 1000 + 77, X0Y0.Y + 14800 - 230, 0));
                    Autodesk.AutoCAD.DatabaseServices.Line acline4b = new Line(new Point3d(X0Y0.X - 280 + 1000 - 56, X0Y0.Y + 14800 - 230 - 40, 0), new Point3d(X0Y0.X - 280 + 1000 + 56, X0Y0.Y + 14800 - 230 - 40, 0));
                    Autodesk.AutoCAD.DatabaseServices.Line acline5b = new Line(new Point3d(X0Y0.X - 280 + 1000 - 20, X0Y0.Y + 14800 - 230 - 80, 0), new Point3d(X0Y0.X - 280 + 1000 + 20, X0Y0.Y + 14800 - 230 - 80, 0));

                    acline2b.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 18]);
                    acline3b.Layer = acline2.Layer;
                    acline4b.Layer = acline2.Layer;
                    acline5b.Layer = acline2.Layer;

                    space.AppendEntity(acline2b);
                    space.AppendEntity(acline3b);
                    space.AppendEntity(acline4b);
                    space.AppendEntity(acline5b);

                    acTrans.AddNewlyCreatedDBObject(acline2b, true);
                    acTrans.AddNewlyCreatedDBObject(acline3b, true);
                    acTrans.AddNewlyCreatedDBObject(acline4b, true);
                    acTrans.AddNewlyCreatedDBObject(acline5b, true);






                    //holaiter
                    Autodesk.AutoCAD.DatabaseServices.Polyline acPolyhol2 = new Autodesk.AutoCAD.DatabaseServices.Polyline();
                    acPolyhol2.SetDatabaseDefaults();
                    acPolyhol2.AddVertexAt(0, new Point2d(X0Y0.X + 1000, X0Y0.Y + 14800 + 560), 0, 0, 0);
                    acPolyhol2.AddVertexAt(1, new Point2d(X0Y0.X + 1000, X0Y0.Y + 14800 + 3120), 0, 0, 0);
                    acPolyhol2.AddVertexAt(2, new Point2d(X0Y0.X + 1470, X0Y0.Y + 14800 + 3120), 0, 0, 0);

                    acPolyhol2.Layer = zmiana_warstwy_tabelka_na_schemat(tabelka.napisy_z_excel_kolor[wiersz, 13]);
                    space.AppendEntity(acPolyhol2);
                    acTrans.AddNewlyCreatedDBObject(acPolyhol2, true);

                    //holl
                    MText acMTexthol2 = new MText();
                    acMTexthol2.SetDatabaseDefaults();
                    acMTexthol2.Rotation = Math.PI / 2;
                    acMTexthol2.Attachment = AttachmentPoint.MiddleLeft;
                    acMTexthol2.Location = new Point3d(X0Y0.X + 1000, X0Y0.Y + 14800 + 1230, 0);
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
            if ((wartwa_in == null) || (wartwa_in == "")) wartwa_in = "0";

            switch (wartwa_in)
            {
                case "00_AntennenTabelle_Text_NEU":
                    wartwa_wyjsc = "10_30_RiFuAntenne_Option";
                    break;

                case "00_AntennenTabelle":

                    wartwa_wyjsc = "20_30_Kabel_Rifu";
                    break;
                default:
                    wartwa_wyjsc = wartwa_in;
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
                var ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
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
                var ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                ed.WriteMessage($"{ex.Message}\n{ex.StackTrace}");
            }
        }

        public void JPP_HKT_RLzcad()
        {
            KHT_odczyt_tabeli29kol();
            Dodaj_properties__z_cad_do_cad();
        }

        public void HKT_odczyt_z_excel()
        {
            //otwórz execla

            excelAll = new ExcelAll();
            excelAll.OpentemplateExcel();

            //odczytaj dane
            excelAll.zexcelodczytdanych();
            excelAll.Excel_close();

            //przygotuj properties

            //spradz czy wyczyszone poprezdnie

            //nie wyczyszkoce

            //wyczyszczone

            Dodaj_properties_do_cad();
        }

        private void Dodaj_properties_do_cad()
        {
            Obsluga_prop_cad obsluga_Prop_Cad = new Obsluga_prop_cad();

            //if ((excelAll.ile_wierszy_w_cad != null) && (excelAll.ile_kolumn_w_cad != null))
            //{
            for (int w = 1; w <= excelAll.ile_wierszy_w_cad; w++)
            {
                for (int k = 1; k <= excelAll.ile_kolumn_w_cad; k++)
                {

                    obsluga_Prop_Cad.setDwgProp("JPP-W" + w + "K" + k, excelAll.napisy_z_excel[w, k]);

                }
            }

            obsluga_Prop_Cad.setDwgProp("JPP-ile_wierszy", excelAll.ile_wierszy_w_cad.ToString()); ;
            obsluga_Prop_Cad.setDwgProp("JPP-ile_kolumn", excelAll.ile_kolumn_w_cad.ToString());

            //}
        }

        private void Dodaj_properties__z_cad_do_cad()
        {
            Obsluga_prop_cad obsluga_Prop_Cad = new Obsluga_prop_cad();

            //if ((tabelka.ilewierszy != null) && (tabelka.ilekolumn != null))
            //{
            for (int w = 1; w <= tabelka.ilewierszy; w++)
            {
                for (int k = 1; k <= tabelka.ilekolumn; k++)
                {

                    obsluga_Prop_Cad.setDwgProp("JPP-W" + w + "K" + k, tabelka.napisy_z_excel[w, k]);
                    obsluga_Prop_Cad.setDwgProp("JPP-LayerW" + w + "K" + k, tabelka.napisy_z_excel_kolor[w, k]);

                }
                //dla jakich kolumn robimy zmiany
                //kolumna 3 dla obu tabel to srednica
                // dla tabeli 20kolumnowej azymut to 7 kolumna dla tameli 29kolumnowej to 8 kolumna
                obsluga_Prop_Cad.setDwgProp("JPP-W" + w + "K3", czysc_liczbe_z_char(obsluga_Prop_Cad.GetCustomProperty("JPP-W" + w + "K3")));
                if (tabelka.ilekolumn == 29)
                {
                    obsluga_Prop_Cad.setDwgProp("JPP-W" + w + "K8", czysc_liczbe_z_char(obsluga_Prop_Cad.GetCustomProperty("JPP-W" + w + "K8")));
                }

                if (tabelka.ilekolumn == 20)
                {
                    obsluga_Prop_Cad.setDwgProp("JPP-W" + w + "K7", czysc_liczbe_z_char(obsluga_Prop_Cad.GetCustomProperty("JPP-W" + w + "K7")));

                }



            }
            obsluga_Prop_Cad.setDwgProp("JPP-ile_wierszy", tabelka.ilewierszy.ToString()); ;
            obsluga_Prop_Cad.setDwgProp("JPP-ile_kolumn", tabelka.ilekolumn.ToString());
            //}
        }

        /// <summary>
        /// czysci string zostawiajac tylkoo wartości liczbowe
        /// </summary>
        /// <param name="liczba_str"></param>
        /// <returns></returns>
        private string czysc_liczbe_z_char(string liczba_str)
        {
            string numericString = string.Empty;

            foreach (var c in liczba_str)
            {
                // Check for numeric characters (hex in this case) or leading or trailing spaces.
                if ((c >= '0' && c <= '9') || (c == '.') || (c == ','))
                {
                    numericString = string.Concat(numericString, c.ToString());
                }
                
            }
            return numericString;
        }




        public void HKT_czysc_properties_jpp()
        {
            Obsluga_prop_cad obsluga_Prop_Cad = new Obsluga_prop_cad();
            obsluga_Prop_Cad.czysc_properties();

        }

        public void HKT_sprawdz_properties_jpp()
        {

            Obsluga_prop_cad obsluga_Prop_Cad = new Obsluga_prop_cad();
            tabelka = obsluga_Prop_Cad.odczyt_properties();
            
            if (tabelka.ilekolumn == 20)
            {
                List<tabelkapokaz20> tabelkapokazs20 = obsluga_Prop_Cad.odczyt_properties_dotabelkipokaz20();
                UserControl2 userControl2 = new UserControl2(tabelkapokazs20, tabelka);
              
                userControl2.Show();

            }
            else
            {

                //todo
                List<tabelkapokaz> tabelkapokazs = obsluga_Prop_Cad.odczyt_properties_dotabelkipokaz();
               
                UserControl1 userControl1 = new UserControl1(tabelkapokazs, tabelka);
                userControl1.Show();
            }
        }
   



        public void pobierz_kierunek_polnocy()
        {

            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Editor acDocEd = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;



            PromptDoubleResult kat = acDocEd.GetAngle("wskaz kierunek połnocy dwa punkty");

            if (kat.Status != PromptStatus.OK) return;


            Obsluga_prop_cad obsluga_Prop_Cad = new Obsluga_prop_cad();
            obsluga_Prop_Cad.setDwgProp("JPP-PN_rad", kat.Value.ToString());
            obsluga_Prop_Cad.setDwgProp("JPP-PN_deg", (kat.Value * 180 / Math.PI).ToString());

        }

        public void tabela_zmiana20na29_z_prop()
        {
            Obsluga_prop_cad obsluga_Prop_Cad = new Obsluga_prop_cad();
            tabelka = obsluga_Prop_Cad.odczyt_properties();

            if (tabelka.ilekolumn == 20)
            {
                List<tabelkapokaz20> tabelkapokazs20 = obsluga_Prop_Cad.odczyt_properties_dotabelkipokaz20();
                List<tabelkapokaz> tabelkapokazs = obsluga_Prop_Cad.przerobtabelepokaz20_na_29(tabelkapokazs20);
                obsluga_Prop_Cad.czysc_properties();
                obsluga_Prop_Cad.przerobtabelepokaz20_na_29_zapisz(tabelkapokazs);
                Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("Wykonano" );

            }
        }

        public void generuj_rzut_1radiolinii(int ant_nr)
        {



            Rzuty_radiolinii rzuty_Radiolinii = new Rzuty_radiolinii();

            Obsluga_prop_cad obsluga_Prop_Cad = new Obsluga_prop_cad();
            tabelka = obsluga_Prop_Cad.odczyt_properties();
            string PN_text = "";
            PN_text = obsluga_Prop_Cad.GetCustomProperty("JPP-PN_rad");
            double PN = Math.PI / 2;
            if (PN_text != null) PN = Convert.ToDouble(PN_text);

            if (tabelka.napisy_z_excel[ant_nr, 1] == null)
            {
                Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("Brak danych dla Rifu: " + ant_nr.ToString());
                return;
            }

            //pobierz wartości położenia

            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                Editor acDocEd = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                Point3d Pointbazowy = acDocEd.GetPoint("\n Wskaż miejsce wstwienia radiolinii").Value;



                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForWrite) as BlockTable;


                // odczyt tyou anteny  13


                if ((tabelka.napisy_z_excel[ant_nr, 13] == "2") || (tabelka.napisy_z_excel[ant_nr, 13] == "1"))
                {
                    //ma odu osobne
                    
                    switch (tabelka.napisy_z_excel[ant_nr, 3])
                    {
                        case "300":
                            rzuty_Radiolinii.rzut_300_osobne_odu(ant_nr, tabelka, PN, Pointbazowy);

                            break;
                        case "600":
                            rzuty_Radiolinii.rzut_600_osobne_odu(ant_nr, tabelka, PN, Pointbazowy);

                            break;

                        case "1200":

                            rzuty_Radiolinii.rzut_1200_osobne_odu(ant_nr, tabelka, PN, Pointbazowy);

                            break;



                    }


                }
                else
                {
                    switch (tabelka.napisy_z_excel[ant_nr, 3])
                    {
                        case "300" :
                            rzuty_Radiolinii.rzut_300_zintegrowane_odu(ant_nr, tabelka, PN, Pointbazowy);
                            break;
                        case "600":
                            rzuty_Radiolinii.rzut_600_zintegrowane_odu(ant_nr, tabelka, PN, Pointbazowy);
                            // 
                            break;

                        case "1200":
                            rzuty_Radiolinii.rzut_1200_osobne_odu(ant_nr, tabelka, PN, Pointbazowy);

                            // RiFu_Marconi_abgesetzt_1200_Grundriss
                            break;



                    }
                }
                acTrans.Commit();


            }
        }
        public void generuj_rzut_wszystkich_radiolinii()
        {

            Rzuty_radiolinii rzuty_Radiolinii = new Rzuty_radiolinii();

            Obsluga_prop_cad obsluga_Prop_Cad = new Obsluga_prop_cad();
            tabelka = obsluga_Prop_Cad.odczyt_properties();
            string PN_text = "";
            PN_text = obsluga_Prop_Cad.GetCustomProperty("JPP-PN_rad");
            double PN = Math.PI / 2;
            if (PN_text != null) PN = Convert.ToDouble(PN_text);


            //pobierz wartości położenia

            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                Editor acDocEd = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                Point3d Pointbazowy = acDocEd.GetPoint("\n Wskaż miejsce wstwienia radiolinii").Value;



                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForWrite) as BlockTable;


                // odczyt tyou anteny  13
                for (int ant_nr = 1; ant_nr <= tabelka.ilewierszy; ant_nr++)
                {



                    if ((tabelka.napisy_z_excel[ant_nr, 13] == "2") || (tabelka.napisy_z_excel[ant_nr, 13] == "1"))
                    {
                        //ma odu osobne

                        switch (tabelka.napisy_z_excel[ant_nr, 3])
                        {
                            case "300":
                                rzuty_Radiolinii.rzut_300_osobne_odu(ant_nr, tabelka, PN, Pointbazowy);

                                break;
                            case "600":
                                rzuty_Radiolinii.rzut_600_osobne_odu(ant_nr, tabelka, PN, Pointbazowy);

                                break;

                            case "1200":

                                rzuty_Radiolinii.rzut_1200_osobne_odu(ant_nr, tabelka, PN, Pointbazowy);

                                break;



                        }


                    }
                    else
                    {
                        switch (tabelka.napisy_z_excel[ant_nr, 3])
                        {
                            case "300":
                                rzuty_Radiolinii.rzut_300_zintegrowane_odu(ant_nr, tabelka, PN, Pointbazowy);

                                break;
                            case "600":
                                rzuty_Radiolinii.rzut_600_zintegrowane_odu(ant_nr, tabelka, PN, Pointbazowy);
                                // 
                                break;

                            case "1200":

                                rzuty_Radiolinii.rzut_1200_osobne_odu(ant_nr, tabelka, PN, Pointbazowy);

                                // RiFu_Marconi_abgesetzt_1200_Grundriss
                                break;



                        }
                    }
                    Pointbazowy = new Point3d(Pointbazowy.X, Pointbazowy.Y - 2000, Pointbazowy.Z);

                }
                acTrans.Commit();

            }
        }

        public void generuj_rzut_rozy_wiatrow()
        {

            Rzuty_radiolinii rzuty_Radiolinii = new Rzuty_radiolinii();

            Obsluga_prop_cad obsluga_Prop_Cad = new Obsluga_prop_cad();
            tabelka = obsluga_Prop_Cad.odczyt_properties();
            string PN_text = "";
            PN_text = obsluga_Prop_Cad.GetCustomProperty("JPP-PN_rad");
            double PN = Math.PI / 2;
            if (PN_text != null) PN = Convert.ToDouble(PN_text);


            //pobierz wartości położenia

            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                Editor acDocEd = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                Point3d Pointbazowy = acDocEd.GetPoint("\n Wskaż miejsce wstwienia rozy wiatrow").Value;



                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForWrite) as BlockTable;


                // odczyt typu anteny  13
                for (int ant_nr = 1; ant_nr <= tabelka.ilewierszy; ant_nr++)
                {
                        rzuty_Radiolinii.rzut_Rifu_rozawiatrow(ant_nr, tabelka, PN, Pointbazowy);
                    

                }
                acTrans.Commit();

            }
        }



        private double ConvertToRadians(double angle)
        {
            return (Math.PI / 180) * angle;
        }



        public void Rysuj_tabelke_w_cad_z_properties()
        {
            RifuCAD rifuCAD = new RifuCAD();

            rifuCAD.odczytajz_dane_zpoperties();
            rifuCAD.generuj_wymiarynowej_tabelki();
            rifuCAD.UniwTabTextUzupelnij_oXY();
            rifuCAD.generuj_nowa_tabelke();
            rifuCAD.wstawiajtexttabelkidoCAD();

        }

        public void akualizacja_Rifu_planowane()
        {
            Obsluga_prop_cad obsluga_Prop_Cad = new Obsluga_prop_cad();
           //odczyt tabeli z properties cadowego
            tabelka = obsluga_Prop_Cad.odczyt_properties();

            if (tabelka.ilekolumn == 20)
            {
                Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("W properties zapisano stara tabelkę 20 kolumnową. Musisz ją wczesniej zamienić na nową 20 kolumnowa");

            }
            else
            {

                //todo
                List<tabelkapokaz> tabelkapokazs = obsluga_Prop_Cad.odczyt_properties_dotabelkipokaz();

                UserControl_plan userControl_plan = new UserControl_plan(tabelka);
                userControl_plan.Show();
            }
        }

    }

    public class Rzuty_radiolinii
    {
        private double ConvertToRadians(double angle)
        {
            return (Math.PI / 180) * angle;
        }

        public void rzut_300_osobne_odu(int ant_nr, Tabelka tabelka, double PN, Point3d Pointbazowy)
        {
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                Editor acDocEd = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;

                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForWrite) as BlockTable;

                if (!acBlkTbl.Has("RiFu_Marconi_abgesetzt_300_Grundriss"))
                {
                    try
                    {
                        // search for a dwg file named 'blockName' in AutoCAD search paths
                        var filename = HostApplicationServices.Current.FindFile("RiFu_Marconi_abgesetzt_300_Grundriss.dwg", acCurDb, FindFileHint.Default);
                        // add the dwg model space as 'blockName' block definition in the current database block table
                        using (var sourceDb = new Database(false, true))
                        {
                            sourceDb.ReadDwgFile(filename, FileOpenMode.OpenForReadAndAllShare, true, "");
                            acCurDb.Insert("RiFu_Marconi_abgesetzt_300_Grundriss", sourceDb, true);
                        }
                    }
                    catch
                    {
                        acDocEd.WriteMessage($"\nBlock RiFu_Marconi_abgesetzt_300_Grundriss.dwg not found.");
                        return;
                    }
                }

                using (var br = new BlockReference(Pointbazowy, acBlkTbl["RiFu_Marconi_abgesetzt_300_Grundriss"]))
                {
                    br.Rotation = ConvertToRadians(-270 - Convert.ToDouble(tabelka.napisy_z_excel[ant_nr, 8].Replace(",", "."))) + PN;
                    var space = (BlockTableRecord)acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite);
                    space.AppendEntity(br);
                    acTrans.AddNewlyCreatedDBObject(br, true);

                    //wstaw opis 

                    MText acMText = new MText();
                    acMText.SetDatabaseDefaults();
                    acMText.SetAttachmentMovingLocation(AttachmentPoint.MiddleCenter);
                    acMText.Location = new Point3d(2000 + Pointbazowy.X, Pointbazowy.Y, 0);
                    acMText.ColorIndex = 7;
                    acMText.Contents = tabelka.napisy_z_excel[ant_nr, 1] + ", %%C" + tabelka.napisy_z_excel[ant_nr, 3] + "\n" + tabelka.napisy_z_excel[ant_nr, 8] + "%%d";
                    acMText.TextHeight = 200;


                    space.AppendEntity(acMText);
                    acTrans.AddNewlyCreatedDBObject(acMText, true);
                }
                acTrans.Commit();
            }
        }
        public void rzut_600_osobne_odu(int ant_nr, Tabelka tabelka, double PN, Point3d Pointbazowy)
        {
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                Editor acDocEd = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;

                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForWrite) as BlockTable;

                if (!acBlkTbl.Has("RiFu_Marconi_abgesetzt_600_Grundriss"))
                {
                    try
                    {
                        // search for a dwg file named 'blockName' in AutoCAD search paths
                        var filename = HostApplicationServices.Current.FindFile("RiFu_Marconi_abgesetzt_600_Grundriss.dwg", acCurDb, FindFileHint.Default);
                        // add the dwg model space as 'blockName' block definition in the current database block table
                        using (var sourceDb = new Database(false, true))
                        {
                            sourceDb.ReadDwgFile(filename, FileOpenMode.OpenForReadAndAllShare, true, "");
                            acCurDb.Insert("RiFu_Marconi_abgesetzt_600_Grundriss", sourceDb, true);
                        }
                    }
                    catch
                    {
                        acDocEd.WriteMessage($"\nBlock RiFu_Marconi_abgesetzt_600_Grundriss.dwg not found.");
                        return;
                    }
                }

                using (var br = new BlockReference(Pointbazowy, acBlkTbl["RiFu_Marconi_abgesetzt_600_Grundriss"]))
                {
                    br.Rotation = ConvertToRadians(-270 - Convert.ToDouble(tabelka.napisy_z_excel[ant_nr, 8].Replace(",", "."))) + PN;
                    var space = (BlockTableRecord)acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite);
                    space.AppendEntity(br);
                    acTrans.AddNewlyCreatedDBObject(br, true);

                    //wstaw opis 

                    MText acMText = new MText();
                    acMText.SetDatabaseDefaults();
                    acMText.SetAttachmentMovingLocation(AttachmentPoint.MiddleCenter);
                    acMText.Location = new Point3d(2000 + Pointbazowy.X, Pointbazowy.Y, 0);
                    acMText.ColorIndex = 7;
                    acMText.Contents = tabelka.napisy_z_excel[ant_nr, 1] + ", %%C" + tabelka.napisy_z_excel[ant_nr, 3] + "\n" + tabelka.napisy_z_excel[ant_nr, 8] + "%%d";
                    acMText.TextHeight = 200;


                    space.AppendEntity(acMText);
                    acTrans.AddNewlyCreatedDBObject(acMText, true);
                }
                acTrans.Commit();
            }
        }

        public void rzut_1200_osobne_odu(int ant_nr, Tabelka tabelka, double PN, Point3d Pointbazowy)
        {
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                Editor acDocEd = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;

                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForWrite) as BlockTable;

                if (!acBlkTbl.Has("RiFu_Marconi_abgesetzt_1200_Grundriss"))
                {
                    try
                    {
                        // search for a dwg file named 'blockName' in AutoCAD search paths
                        var filename = HostApplicationServices.Current.FindFile("RiFu_Marconi_abgesetzt_1200_Grundriss.dwg", acCurDb, FindFileHint.Default);
                        // add the dwg model space as 'blockName' block definition in the current database block table
                        using (var sourceDb = new Database(false, true))
                        {
                            sourceDb.ReadDwgFile(filename, FileOpenMode.OpenForReadAndAllShare, true, "");
                            acCurDb.Insert("RiFu_Marconi_abgesetzt_1200_Grundriss", sourceDb, true);
                        }
                    }
                    catch
                    {
                        acDocEd.WriteMessage($"\nBlock RiFu_Marconi_abgesetzt_1200_Grundriss.dwg not found.");
                        return;
                    }
                }

                using (var br = new BlockReference(Pointbazowy, acBlkTbl["RiFu_Marconi_abgesetzt_1200_Grundriss"]))
                {
                    br.Rotation = ConvertToRadians(-270 - Convert.ToDouble(tabelka.napisy_z_excel[ant_nr, 8].Replace(",", "."))) + PN;
                    var space = (BlockTableRecord)acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite);
                    space.AppendEntity(br);
                    acTrans.AddNewlyCreatedDBObject(br, true);

                    //wstaw opis 

                    MText acMText = new MText();
                    acMText.SetDatabaseDefaults();
                    acMText.SetAttachmentMovingLocation(AttachmentPoint.MiddleCenter);
                    acMText.Location = new Point3d(2000 + Pointbazowy.X, Pointbazowy.Y, 0);
                    acMText.ColorIndex = 7;
                    acMText.Contents = tabelka.napisy_z_excel[ant_nr, 1] + ", %%C" + tabelka.napisy_z_excel[ant_nr, 3] + "\n" + tabelka.napisy_z_excel[ant_nr, 8] + "%%d";
                    acMText.TextHeight = 200;


                    space.AppendEntity(acMText);
                    acTrans.AddNewlyCreatedDBObject(acMText, true);
                }
                acTrans.Commit();
            }
        }
        public void rzut_300_zintegrowane_odu(int ant_nr, Tabelka tabelka, double PN, Point3d Pointbazowy)
        {
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                Editor acDocEd = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;

                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForWrite) as BlockTable;

                if (!acBlkTbl.Has("RiFu_Marconi_integriert_300_Grundriss"))
                {
                    try
                    {
                        // search for a dwg file named 'blockName' in AutoCAD search paths
                        var filename = HostApplicationServices.Current.FindFile("RiFu_Marconi_integriert_300_Grundriss.dwg", acCurDb, FindFileHint.Default);
                        // add the dwg model space as 'blockName' block definition in the current database block table
                        using (var sourceDb = new Database(false, true))
                        {
                            sourceDb.ReadDwgFile(filename, FileOpenMode.OpenForReadAndAllShare, true, "");
                            acCurDb.Insert("RiFu_Marconi_integriert_300_Grundriss", sourceDb, true);
                        }
                    }
                    catch
                    {
                        acDocEd.WriteMessage($"\nBlock RiFu_Marconi_integriert_300_Grundriss.dwg not found.");
                        return;
                    }
                }

                using (var br = new BlockReference(Pointbazowy, acBlkTbl["RiFu_Marconi_integriert_300_Grundriss"]))
                {
                    br.Rotation = ConvertToRadians(-270 - Convert.ToDouble(tabelka.napisy_z_excel[ant_nr, 8].Replace(",", "."))) + PN;
                    var space = (BlockTableRecord)acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite);
                    space.AppendEntity(br);
                    acTrans.AddNewlyCreatedDBObject(br, true);

                    //wstaw opis 

                    MText acMText = new MText();
                    acMText.SetDatabaseDefaults();
                    acMText.SetAttachmentMovingLocation(AttachmentPoint.MiddleCenter);
                    acMText.Location = new Point3d(2000 + Pointbazowy.X, Pointbazowy.Y, 0);
                    acMText.ColorIndex = 7;
                    acMText.Contents = tabelka.napisy_z_excel[ant_nr, 1] + ", %%C" + tabelka.napisy_z_excel[ant_nr, 3] + "\n" + tabelka.napisy_z_excel[ant_nr, 8] + "%%d";
                    acMText.TextHeight = 200;


                    space.AppendEntity(acMText);
                    acTrans.AddNewlyCreatedDBObject(acMText, true);
                }
                acTrans.Commit();
            }
        }
        public void rzut_600_zintegrowane_odu(int ant_nr, Tabelka tabelka, double PN, Point3d Pointbazowy)
        {
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                Editor acDocEd = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;

                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForWrite) as BlockTable;

                if (!acBlkTbl.Has("RiFu_Marconi_integriert_600_Grundriss"))
                {
                    try
                    {
                        // search for a dwg file named 'blockName' in AutoCAD search paths
                        var filename = HostApplicationServices.Current.FindFile("RiFu_Marconi_integriert_600_Grundriss.dwg", acCurDb, FindFileHint.Default);
                        // add the dwg model space as 'blockName' block definition in the current database block table
                        using (var sourceDb = new Database(false, true))
                        {
                            sourceDb.ReadDwgFile(filename, FileOpenMode.OpenForReadAndAllShare, true, "");
                            acCurDb.Insert("RiFu_Marconi_integriert_600_Grundriss", sourceDb, true);
                        }
                    }
                    catch
                    {
                        acDocEd.WriteMessage($"\nBlock RiFu_Marconi_integriert_600_Grundriss.dwg not found.");
                        return;
                    }
                }

                using (var br = new BlockReference(Pointbazowy, acBlkTbl["RiFu_Marconi_integriert_600_Grundriss"]))
                {
                    br.Rotation = ConvertToRadians(-270 - Convert.ToDouble(tabelka.napisy_z_excel[ant_nr, 8].Replace(",", "."))) + PN;
                    var space = (BlockTableRecord)acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite);
                    space.AppendEntity(br);
                    acTrans.AddNewlyCreatedDBObject(br, true);

                    //wstaw opis 

                    MText acMText = new MText();
                    acMText.SetDatabaseDefaults();
                    acMText.SetAttachmentMovingLocation(AttachmentPoint.MiddleCenter);
                    acMText.Location = new Point3d(2000 + Pointbazowy.X, Pointbazowy.Y, 0);
                    acMText.ColorIndex = 7;
                    acMText.Contents = tabelka.napisy_z_excel[ant_nr, 1] + ", %%C" + tabelka.napisy_z_excel[ant_nr, 3] + "\n" + tabelka.napisy_z_excel[ant_nr, 8] + "%%d";
                    acMText.TextHeight = 200;


                    space.AppendEntity(acMText);
                    acTrans.AddNewlyCreatedDBObject(acMText, true);
                }
                acTrans.Commit();
            }
        }


        public void rzut_Rifu_rozawiatrow(int ant_nr, Tabelka tabelka, double PN, Point3d Pointbazowy)
        {
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                Editor acDocEd = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;

                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForWrite) as BlockTable;

                if (!acBlkTbl.Has("Richtungspfeil_dyn"))
                {
                    try
                    {
                        // search for a dwg file named 'blockName' in AutoCAD search paths
                        var filename = HostApplicationServices.Current.FindFile("Richtungspfeil_dyn.dwg", acCurDb, FindFileHint.Default);
                        // add the dwg model space as 'blockName' block definition in the current database block table
                        using (var sourceDb = new Database(false, true))
                        {
                            sourceDb.ReadDwgFile(filename, FileOpenMode.OpenForReadAndAllShare, true, "");
                            acCurDb.Insert("Richtungspfeil_dyn", sourceDb, true);
                        }
                    }
                    catch
                    {
                        acDocEd.WriteMessage($"\nBlock Richtungspfeil_dyn.dwg not found.");
                        return;
                    }
                }

                using (var br = new BlockReference(Pointbazowy, acBlkTbl["Richtungspfeil_dyn"]))
                {


                    br.Rotation = ConvertToRadians(-90 - Convert.ToDouble(tabelka.napisy_z_excel[ant_nr, 8].Replace(",", "."))) + PN;
                    br.TransformBy(Matrix3d.Scaling(100, Pointbazowy));
                    br.Layer = tabelka.napisy_z_excel_kolor[ant_nr, 8];
                    var space = (BlockTableRecord)acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite);
                    space.AppendEntity(br);
                    acTrans.AddNewlyCreatedDBObject(br, true);

                    //wstaw opis 

                    MText acMText = new MText();
                    acMText.SetDatabaseDefaults();
                    acMText.SetAttachmentMovingLocation(AttachmentPoint.MiddleCenter);
                    double Y1 = 3000 * Math.Sin(ConvertToRadians(0 - Convert.ToDouble(tabelka.napisy_z_excel[ant_nr, 8].Replace(",", "."))) + PN);
                    double X1 = 3000 * Math.Cos(ConvertToRadians(0 - Convert.ToDouble(tabelka.napisy_z_excel[ant_nr, 8].Replace(",", "."))) + PN);
                    acMText.Location = new Point3d( X1 + Pointbazowy.X, Y1+Pointbazowy.Y  , 0);
                   
                    acMText.Contents = tabelka.napisy_z_excel[ant_nr, 1] + ", %%C" + tabelka.napisy_z_excel[ant_nr, 3] + "\n" + tabelka.napisy_z_excel[ant_nr, 8] + "%%d";
                    acMText.TextHeight = 200;
                    double obrot = ConvertToRadians(-90 - Convert.ToDouble(tabelka.napisy_z_excel[ant_nr, 8].Replace(",", "."))) + PN;

                    if (br.Rotation >= 0 && br.Rotation < Math.PI)
                    {
                        acMText.Rotation = br.Rotation - 0.5 * Math.PI;
                    }
                    else
                    {
                        acMText.Rotation = br.Rotation + 0.5 * Math.PI;
                    }
                    acMText.Layer = tabelka.napisy_z_excel_kolor[ant_nr, 8];

                    //double obrot = ConvertToRadians( - Convert.ToDouble(tabelka.napisy_z_excel[ant_nr, 8].Replace(",", ".")) ) + PN;
                    ////if ((obrot <0 || obrot>(-Math.PI/2)) && (obrot<(-1.5*Math.PI) || obrot>(-2*Math.PI)))    obrot = obrot - Math.PI; 


                    //acMText.Rotation = ConvertToRadians(180 - Convert.ToDouble(tabelka.napisy_z_excel[ant_nr, 8].Replace(",", "."))) + PN;


                    space.AppendEntity(acMText);
                    acTrans.AddNewlyCreatedDBObject(acMText, true);
                }
                acTrans.Commit();
            }
        }




        private string zmiana_warstwy_tabelka_na_rzutach(string wartwa_in, bool czy_text)
        {
            //    string wartwa_wyjsc;
            //    if ((wartwa_in == null) || (wartwa_in == "")) wartwa_in = "0";

            //    switch (wartwa_in)
            //    {
            //        case "00_AntennenTabelle_Text_NEU":
            //            wartwa_wyjsc = "10_30_RiFuAntenne_Option";
            //            break;

            //        case "00_AntennenTabelle":

            //            wartwa_wyjsc = "20_30_Kabel_Rifu";
            //            break;
            //        default:
            //            wartwa_wyjsc = wartwa_in;
            //            break;
            //    }
            //    return wartwa_wyjsc;
            return null;
        }



    }


    public class Obsluga_prop_cad
    {




        public Tabelka odczyt_properties()
        {
            Tabelka tabelka_tmp = new Tabelka();
            tabelka_tmp.ilewierszy =Convert.ToInt32(GetCustomProperty("JPP-ile_wierszy"));
            tabelka_tmp.ilekolumn = Convert.ToInt32(GetCustomProperty("JPP-ile_kolumn"));
            tabelka_tmp.kierpolnocy_deg = GetCustomProperty("JPP-PN_deg");
            tabelka_tmp.kierpolnocy_rad = GetCustomProperty("JPP-PN_rad");




            //if ((tabelka_tmp.ilewierszy != null) && (tabelka_tmp.ilekolumn != null))
            //{
            for (int w = 1; w <= tabelka_tmp.ilewierszy; w++)
            {
                for (int k = 1; k <= tabelka_tmp.ilekolumn; k++)
                {

                    tabelka_tmp.napisy_z_excel[w, k] = GetCustomProperty("JPP-W" + w + "K" + k);
                    tabelka_tmp.napisy_z_excel_kolor[w, k] = GetCustomProperty("JPP-LayerW" + w + "K" + k);

                }
            }
            //}
            return tabelka_tmp;


        }

        public List<tabelkapokaz> odczyt_properties_dotabelkipokaz()
        {
            Tabelka tabelka_tmp = new Tabelka();
            tabelka_tmp.ilewierszy = Convert.ToInt32(GetCustomProperty("JPP-ile_wierszy"));
            tabelka_tmp.ilekolumn = Convert.ToInt32(GetCustomProperty("JPP-ile_kolumn"));
            List<tabelkapokaz> listapokaz_tmp = new List<tabelkapokaz>();


            //if ((tabelka_tmp.ilewierszy != null) && (tabelka_tmp.ilekolumn != null))
            //{
            for (int w = 1; w <= tabelka_tmp.ilewierszy; w++)
            {
                if (tabelka_tmp.ilekolumn == 29)
                {
                    tabelkapokaz tabelkapok = new tabelkapokaz();
                    tabelkapok.RIFU_NR = GetCustomProperty("JPP-W" + w + "K1");
                    tabelkapok.NETZ = GetCustomProperty("JPP-W" + w + "K2");
                    tabelkapok.RIFU = GetCustomProperty("JPP-W" + w + "K3");
                    tabelkapok.AUFBAU = GetCustomProperty("JPP-W" + w + "K4");
                    tabelkapok.OPTION = GetCustomProperty("JPP-W" + w + "K5");
                    tabelkapok.FREQUENZ = GetCustomProperty("JPP-W" + w + "K6");
                    tabelkapok.Farbe = GetCustomProperty("JPP-W" + w + "K7");
                    tabelkapok.RICHTUNG = GetCustomProperty("JPP-W" + w + "K8");
                    tabelkapok.HÖHE = GetCustomProperty("JPP-W" + w + "K9");
                    tabelkapok.GEGENSTELLE = GetCustomProperty("JPP-W" + w + "K10");
                    tabelkapok.Linknummer = GetCustomProperty("JPP-W" + w + "K11");
                    tabelkapok.HOHLLEITER_TYP = GetCustomProperty("JPP-W" + w + "K12");
                    tabelkapok.HOHLLEITER_ANZAHL = GetCustomProperty("JPP-W" + w + "K13");
                    tabelkapok.HOHLLEITER_LÄNGE = GetCustomProperty("JPP-W" + w + "K14");
                    tabelkapok.HOHLLEITER_AUFBAU = GetCustomProperty("JPP-W" + w + "K15");
                    tabelkapok.HOHLLEITER_OPTION = GetCustomProperty("JPP-W" + w + "K16");
                    tabelkapok.ODU_TYP = GetCustomProperty("JPP-W" + w + "K17");
                    tabelkapok.ODU_ANZAHL = GetCustomProperty("JPP-W" + w + "K18");
                    tabelkapok.DATENKABEL_TYP = GetCustomProperty("JPP-W" + w + "K19");
                    tabelkapok.DATENKABEL_ANZAHL = GetCustomProperty("JPP-W" + w + "K20");
                    tabelkapok.DATENKABEL_LÄNGE = GetCustomProperty("JPP-W" + w + "K21");
                    tabelkapok.POWERKABEL_TYP = GetCustomProperty("JPP-W" + w + "K22");
                    tabelkapok.POWERKABEL_ANZAHL = GetCustomProperty("JPP-W" + w + "K23");
                    tabelkapok.POWERKABEL_LÄNGE = GetCustomProperty("JPP-W" + w + "K24");
                    tabelkapok.EISSCHUTZ = GetCustomProperty("JPP-W" + w + "K25");
                    tabelkapok.STATI_VERDREHUNG = GetCustomProperty("JPP-W" + w + "K26");
                    tabelkapok.ANT_TAEGER_NR = GetCustomProperty("JPP-W" + w + "K27");
                    tabelkapok.ANT_TÄGER_DURCHM = GetCustomProperty("JPP-W" + w + "K28");
                    tabelkapok.BEMERKUNG = GetCustomProperty("JPP-W" + w + "K29");





                    listapokaz_tmp.Add(tabelkapok);
                }


                //}
            }
            return listapokaz_tmp;
        }

        public List<tabelkapokaz20> odczyt_properties_dotabelkipokaz20()
        {
            Tabelka tabelka_tmp = new Tabelka();
            tabelka_tmp.ilewierszy = Convert.ToInt32(GetCustomProperty("JPP-ile_wierszy"));
            tabelka_tmp.ilekolumn = Convert.ToInt32(GetCustomProperty("JPP-ile_kolumn"));
            List<tabelkapokaz20> listapokaz_tmp = new List<tabelkapokaz20>();



            for (int w = 1; w <= tabelka_tmp.ilewierszy; w++)
            {
                if (tabelka_tmp.ilekolumn == 20)
                {
                    tabelkapokaz20 tabelkapok = new tabelkapokaz20();
                    tabelkapok.RIFU_NR = GetCustomProperty("JPP-W" + w + "K1");
                    tabelkapok.NETZ = GetCustomProperty("JPP-W" + w + "K2");
                    tabelkapok.RIFU = GetCustomProperty("JPP-W" + w + "K3");
                    tabelkapok.AUFBAU = GetCustomProperty("JPP-W" + w + "K4");
                    tabelkapok.OPTION = GetCustomProperty("JPP-W" + w + "K5");
                    tabelkapok.Farbe = GetCustomProperty("JPP-W" + w + "K6");
                    tabelkapok.RICHTUNG = GetCustomProperty("JPP-W" + w + "K7");
                    tabelkapok.HÖHE = GetCustomProperty("JPP-W" + w + "K8");
                    tabelkapok.DATENKABEL_TYP = GetCustomProperty("JPP-W" + w + "K9");
                    tabelkapok.DATENKABEL_ANZAHL = GetCustomProperty("JPP-W" + w + "K10");
                    tabelkapok.DATENKABEL_LÄNGE = GetCustomProperty("JPP-W" + w + "K11");
                    tabelkapok.Farbe_kable = GetCustomProperty("JPP-W" + w + "K12");

                    tabelkapok.GEGENSTELLE = GetCustomProperty("JPP-W" + w + "K13");
                    tabelkapok.Linknummer = GetCustomProperty("JPP-W" + w + "K14");


                    tabelkapok.HOHLLEITER_LÄNGE = GetCustomProperty("JPP-W" + w + "K15");


                    tabelkapok.ODU_TYP = GetCustomProperty("JPP-W" + w + "K16");
                    tabelkapok.ODU_ANZAHL = GetCustomProperty("JPP-W" + w + "K17");


                    tabelkapok.ANT_TAEGER_NR = GetCustomProperty("JPP-W" + w + "K18");
                    tabelkapok.ANT_TÄGER_DURCHM = GetCustomProperty("JPP-W" + w + "K19");
                    tabelkapok.BEMERKUNG = GetCustomProperty("JPP-W" + w + "K20");

                    listapokaz_tmp.Add(tabelkapok);
                }
            }
            return listapokaz_tmp;
        }

        public List<tabelkapokaz> przerobtabelepokaz20_na_29(List<tabelkapokaz20> tabelkapokaz20s)

        {

                        ///TODO
            List<tabelkapokaz> listapokaz_tmp = new List<tabelkapokaz>();

            foreach (tabelkapokaz20 tabelkapokaz20 in tabelkapokaz20s)
            {

                tabelkapokaz tabelkapok = new tabelkapokaz();
                tabelkapok.RIFU_NR = tabelkapokaz20.RIFU_NR;
                tabelkapok.NETZ = tabelkapokaz20.NETZ;
                tabelkapok.RIFU = tabelkapokaz20.RIFU;
                tabelkapok.AUFBAU = tabelkapokaz20.AUFBAU;
                tabelkapok.OPTION = tabelkapokaz20.OPTION;
                tabelkapok.FREQUENZ = "";
                tabelkapok.Farbe = tabelkapokaz20.Farbe;
                tabelkapok.RICHTUNG = tabelkapokaz20.RICHTUNG;
                tabelkapok.HÖHE = tabelkapokaz20.HÖHE;
                tabelkapok.GEGENSTELLE = tabelkapokaz20.GEGENSTELLE;
                tabelkapok.Linknummer = tabelkapokaz20.Linknummer;
                tabelkapok.HOHLLEITER_TYP = "";
                tabelkapok.HOHLLEITER_ANZAHL = "";
                tabelkapok.HOHLLEITER_LÄNGE = tabelkapokaz20.HOHLLEITER_LÄNGE;
                tabelkapok.HOHLLEITER_AUFBAU = "";
                tabelkapok.HOHLLEITER_OPTION = "";
                tabelkapok.ODU_TYP = tabelkapokaz20.ODU_TYP;
                tabelkapok.ODU_ANZAHL = tabelkapokaz20.ODU_ANZAHL;
                tabelkapok.DATENKABEL_TYP = tabelkapokaz20.DATENKABEL_TYP;
                tabelkapok.DATENKABEL_ANZAHL = tabelkapokaz20.DATENKABEL_ANZAHL;
                tabelkapok.DATENKABEL_LÄNGE = tabelkapokaz20.DATENKABEL_LÄNGE;
                tabelkapok.POWERKABEL_TYP = "";
                tabelkapok.POWERKABEL_ANZAHL = "";
                tabelkapok.POWERKABEL_LÄNGE = "";
                tabelkapok.EISSCHUTZ = "";
                tabelkapok.STATI_VERDREHUNG = "";
                tabelkapok.ANT_TAEGER_NR = tabelkapokaz20.ANT_TAEGER_NR;
                tabelkapok.ANT_TÄGER_DURCHM = tabelkapokaz20.ANT_TÄGER_DURCHM;
                tabelkapok.BEMERKUNG = tabelkapokaz20.BEMERKUNG;


                listapokaz_tmp.Add(tabelkapok);

            }

            return listapokaz_tmp;
        }


       public void  przerobtabelepokaz20_na_29_zapisz(List<tabelkapokaz> tabelkapokazs)
        {

            
            //czysc_properties();
            int w = 0;
            foreach (tabelkapokaz tabelkapok in tabelkapokazs)
            {
                w = w + 1;
                setDwgProp("JPP-W" + w + "K1", tabelkapok.RIFU_NR);
                setDwgProp("JPP-W" + w + "K2", tabelkapok.NETZ);
                //if ((tabelkapok.RIFU.Length > 3) && (tabelkapok.RIFU.Contains("%%C")))
                   //{
                     tabelkapok.RIFU= tabelkapok.RIFU.Replace("%%C","");
                 
                setDwgProp("JPP-W" + w + "K3", tabelkapok.RIFU);
                
                setDwgProp("JPP-W" + w + "K4", tabelkapok.AUFBAU);
                setDwgProp("JPP-W" + w + "K5", tabelkapok.OPTION);
                setDwgProp("JPP-W" + w + "K6", tabelkapok.FREQUENZ);
                setDwgProp("JPP-W" + w + "K7", tabelkapok.Farbe);
                setDwgProp("JPP-W" + w + "K8", tabelkapok.RICHTUNG);
                setDwgProp("JPP-W" + w + "K9", tabelkapok.HÖHE);
                setDwgProp("JPP-W" + w + "K10", tabelkapok.GEGENSTELLE);
                setDwgProp("JPP-W" + w + "K11", tabelkapok.Linknummer);
                setDwgProp("JPP-W" + w + "K12", tabelkapok.HOHLLEITER_TYP);
                setDwgProp("JPP-W" + w + "K13", tabelkapok.HOHLLEITER_ANZAHL);
                setDwgProp("JPP-W" + w + "K14", tabelkapok.HOHLLEITER_LÄNGE);
                setDwgProp("JPP-W" + w + "K15", tabelkapok.HOHLLEITER_AUFBAU);
                setDwgProp("JPP-W" + w + "K16", tabelkapok.HOHLLEITER_OPTION);
                setDwgProp("JPP-W" + w + "K17", tabelkapok.ODU_TYP);
                setDwgProp("JPP-W" + w + "K18", tabelkapok.ODU_ANZAHL);
                setDwgProp("JPP-W" + w + "K19", tabelkapok.DATENKABEL_TYP);
                setDwgProp("JPP-W" + w + "K20", tabelkapok.DATENKABEL_ANZAHL);
                setDwgProp("JPP-W" + w + "K21", tabelkapok.DATENKABEL_LÄNGE);
                setDwgProp("JPP-W" + w + "K22", tabelkapok.POWERKABEL_TYP);

                setDwgProp("JPP-W" + w + "K23", tabelkapok.POWERKABEL_ANZAHL);
                setDwgProp("JPP-W" + w + "K24", tabelkapok.POWERKABEL_LÄNGE);
                setDwgProp("JPP-W" + w + "K25", tabelkapok.EISSCHUTZ);
                setDwgProp("JPP-W" + w + "K26", tabelkapok.STATI_VERDREHUNG);
                setDwgProp("JPP-W" + w + "K27", tabelkapok.ANT_TAEGER_NR);
                setDwgProp("JPP-W" + w + "K28", tabelkapok.ANT_TÄGER_DURCHM);
                setDwgProp("JPP-W" + w + "K29", tabelkapok.BEMERKUNG);
            
            }
           
            setDwgProp("JPP-ile_wierszy", w.ToString()); 
            setDwgProp("JPP-ile_kolumn", "29");
        }


        public void czysc_properties()
        {
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;


            IDictionaryEnumerator denum = acCurDb.SummaryInfo.CustomProperties;

            while (denum.MoveNext())
            {
                DictionaryEntry entry = denum.Entry;

                if ((entry.Key.ToString().Length > 4) && (entry.Key.ToString().ToUpper().Substring(0, 4) == "JPP-"))
                {
                    DatabaseSummaryInfoBuilder dpbuilder = new DatabaseSummaryInfoBuilder(acCurDb.SummaryInfo);
                    IDictionary customProps = dpbuilder.CustomPropertyTable;
                    customProps.Remove(entry.Key);
                    acCurDb.SummaryInfo = dpbuilder.ToDatabaseSummaryInfo();

                }
            }
        }


        public string GetCustomProperty(string key)
        {
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
           
            Database acCurDb = acDoc.Database;

            DatabaseSummaryInfoBuilder sumInfo = new DatabaseSummaryInfoBuilder(acCurDb.SummaryInfo);
            IDictionary custProps = sumInfo.CustomPropertyTable;
            var napis = (string)custProps[key];
            if (!string.IsNullOrEmpty(napis))
            {
                string wartosc = string.Copy(napis);
            
                return wartosc;
            }

            else return "";

            //return (string)custProps[key];
        }

        public void setDwgProp(string propkey, string propval)
        {
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            acDoc.LockDocument();
            Database acCurDb = acDoc.Database;
            bool jest = false;

            IDictionaryEnumerator denum = acCurDb.SummaryInfo.CustomProperties;
          
           
            while (denum.MoveNext())
            {
                DictionaryEntry entry = denum.Entry;
                if (entry.Key.ToString().ToUpper() == propkey.ToUpper())
                {
                    DatabaseSummaryInfoBuilder dpbuilder = new DatabaseSummaryInfoBuilder(acCurDb.SummaryInfo);
                    IDictionary customProps = dpbuilder.CustomPropertyTable;
                    if (customProps.Contains(entry.Key))
                    {
                        customProps[entry.Key] = propval;
                        jest = true;
                       
                    }

                     acCurDb.SummaryInfo = dpbuilder.ToDatabaseSummaryInfo();
                }
            }
            if (!jest)
            {
                DatabaseSummaryInfoBuilder dpbuilder = new DatabaseSummaryInfoBuilder(acCurDb.SummaryInfo);
                IDictionary customProps = dpbuilder.CustomPropertyTable;
                customProps.Add(propkey, propval);
                acCurDb.SummaryInfo = dpbuilder.ToDatabaseSummaryInfo();
            }

           
        }


    }

    public class RifuCAD
    {

        public Tabelka rifutabelka2 = new Tabelka();
        // public Tabelka rifutabelka_old = new Tabelka();
        public List<Textydocad> textydocads = new List<Textydocad>();

        public Point3d Pointbazowy = new Point3d(0, 0, 0);
        public Autodesk.AutoCAD.ApplicationServices.Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        public Database acCurDb;

        public RifuCAD()
        {
            acDoc.LockDocument();
        }

        public void odczytajz_dane_zpoperties()
        {
            Obsluga_prop_cad obsluga_Prop_Cad = new Obsluga_prop_cad();
            rifutabelka2 = obsluga_Prop_Cad.odczyt_properties();

            for (int w = 1; w <= rifutabelka2.ilewierszy; w++)
            {
                for (int k = 1; k <= rifutabelka2.ilekolumn; k++)
                {
                    textydocads.Add(new Textydocad(rifutabelka2.napisy_z_excel[w, k], w, k, "", ""));



                }
            }

        }



        public void UniwTabTextUzupelnij_oXY()
        //uzupełnia objekty tekstiowe o wartości współrzednych X i Y na podstawie szerokości i wysokości tabelki zapisanej jako Rifutabeka2
        {
            foreach (Textydocad textydocad in textydocads)
            {
                textydocad.X0 = (rifutabelka2.zesatwkolumn[textydocad.Kolumna].X1 + rifutabelka2.zesatwkolumn[textydocad.Kolumna].X0) / 2;
                textydocad.SzerTla = rifutabelka2.zesatwkolumn[textydocad.Kolumna].X1 - rifutabelka2.zesatwkolumn[textydocad.Kolumna].X0;

                textydocad.Y0 = (rifutabelka2.zesatwwierszy[textydocad.Wiersz].Y0 + rifutabelka2.zesatwwierszy[textydocad.Wiersz].Y1) / 2;
                textydocad.WysTla = (rifutabelka2.zesatwwierszy[textydocad.Wiersz].Y0 - rifutabelka2.zesatwwierszy[textydocad.Wiersz].Y1);
            }
        }



        public void wstawiajtexttabelkidoCAD()
        //wstawia opisy dla sterj tabelki 20 kolumn
        {
            acCurDb = acDoc.Database;


       





            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                Autodesk.AutoCAD.Colors.Color cl = new Autodesk.AutoCAD.Colors.Color();
                Editor acDocEd = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;

                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;

                // Open the Block table record Model space for write
                BlockTableRecord acBlkTblRec;
                acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                OpenMode.ForWrite) as BlockTableRecord;

                // check if the block table already has the 'blockName'" block

                foreach (Textydocad textydocad in textydocads)
                {
                    // Create a multiline text object
                    MText acMText = new MText();
                    acMText.SetDatabaseDefaults();
                    acMText.SetAttachmentMovingLocation(AttachmentPoint.MiddleCenter);
                    acMText.Location = new Point3d(textydocad.X0 + Pointbazowy.X, textydocad.Y0 + Pointbazowy.Y, 0);
                    acMText.ColorIndex = 7;
                    acMText.Contents = textydocad.Text;
                    acMText.TextHeight = 200;

                    if (textydocad.KolorTla == "65535")
                    {
                        acMText.ShowBorders = true;
                        acMText.BackgroundFill = true;
                        acMText.BackgroundFillColor = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 2);
                        acMText.BackgroundScaleFactor = 1;
                        acMText.UseBackgroundColor = false;
                        acMText.ColorIndex = 1;
                    }

                    acMText.Width = (double)textydocad.SzerTla;
                    acMText.Height = (double)textydocad.WysTla;

                    acBlkTblRec.AppendEntity(acMText);
                    acTrans.AddNewlyCreatedDBObject(acMText, true);
                }

                acTrans.Commit();

            }

        }

        //  rifu_ant_tab_n
        //AntennenTabelle_Rifu_Kopf_11
        public void generuj_nowa_tabelke()
        {
            // wstawia blok tabeli do cad 
            acCurDb = acDoc.Database;

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                Editor acDocEd = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                // Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("Wskaż miejsce wstwienia tebelki");

                Pointbazowy = acDocEd.GetPoint("\n Wskaż miejsce wstwienia tebelki").Value;

                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForWrite) as BlockTable;



                if (!acBlkTbl.Has("AntennenTabelle_Rifu_Kopf"))
                {
                    try
                    {
                        // search for a dwg file named 'blockName' in AutoCAD search paths
                        var filename = HostApplicationServices.Current.FindFile("AntennenTabelle_Rifu_Kopf.dwg", acCurDb, FindFileHint.Default);
                        // add the dwg model space as 'blockName' block definition in the current database block table
                        using (var sourceDb = new Database(false, true))
                        {
                            sourceDb.ReadDwgFile(filename, FileOpenMode.OpenForReadAndAllShare, true, "");
                            acCurDb.Insert("AntennenTabelle_Rifu_Kopf", sourceDb, true);
                        }
                    }
                    catch
                    {
                        acDocEd.WriteMessage($"\nBlock AntennenTabelle_Rifu_Kopf.dwg not found.");
                        return;
                    }
                }

                if (!acBlkTbl.Has("tabrifu1"))
                {
                    try
                    {
                        // search for a dwg file named 'blockName' in AutoCAD search paths
                        var filename = HostApplicationServices.Current.FindFile("tabrifu1.dwg", acCurDb, FindFileHint.Default);
                        // add the dwg model space as 'blockName' block definition in the current database block table
                        using (var sourceDb = new Database(false, true))
                        {
                            sourceDb.ReadDwgFile(filename, FileOpenMode.OpenForReadAndAllShare, true, "");
                            acCurDb.Insert("tabrifu1", sourceDb, true);
                        }
                    }
                    catch
                    {
                        acDocEd.WriteMessage($"\nBlock tabrifu1.dwg not found.");
                        return;
                    }
                }

                Point3d punktwstawienia;
                punktwstawienia = new Point3d(Pointbazowy.X, Pointbazowy.Y + 1750, 0);

                // create a new block reference
                using (var br = new BlockReference(punktwstawienia, acBlkTbl["AntennenTabelle_Rifu_Kopf"]))
                {
                  
                    var space = (BlockTableRecord)acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite);
                    space.AppendEntity(br);
                    acTrans.AddNewlyCreatedDBObject(br, true);
                     
                }
                punktwstawienia = Pointbazowy;
                using (var br = new BlockReference(punktwstawienia, acBlkTbl["tabrifu1"]))
                {
                    
                    var space = (BlockTableRecord)acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite);
                    space.AppendEntity(br);
                    acTrans.AddNewlyCreatedDBObject(br, true);
                  
                }
                                
                acTrans.Commit();

            }

        }


        public void generuj_wymiarynowej_tabelki()
        {
            //generuje wymiary dla kolumn i wierszy dla tekstów do nowej tabelki
            rifutabelka2.zesatwkolumn[1].X0 = Convert.ToInt32(0);
            rifutabelka2.zesatwkolumn[1].X1 = rifutabelka2.zesatwkolumn[1].X0 + 2500;

            rifutabelka2.zesatwkolumn[2].X0 = rifutabelka2.zesatwkolumn[1].X1;
            rifutabelka2.zesatwkolumn[2].X1 = rifutabelka2.zesatwkolumn[2].X0 + 1500;

            rifutabelka2.zesatwkolumn[3].X0 = rifutabelka2.zesatwkolumn[2].X1;
            rifutabelka2.zesatwkolumn[3].X1 = rifutabelka2.zesatwkolumn[3].X0 + 2000;

            rifutabelka2.zesatwkolumn[4].X0 = rifutabelka2.zesatwkolumn[3].X1;
            rifutabelka2.zesatwkolumn[4].X1 = rifutabelka2.zesatwkolumn[4].X0 + 800;

            rifutabelka2.zesatwkolumn[5].X0 = rifutabelka2.zesatwkolumn[4].X1;
            rifutabelka2.zesatwkolumn[5].X1 = rifutabelka2.zesatwkolumn[5].X0 + 800;

            rifutabelka2.zesatwkolumn[6].X0 = rifutabelka2.zesatwkolumn[5].X1;
            rifutabelka2.zesatwkolumn[6].X1 = rifutabelka2.zesatwkolumn[6].X0 + 1200;

            rifutabelka2.zesatwkolumn[7].X0 = rifutabelka2.zesatwkolumn[6].X1;
            rifutabelka2.zesatwkolumn[7].X1 = rifutabelka2.zesatwkolumn[7].X0 + 1500;

            rifutabelka2.zesatwkolumn[8].X0 = rifutabelka2.zesatwkolumn[7].X1;
            rifutabelka2.zesatwkolumn[8].X1 = rifutabelka2.zesatwkolumn[8].X0 + 1000;

            rifutabelka2.zesatwkolumn[9].X0 = rifutabelka2.zesatwkolumn[8].X1;
            rifutabelka2.zesatwkolumn[9].X1 = rifutabelka2.zesatwkolumn[9].X0 + 1000;

            rifutabelka2.zesatwkolumn[10].X0 = rifutabelka2.zesatwkolumn[9].X1;
            rifutabelka2.zesatwkolumn[10].X1 = rifutabelka2.zesatwkolumn[10].X0 + 2400;

            rifutabelka2.zesatwkolumn[11].X0 = rifutabelka2.zesatwkolumn[10].X1;
            rifutabelka2.zesatwkolumn[11].X1 = rifutabelka2.zesatwkolumn[11].X0 + 2400;

            rifutabelka2.zesatwkolumn[12].X0 = rifutabelka2.zesatwkolumn[11].X1;
            rifutabelka2.zesatwkolumn[12].X1 = rifutabelka2.zesatwkolumn[12].X0 + 1700;

            rifutabelka2.zesatwkolumn[13].X0 = rifutabelka2.zesatwkolumn[12].X1;
            rifutabelka2.zesatwkolumn[13].X1 = rifutabelka2.zesatwkolumn[13].X0 + 1000;

            rifutabelka2.zesatwkolumn[14].X0 = rifutabelka2.zesatwkolumn[13].X1;
            rifutabelka2.zesatwkolumn[14].X1 = rifutabelka2.zesatwkolumn[14].X0 + 1000;

            rifutabelka2.zesatwkolumn[15].X0 = rifutabelka2.zesatwkolumn[14].X1;
            rifutabelka2.zesatwkolumn[15].X1 = rifutabelka2.zesatwkolumn[15].X0 + 1000;

            rifutabelka2.zesatwkolumn[16].X0 = rifutabelka2.zesatwkolumn[15].X1;
            rifutabelka2.zesatwkolumn[16].X1 = rifutabelka2.zesatwkolumn[16].X0 + 1200;

            rifutabelka2.zesatwkolumn[17].X0 = rifutabelka2.zesatwkolumn[16].X1;
            rifutabelka2.zesatwkolumn[17].X1 = rifutabelka2.zesatwkolumn[17].X0 + 2000;

            rifutabelka2.zesatwkolumn[18].X0 = rifutabelka2.zesatwkolumn[17].X1;
            rifutabelka2.zesatwkolumn[18].X1 = rifutabelka2.zesatwkolumn[18].X0 + 1350;

            rifutabelka2.zesatwkolumn[19].X0 = rifutabelka2.zesatwkolumn[18].X1;
            rifutabelka2.zesatwkolumn[19].X1 = rifutabelka2.zesatwkolumn[19].X0 + 1250;

            rifutabelka2.zesatwkolumn[20].X0 = rifutabelka2.zesatwkolumn[19].X1;
            rifutabelka2.zesatwkolumn[20].X1 = rifutabelka2.zesatwkolumn[20].X0 + 1000;

            rifutabelka2.zesatwkolumn[21].X0 = rifutabelka2.zesatwkolumn[20].X1;
            rifutabelka2.zesatwkolumn[21].X1 = rifutabelka2.zesatwkolumn[21].X0 + 1000;

            rifutabelka2.zesatwkolumn[22].X0 = rifutabelka2.zesatwkolumn[21].X1;
            rifutabelka2.zesatwkolumn[22].X1 = rifutabelka2.zesatwkolumn[22].X0 + 1400;

            rifutabelka2.zesatwkolumn[23].X0 = rifutabelka2.zesatwkolumn[22].X1;
            rifutabelka2.zesatwkolumn[23].X1 = rifutabelka2.zesatwkolumn[23].X0 + 1000;

            rifutabelka2.zesatwkolumn[24].X0 = rifutabelka2.zesatwkolumn[23].X1;
            rifutabelka2.zesatwkolumn[24].X1 = rifutabelka2.zesatwkolumn[24].X0 + 1000;

            rifutabelka2.zesatwkolumn[25].X0 = rifutabelka2.zesatwkolumn[24].X1;
            rifutabelka2.zesatwkolumn[25].X1 = rifutabelka2.zesatwkolumn[25].X0 + 1000;

            rifutabelka2.zesatwkolumn[26].X0 = rifutabelka2.zesatwkolumn[25].X1;
            rifutabelka2.zesatwkolumn[26].X1 = rifutabelka2.zesatwkolumn[26].X0 + 1000;

            rifutabelka2.zesatwkolumn[27].X0 = rifutabelka2.zesatwkolumn[26].X1;
            rifutabelka2.zesatwkolumn[27].X1 = rifutabelka2.zesatwkolumn[27].X0 + 800;

            rifutabelka2.zesatwkolumn[28].X0 = rifutabelka2.zesatwkolumn[27].X1;
            rifutabelka2.zesatwkolumn[28].X1 = rifutabelka2.zesatwkolumn[28].X0 + 1000;

            rifutabelka2.zesatwkolumn[29].X0 = rifutabelka2.zesatwkolumn[28].X1;
            rifutabelka2.zesatwkolumn[29].X1 = rifutabelka2.zesatwkolumn[29].X0 + 2700;


            //yyyy

            rifutabelka2.zesatwwierszy[1].Y0 = Convert.ToInt32(0);
            rifutabelka2.zesatwwierszy[1].Y1 = rifutabelka2.zesatwwierszy[1].Y0 - 500;

            for (int k = 2; k < 20; k++)
            {

                rifutabelka2.zesatwwierszy[k].Y0 = rifutabelka2.zesatwwierszy[k - 1].Y1;
                rifutabelka2.zesatwwierszy[k].Y1 = rifutabelka2.zesatwwierszy[k].Y0 - 500;
            }


        }


        public void staratabelazamianaKol_naX()
        {
            //foreach (Textydocad textydocad in textydocads)
            //{
            //   // textydocad.X0 = rifutabelka_old.zesatwkolumn[textydocad.Kolumna].X0;
            //   // textydocad.Y0 = rifutabelka_old.zesatwwierszy[textydocad.Wiersz].Y0;


            //    textydocad.X0 = (rifutabelka_old.zesatwkolumn[textydocad.Kolumna].X1 + rifutabelka_old.zesatwkolumn[textydocad.Kolumna].X0) / 2;
            //    textydocad.SzerTla = rifutabelka_old.zesatwkolumn[textydocad.Kolumna].X1 - rifutabelka_old.zesatwkolumn[textydocad.Kolumna].X0;

            //    textydocad.Y0 = (rifutabelka_old.zesatwwierszy[textydocad.Wiersz].Y0 + rifutabelka_old.zesatwwierszy[textydocad.Wiersz].Y1) / 2;
            //    textydocad.WysTla = (rifutabelka_old.zesatwwierszy[textydocad.Wiersz].Y0 - rifutabelka_old.zesatwwierszy[textydocad.Wiersz].Y1) ;
            //}
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

        public string kierpolnocy_deg = "";
        public string kierpolnocy_rad = "";


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

    public class Tabelka_plan
    {

        public string Lfd_Nr { get; set; }
        public string USER_LINK_ID        { get; set; }
        public string NE_A { get; set; }
        public string Main_Status { get; set; }

        public string Azimuth        { get; set; }
        public string Typ        { get; set; }
        public string Diameter        { get; set; }
        public string Height        { get; set; }
        public string Trager { get; set; }
        public string Frequenz        { get; set; }
        public string Vendor        { get; set; }
        public string Kapazität        { get; set; }
        public string System        { get; set; }
        public string Distance        { get; set; }
        public string Site_B { get; set; }
        
        public string NE_B       { get; set; }

        public Tabelka_plan()
        { }
        public Tabelka_plan(string[] wiersz)
        {
            if (wiersz.Count() >= 19)
            {

                Lfd_Nr = wiersz[0];
                USER_LINK_ID = wiersz[1];
                NE_A = wiersz[2];
                Main_Status = wiersz[3];

                Azimuth = wiersz[5];
                Typ = wiersz[6];
                Diameter = wiersz[7];
                Height = wiersz[8];
                Trager = wiersz[9];

                Frequenz = wiersz[11];
                Vendor = wiersz[12];
                Kapazität = wiersz[13];
                System = wiersz[14];

                Distance = wiersz[16];
                Site_B = wiersz[17];
                NE_B = wiersz[18];
            }

        }
    }


    public class tabelkapokaz
    {

        public string RIFU_NR { get; set; }
        public string NETZ { get; set; }
        public string RIFU { get; set; }
        public string AUFBAU { get; set; }
        public string OPTION { get; set; }
        public string FREQUENZ { get; set; }
        public string Farbe { get; set; }
        public string RICHTUNG { get; set; }
        public string HÖHE { get; set; }
        public string GEGENSTELLE { get; set; }
        public string Linknummer { get; set; }
        public string HOHLLEITER_TYP { get; set; }
        public string HOHLLEITER_ANZAHL { get; set; }
        public string HOHLLEITER_LÄNGE { get; set; }
        public string HOHLLEITER_AUFBAU { get; set; }
        public string HOHLLEITER_OPTION { get; set; }
        public string ODU_TYP { get; set; }
        public string ODU_ANZAHL { get; set; }
        public string DATENKABEL_TYP { get; set; }
        public string DATENKABEL_ANZAHL { get; set; }
        public string DATENKABEL_LÄNGE { get; set; }
        public string POWERKABEL_TYP { get; set; }
        public string POWERKABEL_ANZAHL { get; set; }
        public string POWERKABEL_LÄNGE { get; set; }
        public string EISSCHUTZ { get; set; }
        public string STATI_VERDREHUNG { get; set; }
        public string ANT_TAEGER_NR { get; set; }
        public string ANT_TÄGER_DURCHM { get; set; }
        public string BEMERKUNG { get; set; }

    }

    public class tabelkapokaz20
    {

        public string RIFU_NR { get; set; } = ""; //1
        public string NETZ { get; set; } = "";   //2
        public string RIFU { get; set; } = "";   //3
        public string AUFBAU { get; set; } = ""; //4
        public string OPTION { get; set; } = ""; //5
        public string Farbe { get; set; } = "";
        public string RICHTUNG { get; set; } = "";
        public string HÖHE { get; set; } = "";
        public string DATENKABEL_TYP { get; set; } = "";
        public string DATENKABEL_ANZAHL { get; set; } = "";


        public string DATENKABEL_LÄNGE { get; set; } = "";
        public string Farbe_kable { get; set; } = "";
        public string GEGENSTELLE { get; set; } = "";
        public string Linknummer { get; set; } = "";
        public string HOHLLEITER_LÄNGE { get; set; } = "";
        public string ODU_TYP { get; set; } = "";
        public string ODU_ANZAHL { get; set; } = "";
        public string ANT_TAEGER_NR { get; set; } = "";
        public string ANT_TÄGER_DURCHM { get; set; } = "";
        public string BEMERKUNG { get; set; } = "";

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


    class ExcelAll
    {


        private Excel.Application m_objExcel = null;
        private Excel.Workbooks m_objBooks = null;
        private Excel._Workbook m_objBook = null;
        private Excel.Sheets m_objSheets = null;
        //private Excel._Worksheet m_objSheet = null;
        //private Excel._Worksheet m_objSheet1 = null;
        //private Excel._Worksheet m_objSheet2 = null;
        private Excel._Worksheet m_objSheet3 = null;
        //private Excel._Worksheet m_objSheet4 = null;
        //private Excel._Worksheet m_objSheet5 = null;
        //private Excel._Worksheet m_objSheet6 = null;

        private Excel.Range m_objRange = null;
        private Excel.Font m_objFont = null;
        // private Excel.QueryTables m_objQryTables = null;
        // private Excel._QueryTable m_objQryTable = null;
        private object m_objOpt = System.Reflection.Missing.Value;
        // private object m_strSampleFolder = "C:\\ExcelData\\";
        string mySheet = @"‪‪‪‪K:\tmp\hkt.xlsx";

        public int ile_wierszy_w_cad = 0;
        public int ile_kolumn_w_cad = 0;

        public List<Textydocad> textydocads = new List<Textydocad>();
        public Tabelka TabelkazExcel = new Tabelka();
        public string[,] napisy_z_excel = new string[20, 30];

        public ExcelAll()
        { }

        public void Excel_close()
        {
            m_objBooks.Close();
        }

        public void OpentemplateExcel()
        {
            m_objExcel = new Excel.Application();
            m_objExcel.Visible = true;
            m_objBooks = (Excel.Workbooks)m_objExcel.Workbooks;



            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            if (openFileDialog.ShowDialog() == true)

                mySheet = openFileDialog.FileName;




            m_objBook = (Excel._Workbook)(m_objBooks.Open(mySheet));
            m_objSheets = (Excel.Sheets)m_objBook.Worksheets;

            //m_objSheet = (Excel._Worksheet)(m_objSheets.get_Item("z ACAD new 29"));
            //m_objSheet1 = (Excel._Worksheet)(m_objSheets.get_Item("z ACAD old 20"));
            //m_objSheet2 = (Excel._Worksheet)(m_objSheets.get_Item("pomocnicze1"));
            m_objSheet3 = (Excel._Worksheet)(m_objSheets.get_Item("Gotowa do ACAD new 29"));
            //m_objSheet4 = (Excel._Worksheet)(m_objSheets.get_Item("Gotowa do ACAD old 20"));
            //m_objSheet5 = (Excel._Worksheet)(m_objSheets.get_Item("Gotowa do ACAD old 20"));
            //m_objSheet6 = (Excel._Worksheet)(m_objSheets.get_Item("Gotowa do ACAD old 20"));
        }




        public void zexcelodczytdanych()
        //odczytuje poprawione dane do listy textów i uzupełnia o nr wiersza i nr kolumny 
        {
            //mozan stworzyć nową tabelkę do naszych potrzeb

            //wier1 wiersz
            // kol1 kolumna
            string text3;
            textydocads.Clear();
            int ilosckolumn = 29;
            ile_kolumn_w_cad = 29;
            m_objRange = m_objSheet3.Range["A5", "AC25"];

            for (int wier1 = 1; wier1 < 20; wier1++)

            {
                if (m_objRange.Cells[wier1, 1].Value != null)

                {
                    for (int kol1 = 1; kol1 <= ilosckolumn; kol1++)
                    {
                        m_objFont = m_objRange.Cells[wier1, kol1].Font;

                        if (m_objRange.Cells[wier1, kol1].Value == null)
                        { text3 = ""; }
                        else { text3 = m_objRange.Cells[wier1, kol1].Value.ToString(); }

                        string kolor = m_objRange.Cells[wier1, kol1].Font.Color.ToString();
                        string backolor = m_objRange.Cells[wier1, kol1].Interior.Color.ToString();

                        //  Textydocad(string text, int wiersz, int kolumna, string kolor, string kolortla)
                        textydocads.Add(new Textydocad(text3, wier1, kol1, kolor, backolor));
                        napisy_z_excel[wier1, kol1] = text3;
                    }
                    //dodanie znaczku st do kol 8
                    napisy_z_excel[wier1, 8] = napisy_z_excel[wier1, 8] + "%%d";
                }
                else
                {
                    ile_wierszy_w_cad = wier1 - 1;

                    break;
                }
            }

            Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("Oczytano " + textydocads.Count.ToString() + " napisow");
        }

    }







}







