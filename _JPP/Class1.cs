
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System.Collections.Generic;

namespace JPP
{
    public class Class1
    {

        List<object> lista = new List<object>();

        public Point3d punkt1 { get; set; }
        public Point3d punkt2 { get; set; }
        public Point3d punkt3 { get; set; }
        public Point3d punkt4 { get; set; }
        public Point3d punkt5 { get; set; }



        [CommandMethod("JPP_blok0")]
        public void JPP_blok0()
        {

            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            // Start a transaction
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Get the current document editor
                Editor acDocEd = Application.DocumentManager.MdiActiveDocument.Editor;

                PromptSelectionResult acSSPrompt;
                acSSPrompt = acDoc.Editor.GetSelection();


                if (acSSPrompt.Status == PromptStatus.OK)
                {
                    SelectionSet acSSet = acSSPrompt.Value;

                    // Step through the objects in the selection set
                    foreach (SelectedObject acSSObj in acSSet)
                    {
                        // Check to make sure a valid SelectedObject object was returned
                        if (acSSObj != null)
                        {
                            Entity e = (Entity)acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForWrite);
                            if (e.GetType().Name == "BlockReference")
                            {
                                BlockReference acBlkRef = (BlockReference)acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as BlockReference;
                                zmiananawartwy(acBlkRef, acTrans);
                            }
                        }
                    }
                }
                acTrans.Commit();
                acDocEd.Regen();
            }
        }


        public void zmiananawartwy(BlockReference acBlkRef1, Transaction acTrans1)
        {

            BlockTableRecord acBlkTblRec = (BlockTableRecord)acTrans1.GetObject(acBlkRef1.BlockTableRecord, OpenMode.ForWrite);

            foreach (ObjectId asObjId in acBlkTblRec)
            {
                Entity e = (Entity)acTrans1.GetObject(asObjId, OpenMode.ForWrite);

                e.Layer = "0";

                if (e.GetType().Name == "BlockReference")
                {
                    zmiananawartwy((BlockReference)e, acTrans1);
                }
            }

        }



        /// <summary>
        /// Schemat procedury do rys schematu radiolinii
        /// </summary>
        
        #region  Schemat RIFU



        [CommandMethod("JPP_HKT_schemat")]
        public void JPP_HKT_schemat()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            hKT.rysuj_schemat();
        }

        [CommandMethod("JPP_HKT_schemat1")]
        public void JPP_HKT_schemat1()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            hKT.rysuj_schematpoj(1);
        }

        [CommandMethod("JPP_HKT_schemat2")]
        public void JPP_HKT_schemat2()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            hKT.rysuj_schematpoj(2);
        }

        [CommandMethod("JPP_HKT_schemat3")]
        public void JPP_HKT_schemat3()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            hKT.rysuj_schematpoj(3);
        }

        [CommandMethod("JPP_HKT_schemat4")]
        public void JPP_HKT_schemat4()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            hKT.rysuj_schematpoj(4);
        }

        [CommandMethod("JPP_HKT_schemat5")]
        public void JPP_HKT_schemat5()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            hKT.rysuj_schematpoj(5);
        }
        [CommandMethod("JPP_HKT_schemat6")]
        public void JPP_HKT_schemat6()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            hKT.rysuj_schematpoj(6);
        }
        [CommandMethod("JPP_HKT_schemat7")]
        public void JPP_HKT_schemat7()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            hKT.rysuj_schematpoj(7);
        }
        [CommandMethod("JPP_HKT_schemat8")]
        public void JPP_HKT_schemat8()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            hKT.rysuj_schematpoj(8);
        }
        [CommandMethod("JPP_HKT_schemat9")]
        public void JPP_HKT_schemat9()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            hKT.rysuj_schematpoj(9);
        }

        #endregion


        #region Odczyt danych z Excel

        /// odczyt napisów z tabeli exela - konkretny format
        [CommandMethod("JPP_HKT_RLzexcel")]
        public void JPP_HKT_RLzexcel()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            hKT.HKT_odczyt_z_excel();
        }
        #endregion

        #region Odczyt danych z tabeli w dwg

        
        /// odczyt napisów z tabeli z CAD - konkretny format
        [CommandMethod("JPP_HKT_RLzcad")]
        public void JPP_HKT_RLzcad()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            hKT.JPP_HKT_RLzcad();
        }
        #endregion

        #region Obsługa properties


        [CommandMethod("JPP_HKT_czysc_properties_jpp")]
        public void JPP_HKT_czysc_properties_jpp()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            hKT.HKT_czysc_properties_jpp();
        }
        [CommandMethod("JPP_HKT_sprawdz_properties_jpp")]
        public void JPP_HKT_sprawdz_properties_jpp()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            hKT.HKT_sprawdz_properties_jpp();
        }

       [CommandMethod("JPP_HKT_rysuj_tabelke_z_prop")]
        public void JPP_HKT_rysuj_tabelke_z_prop()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            hKT.Rysuj_tabelke_w_cad_z_properties();
        }

        [CommandMethod("JPP_HKT_tab_zmiana20na29_z_prop")]
        public void JPP_HKT_tab_zmiana20na29_z_prop()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            hKT.tabela_zmiana20na29_z_prop();
        }
  

      [CommandMethod("JPP_HKT_akualizacja_Rifu_planowane")]
        public void JPP_HKT_akualizacja_Rifu_planowane()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            hKT.akualizacja_Rifu_planowane();
        }
        #endregion



        #region widoki rifu 300

      

        [CommandMethod("JPP_HKT_Rifu_300_0_BOK")]
        public void JPP_HKT_Rifu_300_0_BOK()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            hKT.wstaw_rzut_poziomy_Rifu("JPP_300_0_ODU_BOK");
        }

        [CommandMethod("JPP_HKT_Rifu_300_0_PION")]
        public void JPP_HKT_Rifu_300_0_PION()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            hKT.wstaw_rzut_poziomy_Rifu("JPP_300_0_ODU_PION");
        }

        [CommandMethod("JPP_HKT_Rifu_300_0_PRZOD")]
        public void JPP_HKT_Rifu_300_0_PRZOD()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            hKT.wstaw_rzut_poziomy_Rifu("JPP_300_0_ODU_PRZOD");
        }

        [CommandMethod("JPP_HKT_Rifu_300_0_TYL")]
        public void JPP_HKT_Rifu_300_0_TYL()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            hKT.wstaw_rzut_poziomy_Rifu("JPP_300_0_ODU_TYL");
        }

        [CommandMethod("JPP_HKT_Rifu_300_1_BOK")]
        public void JPP_HKT_Rifu_300_1_BOK()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            hKT.wstaw_rzut_poziomy_Rifu("JPP_300_1_ODU_BOK");
        }
        [CommandMethod("JPP_HKT_Rifu_300_1_PION")]
        public void JPP_HKT_Rifu_300_1_PION()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            hKT.wstaw_rzut_poziomy_Rifu("JPP_300_1_ODU_PION");
        }

        [CommandMethod("JPP_HKT_Rifu_300_1_PRZOD")]
        public void JPP_HKT_Rifu_300_1_PRZOD()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            hKT.wstaw_rzut_poziomy_Rifu("JPP_300_1_ODU_PRZOD");
        }

        [CommandMethod("JPP_HKT_Rifu_300_1_TYL")]
        public void JPP_HKT_Rifu_300_1_TYL()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            hKT.wstaw_rzut_poziomy_Rifu("JPP_300_1_ODU_TYL");
        }



        [CommandMethod("JPP_HKT_Rifu_300_2_BOK")]
        public void JPP_HKT_Rifu_300_2_BOK()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            hKT.wstaw_rzut_poziomy_Rifu("JPP_300_2_ODU_BOK");
        }
        [CommandMethod("JPP_HKT_Rifu_300_2_PION")]
        public void JPP_HKT_Rifu_300_2_PION()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            hKT.wstaw_rzut_poziomy_Rifu("JPP_300_2_ODU_PION");
        }

        [CommandMethod("JPP_HKT_Rifu_300_2_PRZOD")]
        public void JPP_HKT_Rifu_300_2_PRZOD()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            hKT.wstaw_rzut_poziomy_Rifu("JPP_300_2_ODU_PRZOD");
        }

        [CommandMethod("JPP_HKT_Rifu_300_2_TYL")]
        public void JPP_HKT_Rifu_300_2_TYL()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            hKT.wstaw_rzut_poziomy_Rifu("JPP_300_2_ODU_TYL");
        }

        #endregion

        #region widoki rifu 600



        [CommandMethod("JPP_HKT_Rifu_600_0_BOK")]
        public void JPP_HKT_Rifu_600_0_BOK()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            hKT.wstaw_rzut_poziomy_Rifu("JPP_600_0_ODU_BOK");
        }

        [CommandMethod("JPP_HKT_Rifu_600_0_PION")]
        public void JPP_HKT_Rifu_600_0_PION()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            hKT.wstaw_rzut_poziomy_Rifu("JPP_600_0_ODU_PION");
        }

        [CommandMethod("JPP_HKT_Rifu_600_0_PRZOD")]
        public void JPP_HKT_Rifu_600_0_PRZOD()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            hKT.wstaw_rzut_poziomy_Rifu("JPP_600_0_ODU_PRZOD");
        }

        [CommandMethod("JPP_HKT_Rifu_600_0_TYL")]
        public void JPP_HKT_Rifu_600_0_TYL()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            hKT.wstaw_rzut_poziomy_Rifu("JPP_600_0_ODU_TYL");
        }

        [CommandMethod("JPP_HKT_Rifu_600_1_BOK")]
        public void JPP_HKT_Rifu_600_1_BOK()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            hKT.wstaw_rzut_poziomy_Rifu("JPP_600_1_ODU_BOK");
        }
        [CommandMethod("JPP_HKT_Rifu_600_1_PION")]
        public void JPP_HKT_Rifu_600_1_PION()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            hKT.wstaw_rzut_poziomy_Rifu("JPP_600_1_ODU_PION");
        }

        [CommandMethod("JPP_HKT_Rifu_600_1_PRZOD")]
        public void JPP_HKT_Rifu_600_1_PRZOD()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            hKT.wstaw_rzut_poziomy_Rifu("JPP_600_1_ODU_PRZOD");
        }

        [CommandMethod("JPP_HKT_Rifu_600_1_TYL")]
        public void JPP_HKT_Rifu_600_1_TYL()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            hKT.wstaw_rzut_poziomy_Rifu("JPP_600_1_ODU_TYL");
        }



        [CommandMethod("JPP_HKT_Rifu_600_2_BOK")]
        public void JPP_HKT_Rifu_600_2_BOK()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            hKT.wstaw_rzut_poziomy_Rifu("JPP_600_2_ODU_BOK");
        }
        [CommandMethod("JPP_HKT_Rifu_600_2_PION")]
        public void JPP_HKT_Rifu_600_2_PION()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            hKT.wstaw_rzut_poziomy_Rifu("JPP_600_2_ODU_PION");
        }

        [CommandMethod("JPP_HKT_Rifu_600_2_PRZOD")]
        public void JPP_HKT_Rifu_600_2_PRZOD()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            hKT.wstaw_rzut_poziomy_Rifu("JPP_600_2_ODU_PRZOD");
        }

        [CommandMethod("JPP_HKT_Rifu_600_2_TYL")]
        public void JPP_HKT_Rifu_600_2_TYL()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            hKT.wstaw_rzut_poziomy_Rifu("JPP_600_2_ODU_TYL");
        }

        #endregion





        #region Pobranie dodatkowych danych PN i tp

        [CommandMethod("JPP_HKT_roza_wiatrow_Rifu")]
        public void JPP_HKT_roza_wiatrow_Rifu()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            hKT.generuj_rzut_rozy_wiatrow();
        }
       

        [CommandMethod("JPP_HKT_pobierz_kierunek_polnocy")]
        public void pobierz_kierunek_polnocy()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            hKT.pobierz_kierunek_polnocy();
        }
        #endregion


        #region   Obsługa rzurow radiolinii HKT
        
        [CommandMethod("JPP_HKT_rzuty_radiolinii_all")]
        public void JPP_HKT_rzuty_radiolinii_all()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            hKT.generuj_rzut_wszystkich_radiolinii();
        }

        [CommandMethod("JPP_HKT_rzuty_radiolinii1")]
        public void JPP_HKT_rzuty_radiolinii1()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            lista = new List<object>() { ((int)1)};
            hKT.generuj_rzut_1radiolinii(lista);
        }


        [CommandMethod("JPP_HKT_rzuty_radiolinii2")]
        public void JPP_HKT_rzuty_radiolinii2()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            lista = new List<object>() { ((int)2) };
            hKT.generuj_rzut_1radiolinii(lista);
        }

        [CommandMethod("JPP_HKT_rzuty_radiolinii3")]
        public void JPP_HKT_rzuty_radiolinii3()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            lista = new List<object>() { ((int)3) };
            hKT.generuj_rzut_1radiolinii(lista);
        }

        [CommandMethod("JPP_HKT_rzuty_radiolinii4")]
        public void JPP_HKT_rzuty_radiolinii4()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            lista = new List<object>() { ((int)4) };
            hKT.generuj_rzut_1radiolinii(lista);
        }


        [CommandMethod("JPP_HKT_rzuty_radiolinii5")]
        public void JPP_HKT_rzuty_radiolinii5()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            lista = new List<object>() { ((int)5) };
            hKT.generuj_rzut_1radiolinii(lista);
        }

        [CommandMethod("JPP_HKT_rzuty_radiolinii6")]
        public void JPP_HKT_rzuty_radiolinii6()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            lista = new List<object>() { ((int)6) };
            hKT.generuj_rzut_1radiolinii(lista);
        }

        [CommandMethod("JPP_HKT_rzuty_radiolinii7")]
        public void JPP_HKT_rzuty_radiolinii7()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            lista = new List<object>() { ((int)7) };
            hKT.generuj_rzut_1radiolinii(lista);
        }

        [CommandMethod("JPP_HKT_rzuty_radiolinii8")]
        public void JPP_HKT_rzuty_radiolinii8()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            lista = new List<object>() { ((int)8) };
            hKT.generuj_rzut_1radiolinii(lista);
        }
        [CommandMethod("JPP_HKT_rzuty_radiolinii9")]
        public void JPP_HKT_rzuty_radiolinii9()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            lista = new List<object>() { ((int)9) };
            hKT.generuj_rzut_1radiolinii(lista);
        }

        #endregion





     


  

    }
}
    













