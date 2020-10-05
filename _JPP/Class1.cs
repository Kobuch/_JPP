
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;


namespace JPP
{
    public class Class1
    {
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

        [CommandMethod("JPP_HKT_schemat")]
        public void JPP_HKT_schemat()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();
            hKT.KHT_schemat();

        }


        [CommandMethod("JPP_HKT_schemat_rys")]
        public void JPP_HKT_schemat_rys()
        {
            _JPP.HKT_class hKT = new _JPP.HKT_class();

            hKT.rysuj_schemat_rifu_80(new Point3d(100,100,0));

        }


    }
}
    













