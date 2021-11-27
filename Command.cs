#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.IO;
using System.Linq;

#endregion

namespace BergmannTestTask
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        #region Constants
        /// <summary>
        /// Family template filename
        /// </summary>
        const string FAMILY_TEMPLATE = "Generic Model.rft";

        /// <summary>
        /// Family filename
        /// </summary>
        const string FAMILY_NAME = "Cube";

        /// <summary>
        /// Family filename ext
        /// </summary>
        const string FAMILY_NAME_EXT = ".rfa";

        /// <summary>
        /// Plugin Directory
        /// </summary>
        readonly string CURRENT_DIR = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) +
                        @"\Autodesk\REVIT\Addins\2022\";

        #endregion
        private Autodesk.Revit.Creation.FamilyItemFactory familyCreation = null;
        private Autodesk.Revit.DB.Document familyDocument;
        private int errCount = 0;
        private float w = 20;


        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            if (null == doc)
            {
                message = "Please run this command in an open document.";
                return Result.Failed;
            }

            try
            {
                app = commandData.Application.Application;

                if (!doc.IsFamilyDocument)
                {
                    familyDocument = app.NewFamilyDocument(
                        Path.Combine(CURRENT_DIR + FAMILY_TEMPLATE));

                    if (null == familyDocument)
                    {
                        message = "Cannot open family document";
                        return Result.Failed;
                    }
                }
                familyCreation = familyDocument.FamilyCreate;

                CreateGenericModel(familyDocument, w, doc);
                if (0 == errCount)
                {
                    Transaction transaction = new Transaction(doc,
                        "Insert family instance");

                    transaction.Start();

                    Family family = null;
                    string filename = Path.Combine(
                        CURRENT_DIR, FAMILY_NAME + FAMILY_NAME_EXT);

                    FilteredElementCollector families
                      = new FilteredElementCollector(doc)
                        .OfClass(typeof(Family));

                    int n = families.Count<Element>(
                      e => e.Name.Equals(FAMILY_NAME));

                    if (0 < n)
                    {
                        family = families.First<Element>(
                          e => e.Name.Equals(FAMILY_NAME))
                            as Family;
                    }
                    else
                    {
                        doc.LoadFamily(filename, out family);
                    }

                    FamilySymbol familySymbol = new FilteredElementCollector(doc)
                          .OfClass(typeof(FamilySymbol))
                          .OfType<FamilySymbol>()
                          .Where(x => x.FamilyName == FAMILY_NAME)
                          .FirstOrDefault();

                    if (!familySymbol.IsActive)
                        familySymbol.Activate();

                    XYZ point = uidoc.Selection.PickPoint(
                      "Please pick a point for family instance insertion");

                    StructuralType structuralType = StructuralType.UnknownFraming;

                   FamilyInstance familyInstance = doc.Create.NewFamilyInstance(point, familySymbol, structuralType);
                
                    transaction.Commit();

                    return Result.Succeeded;
                }
                else
                {
                    return Result.Failed;
                }
            }
            catch (Exception e)
            {
                message = e.ToString();
                return Result.Failed;
            }
        }


        public void CreateGenericModel(Document familyDocument, float w, Document doc)
        {
            Transaction transaction = new Transaction(familyDocument, "CreateGenericModel");
            transaction.Start();
            CreateExtrusion(familyDocument, w, doc);

            transaction.Commit();

            string filename = Path.Combine(
              CURRENT_DIR, FAMILY_NAME + FAMILY_NAME_EXT);

            SaveAsOptions opt = new SaveAsOptions();
            opt.OverwriteExistingFile = true;

            familyDocument.SaveAs(filename, opt);

            familyDocument.Close(false);

            return;
        }

        /// <summary>
        /// Create extrusion
        /// </summary>
        private void CreateExtrusion(Document familyDocument, float w, Document doc)
        {
            try
            {
                #region Create rectangle profile
                CurveArrArray curveArrArray = new CurveArrArray();
                CurveArray curveArray1 = new CurveArray();

                double halfWidth = w / 2;

                XYZ normal = XYZ.BasisZ;
                SketchPlane sketchPlane = CreateSketchPlane(normal, XYZ.Zero, familyDocument);

                XYZ center = XYZ.Zero;

                XYZ p0 = new XYZ(halfWidth, halfWidth, 0);
                XYZ p1 = new XYZ(-halfWidth, halfWidth, 0);
                XYZ p2 = new XYZ(-halfWidth, -halfWidth, 0);
                XYZ p3 = new XYZ(halfWidth, -halfWidth, 0);
                Line line1 = Line.CreateBound(p0, p1);                
                Line line2 = Line.CreateBound(p1, p2);
                Line line3 = Line.CreateBound(p2, p3);                
                Line line4 = Line.CreateBound(p3, p0);

                curveArray1.Append(line1);
                curveArray1.Append(line2);
                curveArray1.Append(line3);
                curveArray1.Append(line4);

                curveArrArray.Append(curveArray1);
                #endregion

                Extrusion rectExtrusion = familyCreation.NewExtrusion(true, curveArrArray, sketchPlane, w);

                FamilyManager familyManager = familyDocument.FamilyManager;
                FamilyParameter param = familyManager.AddParameter("width",
                    GroupTypeId.Constraints,
                    SpecTypeId.Length, false);
                //dimR.FamilyLabel = param;

                //Dimensions
                int gap = 2;
                Dimension dimL = CreateDimension(p1, p0, p3, p2, gap * 2, doc);
                Dimension dimLu = CreateDimension(p1, p0,
                    center.Add(line1.Direction.Normalize().Negate().Multiply(line1.Length / 2)),
                    center.Add(line1.Direction.Normalize().Multiply(line1.Length / 2)),
                    gap, doc);
                Dimension dimLb = CreateDimension(
                    center.Add(line1.Direction.Normalize().Multiply(line1.Length / 2)),
                    center.Add(line1.Direction.Normalize().Negate().Multiply(line1.Length / 2)),
                    p3, p2,
                    gap, doc);

                Dimension dimB = CreateDimension(p2, p1, p0, p3, gap * 2, doc);
                Dimension dimBl = CreateDimension(p2, p1,
                   center.Add(line2.Direction.Normalize().Negate().Multiply(line2.Length / 2)),
                   center.Add(line2.Direction.Normalize().Multiply(line2.Length / 2)),
                   gap, doc);
                Dimension dimBr = CreateDimension(
                    center.Add(line4.Direction.Normalize().Negate().Multiply(line4.Length / 2)),
                    center.Add(line4.Direction.Normalize().Multiply(line4.Length / 2)),
                    p0, p3,
                    gap, doc);

            }
            catch (Exception e)
            {
                errCount++;
                TaskDialog mainDialog = new TaskDialog("Œ¯Ë·Í‡!");

                mainDialog.MainContent = "Unexpected exceptions occur in CreateExtrusion: " + e.ToString();
                mainDialog.Show();
            }
        }

        /// <summary>
        /// Create sketch plane for generic model profile
        /// </summary>
        /// <param name="normal">plane normal</param>
        /// <param name="origin">origin point</param>
        /// <param name="doc">Document</param>
        /// <returns></returns>
        internal SketchPlane CreateSketchPlane(XYZ normal, XYZ origin, Document familyDocument)
        {
            Plane geometryPlane = Plane.CreateByNormalAndOrigin(normal, origin);
            if (null == geometryPlane)
            {
                throw new Exception("Create the geometry plane failed.");
            }

            SketchPlane plane = SketchPlane.Create(familyDocument, geometryPlane);
            if (null == plane)
            {
                errCount++;

                throw new Exception("Create the sketch plane failed.");
            }
            return plane;
        }

        internal Dimension CreateDimension(XYZ p0, XYZ p1, XYZ p2, XYZ p3, int gap, Document doc)
        {
            Plane plane = Plane.CreateByThreePoints(p0, p1, p2);
            SketchPlane skplane = SketchPlane.Create(familyDocument, plane);
            Line line1 = Line.CreateBound(p0, p1);
            ModelCurve modelcurve1 = familyDocument.FamilyCreate.NewModelCurve(line1, skplane);
            plane = Plane.CreateByThreePoints(p0, p1, p2);
            skplane = SketchPlane.Create(familyDocument, plane);
            Line line3 = Line.CreateBound(p2, p3);
            ModelCurve modelcurve2 = familyDocument.FamilyCreate.NewModelCurve(line3, skplane);

            ReferenceArray ra = new ReferenceArray();
            ra.Append(modelcurve1.GeometryCurve.Reference);
            ra.Append(modelcurve2.GeometryCurve.Reference);

            XYZ direction = line1.Direction.Negate();

            Line line = Line.CreateBound(p0.Add(direction.Multiply(gap)), p3.Add(direction.Multiply(gap)));

            View activeView = doc.ActiveView;

            Dimension dim = familyDocument.FamilyCreate.NewLinearDimension(activeView, line, ra);

            return dim;
        }
    }
}
