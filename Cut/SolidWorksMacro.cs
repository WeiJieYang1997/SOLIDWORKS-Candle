using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;

using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;


namespace Cut
{
    public partial class SolidWorksMacro
    {
        public void Main()
        {
            //ADD THESE LINES OF CODE
            double holeRadius;
            double holeDepth;
            //Create an instance of the user form

            frmCutExtrude myForm = new frmCutExtrude();

            //Set the title for the form
            myForm.Text = "Size of Cut-Extrude in Millimeters";
            //Display the user form and retrieve radius and 
            //depth values typed by the user; divide those values
            //by 1000 to change millimeters to meters
            myForm.ShowDialog();
            holeRadius = myForm.radius / 1000;
            holeDepth = myForm.depth / 1000;
            //Dispose of the user form and remove it from
            //memory because it's no longer needed
            myForm.Dispose();


            ModelDoc2 swDoc = null;
            
            //bool boolstatus = false;
            //var COSMOSWORKSObj = null;
            swDoc = ((ModelDoc2)(swApp.ActiveDoc));
            //COSMOSWORKSObj COSMOSWORKSObj = null;
            //CWAddinCallBackObj CWAddinCallBackObj = null;
            //CWAddinCallBackObj = swApp.GetAddInObject("CosmosWorks.CosmosWorks");
            //COSMOSWORKSObj = CWAddinCallBackObj.COSMOSWORKS;
            ModelView myModelView = null;
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.FrameState = ((int)(swWindowState_e.swWindowMaximized));
            //boolstatus = swDoc.Extension.SelectByRay(0.00053592862681739462, 0.2199999999999136, -0.0045118054006820785, -0.18724816575405859, -0.66456615869841951, -0.723387824845406, 0.0010862903128070283, 2, false, 0, 0);
            //ADD THESE LINES OF CODE 
            //Get coordinates of selection point
            SelectionMgr swSelectionMgr = null;
            swSelectionMgr = (SelectionMgr)swDoc.SelectionManager;
            double[] SelectCoordinates;
            SelectCoordinates = (double[])swSelectionMgr.GetSelectionPoint2(1, -1);
            //If face is selected, then open a sketch;
            //otherwise, stop execution
            object SelectedObject = null;
            SelectedObject = (object)swSelectionMgr.GetSelectedObject6(1, 0);
            int objtype;
            objtype = (int)swSelectionMgr.GetSelectedObjectType3(1, -1);
            if (objtype == (int)swSelectType_e.swSelFACES)
            {
                swDoc.SketchManager.InsertSketch(true);
            }

            //swDoc.ClearSelection2(true);
            SketchSegment skSegment = null;
            skSegment = ((SketchSegment)(swDoc.SketchManager.CreateCircle(0, 0, 0, 0.010505, -0.007509, 0)));
            Feature myFeature = null;
            myFeature = ((Feature)(swDoc.FeatureManager.FeatureCut4(true, false, false, 0, 0, 0.02, 0.01, true, false, false, false, 0.26179938779914946, 0.017453292519943334, false, false, false, false, false, true, true, true, true, false, 0, 0, false, false)));
            swDoc.ISelectionManager.EnableContourSelection = false;
            //StudyManagerObj = null;
            //ActiveDocObj = null;
            //CWAddinCallBackObj = null;
            //COSMOSWORKSObj = null;


            //ADD THESE LINES OF CODE 
            //Get IMathPoint to use when transforming
            //from model space to sketch space
            MathUtility swMathUtility = null;
            MathPoint swMathPoint = null;
            Sketch swSketch = null;
            double dx;
            double dy;
            double dz;
            swMathUtility = (MathUtility)swApp.GetMathUtility();
            swMathPoint = (MathPoint)swMathUtility.CreatePoint(SelectCoordinates);
            //Get reference to sketch
            swSketch = (Sketch)swDoc.SketchManager.ActiveSketch;
            //Translate sketch point into sketch space
            MathTransform swMathTransform = null;
            swMathTransform = (MathTransform)swSketch.ModelToSketchTransform;
            swMathPoint = (MathPoint)swMathPoint.MultiplyTransform(swMathTransform);
            //Retrieve coordinates of the sketch point
            double[] darray;
            darray = (double[])swMathPoint.ArrayData;
            dx = darray[0];
            dy = darray[1];
            dz = darray[2];
            //Use swDoc.SketchManager.CreateCircleByRadius instead of
            //swDoc.SketchManager.CreateCircle because
            //swDoc.SketchManager.CreateCircleByRadius sketches a
            //circle centered on a sketch point and lets you 
            //specify a radius
            double radius = 0.015;
            SketchSegment swSketchSegment = null;
            swSketchSegment = (SketchSegment)swDoc.SketchManager.CreateCircleByRadius(dx, dy, dz, radius);
            //Create the cut extrude feature
            Feature swFeature = null;
            swFeature = (Feature)swDoc.FeatureManager.FeatureCut3(true, false, false, 0, 0, 0.025, 0.01, true, false, false, false, 0, 0, false, false, false, false, false, true, true, false, false, false, (int)swStartConditions_e.swStartSketchPlane, 0, false);
        }


        // The SldWorks swApp variable is pre-assigned for you.
        public SldWorks swApp;

    }
}

