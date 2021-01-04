using System;
using System.Linq;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using JPMorrow.Tools.Diagnostics;
using JPMorrow.Revit.Documents;
using System.Reflection;
using System.IO;
using System.Collections.Generic;

namespace MainApp
{
	/// <summary>
	/// Main Execution
	/// </summary>
	[Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
	[Autodesk.Revit.DB.Macros.AddInId("9BBF529B-520A-4877-B63B-BEF1238B6A05")]
    public partial class ThisApplication : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
			string[] dataDirectories = new string[0];
			bool debugApp = false;

			//set revit model info
			ModelInfo revit_info = ModelInfo.StoreDocuments(commandData, dataDirectories, debugApp);

			//create data directories
			//RAP.GenAppStorageStruct(Settings_Base_Path, Data_Dirs, Directory.CreateDirectory, Directory.Exists);

			View view_to_crop = revit_info.DOC.ActiveView;
			ElementId[] selected_ids = revit_info.UIDOC.Selection.GetElementIds().ToArray();

			//pre-debug
			if(view_to_crop.ViewType != ViewType.FloorPlan)
			{
				debugger.show(err:"This is not a floorplan view. please run the script on a floorplan.");
				return Result.Succeeded;
			}

			if(!view_to_crop.CropBoxActive)
			{
				debugger.show(err:"Crop box for this view is not active please activate it.");
				return Result.Succeeded;
			}

			if(!selected_ids.Any())
			{
				debugger.show(err:"No elements selected. Please select some elements that you want to crop around.");
				return Result.Succeeded;
			}


			//get points from selected elements
			List<XYZ> points = new List<XYZ>();
			foreach(var id in selected_ids)
			{
				BoundingBoxXYZ bb = revit_info.DOC.GetElement(id).get_BoundingBox(view_to_crop);
				points.Add(bb.Min);
				points.Add(bb.Max);
			}

			//make bounding box that encompases them
			BoundingBoxXYZ sb = GetExtents(revit_info, points);

			//make curve that cropbox can accept
			List<Curve> nl = new List<Curve>();
			XYZ top_left = new XYZ(sb.Min.X, sb.Min.Y, 0);
			XYZ top_right = new XYZ(sb.Max.X, sb.Min.Y, 0);
			XYZ bottom_left = new XYZ(sb.Min.X, sb.Max.Y, 0);
			XYZ bottom_right = new XYZ(sb.Max.X, sb.Max.Y, 0);

			nl.Add(Line.CreateBound(top_left, top_right));
			nl.Add(Line.CreateBound(top_right, bottom_right));
			nl.Add(Line.CreateBound(bottom_right, bottom_left));
			nl.Add(Line.CreateBound(bottom_left, top_left));


			CurveLoop cl = CurveLoop.Create(nl);
			Plane plane = cl.GetPlane();

			//check if it is a valid curve
			bool valid = cl.IsRectangular(plane);
			if(!valid)
			{
				debugger.show(err:"The generated rectangle is not rectangular. My bad. Report it to the angles department...");
				return Result.Succeeded;
			}

			//crop view
			using(Transaction tx = new Transaction(revit_info.DOC, "fix crop region"))
			{
				tx.Start();

				ViewCropRegionShapeManager vpcr = view_to_crop.GetCropRegionShapeManager();
				bool cropValid = vpcr.IsCropRegionShapeValid(cl);
				if (cropValid)
					vpcr.SetCropShape(cl);

				tx.Commit();
			}

			return Result.Succeeded;
        }

		/// <summary>
		/// Return a bounding box enclosing all model
		/// elements using only quick filters.
		/// </summary>
		private BoundingBoxXYZ GetExtents(ModelInfo info, IEnumerable<XYZ> points)
		{
			BoundingBoxXYZ sb= new BoundingBoxXYZ();
			sb.Min=new XYZ(points.Min(p=> p.X),
							points.Min(p=> p.Y),
							points.Min(p=> p.Z));
			sb.Max=new XYZ(points.Max(p=> p.X),
							points.Max(p=> p.Y),
							points.Max(p=> p.Z));
			return sb;
		}

		public static void DrawModelLines(ModelInfo info, List<XYZ> pts)
		{
			using (Transaction tx = new Transaction(info.DOC, "placing model line"))
			{
				tx.Start();
				List<XYZ> pts_queue = new List<XYZ>(pts);
				while(pts_queue.Count > 0)
				{
					XYZ[] current_pts = pts_queue.Take(2).ToArray();
					foreach(var pt in current_pts)
						pts_queue.Remove(pt);

					string line_str_style = "<Hidden>"; // system linestyle guaranteed to exist
					Create3DModelLine(info, current_pts[0], current_pts[1], line_str_style);
				}

				tx.Commit();
			}
		}

		public static SketchPlane NewSketchPlanePassLine(ModelInfo info, Line line)
		{
			XYZ p = line.GetEndPoint(0);
			XYZ q = line.GetEndPoint(1);
			XYZ norm = new XYZ(-10000, -10000, 0);
			Plane plane = Plane.CreateByThreePoints(p, q, norm);
			SketchPlane skPlane = SketchPlane.Create(info.DOC, plane);
			return skPlane;
		}

		public static void Create3DModelLine(ModelInfo info, XYZ p, XYZ q, string line_style)
		{
			try
			{
				if (p.IsAlmostEqualTo(q))
				{
					debugger.show(err: "Expected two different points.");
					return;
				}
				Line line = Line.CreateBound(p, q);
				if (null == line)
				{
					debugger.show(err: "Geometry line creation failed.");
					return;
				}
				ModelCurve model_line_curve = null;
				model_line_curve = info.DOC.Create.NewModelCurve(line, NewSketchPlanePassLine(info, line));

				// set linestyle
				ICollection<ElementId> styles = model_line_curve.GetLineStyleIds();
				foreach (ElementId eid in styles)
				{
					Element e = info.DOC.GetElement(eid);
					if (e.Name == line_style)
					{
						model_line_curve.LineStyle = e;
						break;
					}
				}
			}
			catch (Exception ex)
			{
				debugger.show(err:ex.ToString());
			}
		}
    }

	public static class EXT
	{
		/// <summary>
		/// Expand the given bounding box to include
		/// and contain the given point.
		/// </summary>
		public static void ExpandToContain_BTP(
		  this BoundingBoxXYZ bb,
		  XYZ p)
		{
			bb.Min = new XYZ(Math.Min(bb.Min.X, p.X),
			  Math.Min(bb.Min.Y, p.Y),
			  Math.Min(bb.Min.Z, p.Z));

			bb.Max = new XYZ(Math.Max(bb.Max.X, p.X),
			  Math.Max(bb.Max.Y, p.Y),
			  Math.Max(bb.Max.Z, p.Z));
		}
	}
}