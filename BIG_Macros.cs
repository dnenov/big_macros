/*
 * Created by SharpDevelop.
 * User: dene
 * Date: 7/21/2016
 * Time: 10:45 AM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace BIG_Macros
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.DB.Macros.AddInId("22A30AF4-92FD-478E-A202-C1EEBF522AF9")]
	public partial class ThisApplication
	{
      
		private void Module_Startup(object sender, EventArgs e)
		{	
			
		}

		private void Module_Shutdown(object sender, EventArgs e)
		{			
			
		}
		
		public class LineSelectionFilter : ISelectionFilter
		{
			Document doc = null;
			public LineSelectionFilter(Document document)
			{
				doc = document;
			}
			
			public bool AllowElement(Element element)
			{
				if(element.Category.Name == "Lines" ||
				  element.Category.Name == "<Room Separation>" ||
				 element.Category.Name == "<Area Boundary>")
				{
					return true;
				}
				return false;
			}
			
			public bool AllowReference(Reference refer, XYZ point)
			{
				return true;
			}
		}
		
		public class AreaBoundarySelectionFilter : ISelectionFilter
		{
			Document doc = null;
			public AreaBoundarySelectionFilter(Document document)
			{
				doc = document;
			}
			
			public bool AllowElement(Element element)
			{
				if(element.Category.Name == "<Area Boundary>")
				{
					return true;
				}
				return false;
			}
			
			public bool AllowReference(Reference refer, XYZ point)
			{
				return true;
			}
		}
		
		public class ViewPortSelectionFilter : ISelectionFilter
		{
			Document doc = null;
			public ViewPortSelectionFilter(Document document)
			{
				doc = document;
			}
			
			public bool AllowElement(Element element)
			{
				if(element.Category.Name == "Viewports")
				{
					return true;
				}
				return false;
			}
			
			public bool AllowReference(Reference refer, XYZ point)
			{
				return true;
			}
		}
		
		#region Revit Macros generated code
		private void InternalStartup()
		{
			this.Startup += new System.EventHandler(Module_Startup);
			this.Shutdown += new System.EventHandler(Module_Shutdown);
		}
		#endregion
			
			
#region Macros

		public void OpenViewFromViewTemplate()
		{
			UIDocument uidoc = this.ActiveUIDocument;
			Document doc = uidoc.Document;
			
			var tempalteId = new FilteredElementCollector(doc)
				.OfClass(typeof(View))
				.Cast<View>()
				.Where(x => x.IsTemplate)
				.First(x => x.Name.Equals("BIG_Section 1_50"))
				.Id;
			
			var views = new FilteredElementCollector(doc)
				.OfClass(typeof(View))
				.Cast<View>()
				.Where(x => x.ViewTemplateId.IntegerValue == tempalteId.IntegerValue)
				.ToList();
			
			foreach(var v in views)
			{
				uidoc.ActiveView = v;
			}
		}
		public void DuplicateNumberValues()
		{
			Document doc = this.ActiveUIDocument.Document;
			Selection selection = this.ActiveUIDocument.Selection;
			List<ElementId> ids = selection.GetElementIds().ToList();
			
			List<int> numbers = new FilteredElementCollector(doc)
				.OfCategoryId(doc.GetElement(ids.First()).Category.Id)
				.WhereElementIsNotElementType()
				.Select(x => int.Parse(x.LookupParameter("Number").AsString()))
				.ToList();
			
			List<int> result = Enumerable.Range(0, numbers.Max()).Except(numbers).ToList();
			
			using(Transaction t = new Transaction(doc, "Duplicating Number Elements"))
			{
				t.Start();				
				foreach(ElementId id in ids)
				{
					int number = 0;
					Element el = doc.GetElement(id);
					Parameter param = el.LookupParameter("Number");
					
					Int32.TryParse(param.AsString(), out number);
					if(number != 0)
					{
						number = result[0];
						param.Set(number.ToString());
						result.RemoveAt(0);
					}
				}
				t.Commit();
			}
			
		}
		public void ChangeTypeNames()
		{
			Document doc = this.ActiveUIDocument.Document;
						
            FamilyManager familyManager = doc.FamilyManager;
            FamilyTypeSet familyTypes = familyManager.Types;
	        FamilyTypeSetIterator familyTypesItor = familyTypes.ForwardIterator();
	        familyTypesItor.Reset();
	        
	        string types = "";
	        
	        
			
	        using(Transaction t = new Transaction(doc, "Change type name"))
	        {
	        	t.Start();
	        	while (familyTypesItor.MoveNext())
		        {
		            FamilyType familyType = familyTypesItor.Current as FamilyType;
		            if(familyType.Name.Contains("Plant"))
		            {
		            	familyManager.CurrentType = familyType;
		            	familyManager.RenameCurrentType(familyType.Name.Replace("Plant", "Plant "));
		            }
		        }
	        	t.Commit();
	        }
			TaskDialog.Show("test", types);			
		}
		public void SetWorksets()
		{
			Document doc = this.ActiveUIDocument.Document;
			
			// Select all elements of the correct category
			var elements = new FilteredElementCollector(doc)
				.OfCategory(BuiltInCategory.OST_RoomSeparationLines)
				.ToElements();
			
			// Select Workset integer value
			var workset = new FilteredWorksetCollector(doc)
				.OfKind(WorksetKind.UserWorkset)
				.Where(x => x.Name.Equals("X - Admin"))
				.Select(w => w.Id.IntegerValue)
				.First();
			
			using(Transaction t = new Transaction(doc, "Set Worksets"))
			{
				t.Start();
				foreach(Element el in elements)
				{
					el.LookupParameter("Workset").Set(workset);
				}
				t.Commit();
			}
		}
		public void OffAxis()
		{
            Document doc = this.ActiveUIDocument.Document;			
			Selection selection = this.ActiveUIDocument.Selection;
			
//			List<ElementId> ids = selection.PickObjects(ObjectType.Element, "Pick Axis to fix").ToList().Select(x => x.ElementId).ToList();
			List<ElementId> ids = selection.GetElementIds().ToList();
			foreach(ElementId id in ids)
			{
				Fix(id);
			}
		}
		
	    private void Fix(ElementId id)
        {            
            Document doc = this.ActiveUIDocument.Document;
                        
//			  Curve locationLine = (doc.GetElement(id) as Grid).Curve;
                        
            ModelLine modelLine = doc.GetElement(id) as ModelLine;         
            Curve locationLine = modelLine.GeometryCurve;

//            Wall wall = doc.GetElement(id) as Wall;
//            Curve locationLine = (wall.Location as LocationCurve).Curve;
			/*
            // mini procedure to make sure rotation a line doesn't drag other lines with it
            if(doc.GetElement(id).Name.Equals("Model Lines"))
            {
                ElementArray empty = locationLine.get_ElementsAtJoin(0);
                using (Transaction temp = new Transaction(doc, "Temporary cut joined element"))
                {
                    temp.Start();
                    temp.Commit();
                }
            }
			*/
            double rotation = getRotation(locationLine);

            Line axis = getAxis(locationLine);

            using (Transaction t = new Transaction(doc, "Rotate off axis element."))
            {
                t.Start();
                ElementTransformUtils.RotateElement(doc, id, axis, (rotation));
                t.Commit();
            }
        }
		/// <summary>
        ///  Get Axis of rotation (Z) of line
        /// </summary>
        /// <param name="locationLine"></param>
        /// <returns></returns>
        private Line getAxis(Curve locationLine)
        {
            XYZ basePoint = locationLine.GetEndPoint(0);
            XYZ direction = (locationLine as Line).Direction;
            XYZ cross = XYZ.BasisY.CrossProduct(direction).Normalize();
            return cross.IsZeroLength() ? null : Line.CreateBound(basePoint, cross+basePoint);
        }
        /// <summary>
        ///  Get Axis of rotation (Z) of plane
        /// </summary>
        /// <param name="locationLine"></param>
        /// <returns></returns>
        private Line getAxis(Plane plane)
        {
            XYZ basePoint = plane.Origin;
            XYZ direction = plane.XVec;
            XYZ cross = XYZ.BasisY.CrossProduct(direction).Normalize();
            return cross.IsZeroLength() ? null : Line.CreateBound(basePoint, cross + basePoint);
        }
        /// <summary>
        /// Get rotation angle of line
        /// </summary>
        /// <param name="locationLine"></param>
        /// <returns></returns>
        private double getRotation(Curve locationLine)
        {
            Line line = locationLine as Line;
            double angle = XYZ.BasisY.AngleTo(line.Direction);
            if (angle > 0)
            {
                while (angle > 0.01)
                {
                    angle -= Math.PI / 4;
                }
            }
            else
            {
                while (angle < -0.01)
                {
                    angle += Math.PI / 4;
                }
            }
            return -angle;
        }
		public void ToggleLevelBubbules()
		{
			Document doc = this.ActiveUIDocument.Document;
			Selection sel = this.ActiveUIDocument.Selection;
						
			try{
				do
				{
					Level lvl = doc.GetElement(sel.PickObject(ObjectType.Element, "Pick View to Align To")) as Level;
					
					bool end0 = lvl.IsBubbleVisibleInView(DatumEnds.End0, doc.ActiveView);
					bool end1 = lvl.IsBubbleVisibleInView(DatumEnds.End1, doc.ActiveView);
					
					using(Transaction t = new Transaction(doc, "toggle"))
					{
						t.Start();
						if(end0 && end1)
						{
							lvl.HideBubbleInView(DatumEnds.End1,doc.ActiveView);						
						} 
						else if(end0 && !end1)
						{	
							lvl.HideBubbleInView(DatumEnds.End0,doc.ActiveView);
						}
						else if(!end0 && !end1)
						{							
							lvl.ShowBubbleInView(DatumEnds.End1,doc.ActiveView);
						}
						else
						{
							lvl.ShowBubbleInView(DatumEnds.End0,doc.ActiveView);							
						}
						t.Commit();
					}				
				}
				while(true);
			}				
			catch (Autodesk.Revit.Exceptions.OperationCanceledException exception)
			{
				
			}
		}
		public void DeleteId()
		{
			Document doc = this.ActiveUIDocument.Document;
			ElementId id =  new ElementId(3553753);
			using(Transaction t = new Transaction(doc, "Delete ID"))
			{
				t.Start();
				doc.Delete(id);
				t.Commit();
			}
		}
		public void SelectGroupByMember()
		{
			UIDocument uidoc = this.ActiveUIDocument;
			Document doc = this.ActiveUIDocument.Document;			
			
			Element member = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Select Viewport to Renumber")) as Element;
			
			if(member.GroupId.IntegerValue > 0)
			{
				List<ElementId> groupId = new List<ElementId>(){member.GroupId};
				uidoc.Selection.SetElementIds(groupId);
			}			
		}
		public void ExportDetails()
		{
			UIDocument uidoc = this.ActiveUIDocument;
			Document doc = this.ActiveUIDocument.Document;
			
			var details = new FilteredElementCollector(doc)
				.OfCategory(BuiltInCategory.OST_DetailComponents)
				.OfClass(typeof(FamilySymbol))
				.Cast<FamilySymbol>()
				.ToList();
						
			string path = "";
			
			using(var fbd = new FolderBrowserDialog())
			{
				var result = fbd.ShowDialog();
				
				if(result == DialogResult.OK)
				{
					path = fbd.SelectedPath;		
				}
			}							
			
			if (string.IsNullOrWhiteSpace(path)) return;
			
			foreach(var detail in details)
			{
				try{
					Family family = detail.Family;
					if(family == null) continue;
					Document detailDoc = doc.EditFamily(family);
					detailDoc.SaveAs(path + "\\" + family.Name + ".rfa");
					detailDoc.Close(false);						
				}			
				catch(Exception){}
			}
		}
		public void ElementTypes()
		{
			UIDocument uidoc = this.ActiveUIDocument;
			Document doc = this.ActiveUIDocument.Document;			
			
			var details = new FilteredElementCollector(doc)
				.OfCategory(BuiltInCategory.OST_DetailComponents)
				.OfClass(typeof(FamilySymbol))
				.Cast<FamilySymbol>()
				.ToList();
						
			foreach(var detail in details)
			{
				var elType = doc.GetElement(detail.GetTypeId()) as ElementType;
				System.Drawing.Size imgSize = new System.Drawing.Size( 200, 200 );
				elType.GetPreviewImage(imgSize);
			}
		}
		///
		/// Creates a single 
		ension string between the chosen Level lines
		/// 
		public void DimLevels()
		{			
		  UIDocument uidoc = this.ActiveUIDocument;
		  Document doc = uidoc.Document;
		
		  // Pick all the grid lines you want to dimension to
		  LevelSelectionFilter filter = new ThisApplication.LevelSelectionFilter(doc);			
		  var levels = uidoc.Selection.PickElementsByRectangle(filter, "Pick Grid Lines");
		
		  ReferenceArray refArray = new ReferenceArray();
		  XYZ dir = null;
		
		  foreach(Element el in levels)
		  {
		    Level lv = el as Level;
		
		    if(lv == null) continue;
		    refArray.Append(lv.GetPlaneReference());
		    continue;
		    /*
		    if(dir == null)
		    {
		    	Curve crv = (Reference)lv;
		      dir = new XYZ(0,0,1).CrossProduct((crv.GetEndPoint(0) - crv.GetEndPoint(1)));	// Get the direction of the gridline
		    }
			*/
		    Reference gridRef = null;
		
		    // Options to extract the reference geometry needed for the NewDimension method
		    Options opt = new Options();
		    opt.ComputeReferences = true;
		    opt.IncludeNonVisibleObjects = true;
		    opt.View = doc.ActiveView;
		    foreach (GeometryObject obj in lv.get_Geometry(opt))
		    {
		      if (obj is Line)
		      {
		        Line l = obj as Line;
		        gridRef = l.Reference;
		        refArray.Append(gridRef);	// Append to the list of all reference lines 
		      }
		    }
		  }
		
		  XYZ pickPoint = uidoc.Selection.PickPoint();	// Pick a placement point for the dimension line
		  Line line = Line.CreateBound(pickPoint, pickPoint + XYZ.BasisZ * 100);		// Creates the line to be used for the dimension line
		
		  using(Transaction t = new Transaction(doc, "Make Dim"))
		  {
		    t.Start();
		    if( !doc.IsFamilyDocument )
		    {
		      doc.Create.NewDimension( 
		        doc.ActiveView, line, refArray);
		    }				
		    t.Commit();
		  }					
		}
		/// <summary>
		/// Create Dimensions of Levels in Section
		/// </summary>
		public void DimLevelsSection()
		{			
		  UIDocument uidoc = this.ActiveUIDocument;
		  Document doc = uidoc.Document;
		
		  // Pick all the grid lines you want to dimension to
		  LevelSelectionFilter filter = new ThisApplication.LevelSelectionFilter(doc);			
		  var levels = uidoc.Selection.PickElementsByRectangle(filter, "Pick Grid Lines");
		
		  ReferenceArray refArray = new ReferenceArray();
		  XYZ dir = null;
		
		  foreach(Element el in levels)
		  {
		    Level lv = el as Level;
		
		    if(lv == null) continue;
		    refArray.Append(lv.GetPlaneReference());
		    continue;
		  }
				  
		  try{		  	
			  using(Transaction tr = new Transaction(doc, "Set workplane"))
			  {
			    tr.Start();			    
			    
				View view = uidoc.ActiveView;				
						  	
				var origin = view.Origin;
				var direction = view.ViewDirection;
							  	  
				Plane plane = Plane.CreateByNormalAndOrigin(direction, origin);
				SketchPlane sp = SketchPlane.Create(doc, plane);
				
				view.SketchPlane = sp;	
			    doc.Regenerate();		  
				
			    tr.Commit();
			  }
		  }
		  catch(Exception ex)
		  {
		  	TaskDialog.Show("Error", ex.Message);
		  }		
		  
		  try{		  	
			  using(Transaction t = new Transaction(doc, "Make Dim"))
			  {
			    t.Start();			    
			    
				XYZ pickPoint = uidoc.Selection.PickPoint();	// Pick a placement point for the dimension line
				Line line = Line.CreateBound(pickPoint, pickPoint + XYZ.BasisZ * 100);		// Creates the line to be used for the dimension line
		  
			    if( !doc.IsFamilyDocument )
			    {
			      doc.Create.NewDimension( 
			        doc.ActiveView, line , refArray);
			    }				
			    t.Commit();
			  }
		  }
		  catch(Exception ex)
		  {
		  	TaskDialog.Show("Error", ex.Message);
		  }					
		}
		///
		/// Levels Selection Filter (example for selection filters)
		/// 
		public class LevelSelectionFilter : ISelectionFilter
		{
		  Document doc = null;
		  public LevelSelectionFilter(Document document)
		  {
		    doc = document;
		  }
		
		  public bool AllowElement(Element element)
		  {
		    if(element.Category.Name == "Levels")
		    {
		      return true;
		    }
		    return false;
		  }
		
		  public bool AllowReference(Reference refer, XYZ point)
		  {
		    return true;
		  }
		}
		public void RenumberViewports()
		{
			UIDocument uidoc = this.ActiveUIDocument;
			Document doc = this.ActiveUIDocument.Document;
			
			ViewSheet vsheet = doc.ActiveView as ViewSheet;
			
			if(vsheet == null) return;
			
			var vportsElementIds = vsheet.GetAllViewports();
			List<Viewport> vports = new List<Viewport>();
			
			foreach(var vpId in vportsElementIds)
			{
				vports.Add(doc.GetElement(vpId) as Viewport);
			}
			int counter = 0;
			
			try{				
				do
				{
					counter++;
					Viewport vp = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Select Viewport to Renumber")) as Viewport;
					if(vp == null) return;
					Parameter vp_detailnumber = vp.LookupParameter("Detail Number");
									
					using(Transaction t = new Transaction(doc, "toggle"))
					{
						t.Start();
						var vp_change = vports.Where(x => x.LookupParameter("Detail Number").AsString() == counter.ToString());
						string carry = vp_detailnumber.AsString();
						vp_detailnumber.Set("99");
						if(vp_change != null && vp_change.Count() > 0) vp_change.First().LookupParameter("Detail Number").Set(carry);
						vp_detailnumber.Set(counter.ToString());
						t.Commit();
					}				
				}
				while(true);
			}			
			catch (Autodesk.Revit.Exceptions.OperationCanceledException exception)
			{
				
			}
		}
		public void FilterSelect()
		{
			UIDocument uidoc = ActiveUIDocument;
			Document doc = uidoc.Document;
						
			Reference refer = uidoc.Selection.PickObject(ObjectType.Element, "Set the filter"); //Pick an object by which Category you will filter
			
			IList<Element> elements = uidoc.Selection.PickElementsByRectangle();
						
			uidoc.Selection.SetElementIds(elements
				.Where(x => x.Category.Id.IntegerValue.Equals(doc.GetElement(refer).Category.Id.IntegerValue))
				.Select(x => x.Id)
				.ToList());
		}
		public void AssignViewType()
		{
			Document doc = this.ActiveUIDocument.Document;
			
			List<ViewSheet> sheets = new FilteredElementCollector(doc)
				.OfClass(typeof(ViewSheet))
				.Cast<ViewSheet>()
				.Where(x => x.LookupParameter("Planart").HasValue 
				       && x.LookupParameter("Planart").AsString().Contains("70-Leitdetails"))
				.ToList();
			
			var vpp = new FilteredElementCollector(doc).OfClass(typeof(Viewport))
				.Cast<Viewport>()
				.Where(x => x.Name.Equals("Detail Section - Titlebar BIG")).First();
						
			using(Transaction t = new Transaction(doc, "changenames"))
			{
				t.Start();
				foreach(ViewSheet vs in sheets)
				{
					var viewports = vs.GetAllViewports();
					if(viewports.Count > 0)
					{
						foreach(ElementId vId in viewports)
						{
							Viewport vp = doc.GetElement(vId) as Viewport;
							vp.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).Set(vpp.GetTypeId());
						}
					}
				}				
				t.Commit();
			}
		}
		public void ReplaceViewport()
		{
		  UIDocument uidoc = this.ActiveUIDocument;
		  Document doc = uidoc.Document;
		  
		  var viewport = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Pick Viewport to replace").ElementId) as Viewport;
		  var view = doc.GetElement(viewport.ViewId) as View;
		  
		  var sheet = viewport.SheetId;
		  var point = viewport.GetBoxCenter();
		  
		  var replacementView = new FilteredElementCollector(doc)
		  	.OfClass(typeof(View))
		  	.Cast<View>()
		  	.Where(v => v.Name.Equals("B1 - General Arrangement"))
		  	.First();
		  
		  var delete = new List<ElementId>();
		  delete.Add(viewport.Id);
		  
		  using(Transaction t = new Transaction(doc, "Replace viewport"))
		  {
		  	t.Start();
		  	Viewport.Create(doc, sheet, replacementView.Id, point);
		  	doc.Delete(delete);
		  	t.Commit();
		  }	  
		}
		public void RenumberViewports()
		{
			UIDocument uidoc = this.ActiveUIDocument;
			Document doc = this.ActiveUIDocument.Document;
			
			ViewSheet vsheet = doc.ActiveView as ViewSheet;
			
			if(vsheet == null) return;
			
			var vportsElementIds = vsheet.GetAllViewports();
			List<Viewport> vports = new List<Viewport>();
			
			foreach(var vpId in vportsElementIds)
			{
				vports.Add(doc.GetElement(vpId) as Viewport);
			}
			int counter = 0;
			
			try{				
				do
				{
					counter++;
					Viewport vp = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Select Viewport to Renumber")) as Viewport;
					if(vp == null) return;
					Parameter vp_detailnumber = vp.LookupParameter("Detail Number");
									
					using(Transaction t = new Transaction(doc, "toggle"))
					{
						t.Start();
						var vp_change = vports.Where(x => x.LookupParameter("Detail Number").AsString() == counter.ToString());
						string carry = vp_detailnumber.AsString();
						vp_detailnumber.Set("99");
						if(vp_change != null && vp_change.Count() > 0) vp_change.First().LookupParameter("Detail Number").Set(carry);
						vp_detailnumber.Set(counter.ToString());
						t.Commit();
					}				
				}
				while(true);
			}			
			catch (Autodesk.Revit.Exceptions.OperationCanceledException exception)
			{
				
			}
		}
		public void AlignViews()
		{
			Document doc = this.ActiveUIDocument.Document;
			Selection sel = this.ActiveUIDocument.Selection;
			
			Viewport vp = doc.GetElement(sel.PickObject(ObjectType.Element, "Pick View to Align To")) as Viewport;
			
			if(vp == null) 
			{
				TaskDialog.Show("Error", "That's not a View");
				return;
			}
						
			
				
			XYZ loc = vp.GetBoxCenter();
			
			List<ViewSheet> viewSheets = new FilteredElementCollector(doc)
				.OfClass(typeof(ViewSheet))
				.WhereElementIsNotElementType()
				.Cast<ViewSheet>()
				.ToList();
			
			List<Viewport> viewPorts = new List<Viewport>();
			
			foreach(ViewSheet sheet in viewSheets)
			{
				List<Viewport> ports = sheet.GetAllViewports()
					.Select<ElementId, Viewport>(
						id => doc.GetElement(id) as Viewport)
					.ToList<Viewport>();
				
				viewPorts.AddRange(ports);
			}
			
			viewPorts = viewPorts.Where(x => x.LookupParameter("View Name").AsString().Contains("DS -")).ToList();
			
			
			
			using(Transaction t = new Transaction(doc, "Align Views"))
			{
				t.Start();
				foreach(Viewport v in viewPorts)
				{
					if(v.Id == vp.Id) continue;
					XYZ delta = loc - v.GetBoxCenter();
					ElementTransformUtils.MoveElement(doc, v.Id, delta);
				}
				t.Commit();
			}
		}
		#endregion
			
	public void CollectElementIDs()
	{
	    UIDocument uidoc = ActiveUIDocument;
            Document doc = ActiveUIDocument.Document;
            
            View current = doc.ActiveView;
            
            List<ElementId> ids = new FilteredElementCollector(doc, current.Id)
            	.WhereElementIsNotElementType()
            	.ToElementIds()
            	.ToList();
            
            
			var builder = new StringBuilder();
			
			foreach(ElementId id in ids)
			{
				builder.AppendLine(String.Join("\t",id.ToString()));
			}
            
            
			var file = new FileStream("C:/Temp/report.txt",FileMode.Create);
			var writer = new StreamWriter(file);
			writer.Write(builder.ToString());
			writer.Flush();
			writer.Close();
		}     
		
		private string StringDebug(List<String> stringlist)
		{
			string s = "";
			
			foreach(String str in stringlist)
			{
				s += String.Format("{0}{1}", str, Environment.NewLine);
			}
			
			return s;
		}
		public void SetTypeParameter()
		{
			Document doc = this.ActiveUIDocument.Document;
			
			FilteredElementCollector col = new FilteredElementCollector(doc);
			List<Element> wallTypes = col.OfCategory(BuiltInCategory.OST_Assemblies).WhereElementIsElementType().ToList();
			
			string message = "";
			using (Transaction t = new Transaction(doc,"LOD_UniFormat - Elements"))
			{
				t.Start();
				
				foreach(Element element in wallTypes)
				{			
					if  //((element as FamilySymbol).FamilyName.IndexOf("door", StringComparison.OrdinalIgnoreCase) >= 0) 
						//(true) 
						(element.Name.IndexOf("panel", StringComparison.OrdinalIgnoreCase) >= 0)
					{
						element.LookupParameter("LOD_UniFormat").Set("B2080 - Exterior Wall Appurtenances");
						message += element.Name + "\n";					
					}
				}
				
				t.Commit();
			}
			
			TaskDialog.Show("Wall types", message);
		}
		
		public void ImportedDWG()
        {
            Document doc = ActiveUIDocument.Document;
            
            FilteredElementCollector col = new FilteredElementCollector(doc)
                .OfClass(typeof(ImportInstance));
            
            IList<ImportInstance> elements = col
                .Cast<ImportInstance>()
                .Where(x => !x.IsLinked)
                .ToList();
            
            TaskDialog.Show("NumberImports", String.Format("There are {0} number of imported files", 
                                                           elements.Count().ToString()));
            string s = "";
            
            foreach(ImportInstance instance in elements)
            {
                s += instance.LookupParameter("Name").AsString() + " : " + instance.OwnerViewId.ToString() + Environment.NewLine;
            }
            TaskDialog.Show("NumberImports", String.Format("There are {0}{1}", Environment.NewLine,
                                                           s));
            
            return;
        }public void DimRefPlanes()
		{			
			UIDocument uidoc = this.ActiveUIDocument;
			Document doc = uidoc.Document;
			
			RefPlanesSelectionFilter filter = new ThisApplication.RefPlanesSelectionFilter(doc);			
			IList<Element> planes = uidoc.Selection.PickElementsByRectangle(filter, "Pick Ref Plans");
			
			ReferenceArray refArray = new ReferenceArray();
			XYZ dir = null;
			
			foreach(Element el in planes)
			{
				ReferencePlane gr = el as ReferencePlane;				
				Plane pl = gr.GetPlane();
				
				if(gr == null) continue;
				if(dir == null)
				{
					dir = pl.Normal;
				}
				
				
				Reference gridRef = null;
				gridRef = gr.GetReference();
					
				refArray.Append(gridRef);
			}
			
			XYZ pickPoint = uidoc.Selection.PickPoint();
			Line line = Line.CreateBound(pickPoint, pickPoint + dir * 100);
			
			using(Transaction t = new Transaction(doc, "Make Dim"))
			{
				t.Start();
				if( !doc.IsFamilyDocument )
				{
					doc.Create.NewDimension( 
					  doc.ActiveView, line, refArray);
				}				
				t.Commit();
			}					
		}	
        public void CreateWorksetView()
		{
			Document doc = this.ActiveUIDocument.Document;
			ElementId viewTypeId = FindViewTypes(doc, ViewType.ThreeD);
			
			TaskDialog.Show("Test", viewTypeId.ToString());
			
			List<Workset> workCollector = new FilteredWorksetCollector(doc)
				.OfKind(WorksetKind.UserWorkset)
				.Cast<Workset>()
				.ToList();
				
			using(Transaction t = new Transaction(doc, "Create 3D View per Workset"))
			{
				t.Start();
				foreach(var workset in workCollector)
				{
					string name = workset.Name;
					Autodesk.Revit.DB.View v = View3D.CreateIsometric(doc, viewTypeId);
					v.Name = "WORKSET VIEW - " + name;
					SetWorksetVisibility(v, workset, workCollector);
				}
								
				t.Commit();
			}
		}
		internal void SetWorksetVisibility(Autodesk.Revit.DB.View view, Workset workset, List<Workset> workCollector)
		{
			foreach(var w in workCollector)
			{
				if(w == workset)
				{
					view.SetWorksetVisibility(w.Id, WorksetVisibility.Visible);
				}
				else{
					view.SetWorksetVisibility(w.Id, WorksetVisibility.Hidden);
				}
			}
		}
		internal ElementId FindViewTypes(Document doc, ViewType type)
		{
			var viewFamilyType = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>()
                          .FirstOrDefault(x => x.ViewFamily ==  ViewFamily.ThreeDimensional);
			    
		    return viewFamilyType.Id;
		}
        public void PurgeImportedLines()
        {
            Document doc = ActiveUIDocument.Document;
            
            string m = "";
            
            FilteredElementCollector col = new FilteredElementCollector(doc)
                .OfClass(typeof(LinePatternElement));
            
            List<ElementId> lpeIds = col.ToElementIds().Where(x => doc.GetElement(x).Name.Contains("IMPORT")).ToList();
            List<LinePatternElement> linePatterns = col.Cast<LinePatternElement>().Where(x => x.Name.Contains("IMPORT")).ToList();
            
            using (Transaction t = new Transaction(doc, "Purge imported line patterns"))
            {
                t.Start();
                foreach(LinePatternElement lpe in linePatterns)
                {
                    m += lpe.Name + Environment.NewLine;
                }
                doc.Delete(lpeIds);
                t.Commit();
            }
            
            m += Environment.NewLine + String.Format(("A total of {0} imported line patterns have been removed from this project"), linePatterns.Count.ToString());
            TaskDialog.Show("PurgeImportedLines", m);
            return;            
        }
		public void ShowSubcat()
		{
			UIDocument uidoc = ActiveUIDocument;
            Document doc = ActiveUIDocument.Document;
            
            Categories categories = doc.Settings.Categories;
            
            TaskDialog.Show("SubCategories", categories.Size.ToString());
            
            List<Category> subcat = new List<Category>();
            string s = "";
            string n = "";
            
            foreach(Category cat in categories)
            {
            	var sub = getSubCategories(doc, cat, out n);
            	if (sub != null) 
            	{
            		subcat.AddRange(sub);
            	    s += n;
            	}
            }
            TaskDialog.Show("SubCategories", s);
		}
		internal List<Category> getSubCategories(Document doc, Category cat, out string s)
        {
            List<Category> listCat = null;
            var categories = cat.SubCategories;
            s = "";
            if(!categories.IsEmpty)
            {
                listCat = new List<Category>();
                foreach (Category subCat in categories)
                {
                    listCat.Add(subCat);
                    s += subCat.Name + Environment.NewLine;
                }
            }
            return listCat;
        }
		public void RenameWallTypes()
		{
			UIDocument uidoc = ActiveUIDocument;
            Document doc = ActiveUIDocument.Document;
            View current = doc.ActiveView;
            
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            FilteredElementCollector wallCollector = new FilteredElementCollector(doc);
			
            List<Element> wallTypes = collector.OfClass(typeof(WallType)).ToElements().ToList();
  
            wallTypes = wallTypes.OrderBy(z => z.Name).ToList();
                        
            int num = wallTypes.Count;
            string target = "scr";
            string replace = "SCRN";
                       
            using(Transaction t = new Transaction(doc, "FilledRegionPopulate"))
            {
            	t.Start();	                	
	            foreach(Element wall in wallTypes)
	            {	    
	            	Wall w = wallCollector.OfClass(typeof(Wall)).Cast<Wall>().Where(x => x.WallType.Id == wall.Id).FirstOrDefault();
	            	try
	            	{
		            	if(wall.Name.IndexOf(target, StringComparison.OrdinalIgnoreCase) >= 0)
		            	{
		            		string newName = String.Format("{0}_{1}_{2}","GHA",replace,(w.Width*304.8).ToString());
		            		//TaskDialog.Show("test", newName);
		            		wall.Name = newName;
		            	}	            		
	            	}
	            	catch(Exception)
	            	{
	            		
	            	}
	            }  
	            t.Commit();	     
            }			            
	}	
	public void DuplicateNumberValues()
	{
		Document doc = this.ActiveUIDocument.Document;
		Selection selection = this.ActiveUIDocument.Selection;
		List<ElementId> ids = selection.GetElementIds().ToList();

		List<int> numbers = new FilteredElementCollector(doc)
			.OfCategoryId(doc.GetElement(ids.First()).Category.Id)
			.WhereElementIsNotElementType()
			.Select(x => int.Parse(x.LookupParameter("Number").AsString()))
			.ToList();

		List<int> result = Enumerable.Range(0, numbers.Max()).Except(numbers).ToList();

		using(Transaction t = new Transaction(doc, "Duplicating Number Elements"))
		{
			t.Start();				
			foreach(ElementId id in ids)
			{
				int number = 0;
				Element el = doc.GetElement(id);
				Parameter param = el.LookupParameter("Number");

				Int32.TryParse(param.AsString(), out number);
				if(number != 0)
				{
					number = result[0];
					param.Set(number.ToString());
					result.RemoveAt(0);
				}
			}
			t.Commit();
		}

	}
	public void PopulateFloors()
	{
	    UIDocument uidoc = ActiveUIDocument;
            Document doc = ActiveUIDocument.Document;
            View current = doc.ActiveView;
            
            FilteredElementCollector collector = new FilteredElementCollector(doc);
			
            List<Element> floorTypes = collector.OfClass(typeof(FloorType)).ToElements().ToList();
            floorTypes = floorTypes.OrderBy(z => z.Name).ToList();
                        
            string s = "";
            int num = floorTypes.Count;
            
            CurveArray crvArr = new CurveArray();
            
            double width = 3.0;  //horizontal offset
            double co = 2.0;
            
            int length = Convert.ToInt16(Math.Ceiling(Math.Sqrt(num)));
            
            XYZ pickPoint = uidoc.Selection.PickPoint("Pick a starting point");
            
            int x = 0;
            int y = 0;
            
            TextNoteOptions toptions = new TextNoteOptions();
            toptions.TypeId = doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType);
            
            using(Transaction t = new Transaction(doc, "FilledRegionPopulate"))
            {
            	t.Start();	            	
            	int counter = 0;
            	try{
		            foreach(Element floor in floorTypes)
		            {	           
						FloorType fType = floor as FloorType;
						
						if (fType == null) continue;
						TaskDialog.Show("test",fType.Name);
		            	x = counter%length + 1;
		            	y = counter/length + 1;
		            	
		            	XYZ pos = new XYZ(co*x*width, co*y*width,0);
		            	pos += pickPoint;
		            	XYZ halfLength = new XYZ(0,0,0);
		            	
		            	XYZ [] points = new XYZ[]{
		            		new XYZ(-width, -width,0) + pos,
		            		new XYZ(-width, width,0) + pos,
		            		new XYZ(width, width,0) + pos,
		            		new XYZ(width, -width,0) + pos,
		            	};
		            	
		            	for(int i = 0; i < 4; i++)
		            	{
		            		Curve c  = Line.CreateBound(points[i],points[(i+1)%4]);
		            		crvArr.Append(c);
		            	}
	        			
		            	
		            	doc.Create.NewFloor(crvArr, fType, current.GenLevel, false, new XYZ(0,0,1));
		            	
		            	counter ++;
		            	s += counter.ToString();
		            	} 
        			}
		            catch(Exception)
		            {
		            	
		            }
	            t.Commit();	     
            }			            
		}
		public void AssignViewType()
		{
			Document doc = this.ActiveUIDocument.Document;
			
			List<ViewSheet> sheets = new FilteredElementCollector(doc)
				.OfClass(typeof(ViewSheet))
				.Cast<ViewSheet>()
				.Where(x => x.LookupParameter("Planart").HasValue 
				       && x.LookupParameter("Planart").AsString().Contains("70-Leitdetails"))
				.ToList();
			
			var vpp = new FilteredElementCollector(doc).OfClass(typeof(Viewport))
				.Cast<Viewport>()
				.Where(x => x.Name.Equals("Detail Section - Titlebar BIG")).First();
						
			using(Transaction t = new Transaction(doc, "changenames"))
			{
				t.Start();
				foreach(ViewSheet vs in sheets)
				{
					var viewports = vs.GetAllViewports();
					if(viewports.Count > 0)
					{
						foreach(ElementId vId in viewports)
						{
							Viewport vp = doc.GetElement(vId) as Viewport;
							vp.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).Set(vpp.GetTypeId());
						}
					}
				}				
				t.Commit();
			}
		}
		public void PopulateWalls()
		{
			UIDocument uidoc = ActiveUIDocument;
            Document doc = ActiveUIDocument.Document;
            View current = doc.ActiveView;
            
            FilteredElementCollector collector = new FilteredElementCollector(doc);
			
            List<Element> wallTypes = collector.OfClass(typeof(WallType)).ToElements().ToList();
            wallTypes = wallTypes.OrderBy(z => z.Name).ToList();
                        
            string s = "";
            int num = wallTypes.Count;
            
            List<CurveLoop> cloop = new List<CurveLoop>();
            CurveLoop loop = new CurveLoop();
            
            double width = 3.0;  //horizontal offset
            double height = 6.5;  //vertical offset
            double wallLength = 5.0;
            double wallHeight = 10.0;
            double co = 2.0;
            
            int length = Convert.ToInt16(Math.Ceiling(Math.Sqrt(num)));
            
            XYZ pickPoint = uidoc.Selection.PickPoint("Pick a starting point");
            
            int x = 0;
            int y = 0;
            
            TextNoteOptions toptions = new TextNoteOptions();
            toptions.TypeId = doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType);
            
            using(Transaction t = new Transaction(doc, "FilledRegionPopulate"))
            {
            	t.Start();	            	
            	int counter = 0;
            	
	            foreach(Element wall in wallTypes)
	            {	            	
	            	x = counter%length + 1;
	            	y = counter/length + 1;
	            	
	            	XYZ pos = new XYZ(co*x*width, co*y*height,0);
	            	pos += pickPoint;
	            	XYZ halfLength = new XYZ(0,wallLength,0);
	            	
	            	Curve c = Line.CreateBound(pos+halfLength,pos-halfLength);   
        			
	            	Wall.Create(doc,c,wall.Id,current.GenLevel.Id,wallHeight,0.0,false,false);
	            	
	            	counter ++;
	            	s += counter.ToString();
	            }  
	            t.Commit();	     
            }			            
		}
		public void PopulateFilledRegion()
		{
			UIDocument uidoc = ActiveUIDocument;
            Document doc = ActiveUIDocument.Document;
            View current = doc.ActiveView;
            
            FilteredElementCollector collector = new FilteredElementCollector(doc);
			
            List<Element> filledTypes = collector.OfClass(typeof(FilledRegionType)).ToElements().ToList();
            filledTypes = filledTypes.OrderBy(z => z.Name).ToList();
                        
            string s = "";
            int num = filledTypes.Count;
            
            List<CurveLoop> cloop = new List<CurveLoop>();
            CurveLoop loop = new CurveLoop();
            
            double width = 3.0;
            double height = 1.5;
            double co = 2.0;
            
            int length = Convert.ToInt16(Math.Ceiling(Math.Sqrt(num)));
            int x = 0;
            int y = 0;
            
            TextNoteOptions toptions = new TextNoteOptions();
            toptions.TypeId = doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType);            
            
            XYZ pickPoint = uidoc.Selection.PickPoint("Pick a starting point");
            
            using(Transaction t = new Transaction(doc, "FilledRegionPopulate"))
            {
            	t.Start();	            	
            	int counter = 0;
            	
            	List<XYZ> coordinates = new List<XYZ>{
            		new XYZ(0,0,0),
            		new XYZ(width,0,0),
            		new XYZ(width,height,0),
            		new XYZ(0,height,0),
            	};
            	
	            foreach(Element ft in filledTypes)
	            {
	            	List<Curve> curves = new List<Curve>();
	            	
	            	x = counter%length + 1;
	            	y = counter/length + 1;
	            	
	            	XYZ pos = new XYZ(co*x*width, -co*y*height,0);
	            	pos += pickPoint;
	            	
	            	Curve c = null;
	            	
	            	for(int i = 0; i < coordinates.Count; i++)
	            	{
	            		c = Line.CreateBound(coordinates[i].Add(pos),coordinates[(i+1)%(coordinates.Count)].Add(pos));
	            		curves.Add(c);
	            	}
            			            		
	            	loop = CurveLoop.Create(curves);
        			cloop.Add(loop);
        			
	            	FilledRegion.Create(doc,ft.Id,current.Id,cloop);
	            	cloop.Clear();
	            	TextNote.Create(doc,current.Id, pos.Add(new XYZ(0,-0.2,0)), 0.1,Truncate(ft.Name, 20),toptions);
	            	counter ++;
	            	s += counter.ToString();
	            }  
	            t.Commit();	     
            }			            
		}
		private string Truncate(string value, int maxLength)
	    {
	        if (string.IsNullOrEmpty(value)) return value;
	        return value.Length <= maxLength ? value : value.Substring(0, maxLength)+("..");
	    }
		public void RenumberViewports()
		{
			UIDocument uidoc = this.ActiveUIDocument;
			Document doc = this.ActiveUIDocument.Document;
			
			ViewSheet vsheet = doc.ActiveView as ViewSheet;
			
			if(vsheet == null) return;
			
			var vportsElementIds = vsheet.GetAllViewports();
			List<Viewport> vports = new List<Viewport>();
			
			foreach(var vpId in vportsElementIds)
			{
				vports.Add(doc.GetElement(vpId) as Viewport);
			}
			int counter = 0;
			
			try{				
				do
				{
					counter++;
					Viewport vp = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Select Viewport to Renumber")) as Viewport;
					if(vp == null) return;
					Parameter vp_detailnumber = vp.LookupParameter("Detail Number");
									
					using(Transaction t = new Transaction(doc, "toggle"))
					{
						t.Start();
						var vp_change = vports.Where(x => x.LookupParameter("Detail Number").AsString() == counter.ToString()).First();
						string carry = vp_detailnumber.AsString();
						vp_detailnumber.Set("99");
						vp_change.LookupParameter("Detail Number").Set(carry); 
						vp_detailnumber.Set(counter.ToString());
						t.Commit();
					}				
				}
				while(true);
			}			
			catch (Autodesk.Revit.Exceptions.OperationCanceledException exception)
			{
				
			}

		}
		public void FamilyRename()
		{
			UIDocument uidoc = ActiveUIDocument;
            Document doc = ActiveUIDocument.Document;
            
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            //List<Element> famTypes = collector.OfClass(typeof(Family)).ToElements().ToList();
            List<Element> famTypes = collector.OfClass(typeof(FilledRegionType)).ToElements().ToList();
            
            string s = "";
            
            using(Transaction t = new Transaction(doc, "FamilyRename"))
            {
            	try{
	            	t.Start();
		            foreach(Element ft in famTypes)
		            {
		            	if(ft.Name.Contains("wildcard"))
		            	{
		            		ft.Name = ft.Name.Replace("wildcard", "_GHA");
		            		s += string.Format("{0} {1}", ft.Name, Environment.NewLine);
		            	}		            	
		            }  
		            t.Commit();	            		
            	}
            	catch(Exception)
            	{
            		
            	}
            }
			
            TaskDialog.Show("test", s);
		}		
		public void RemoveEmptyElevMarkers()
		{
			UIDocument uidoc = ActiveUIDocument;
            Document doc = ActiveUIDocument.Document;
            
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            
            IEnumerable<Element> markers = collector.
            	WhereElementIsNotElementType().
            	OfClass(typeof(ElevationMarker)).
            	Cast<ElevationMarker>().
            	Where(x => !x.HasElevations());
            
            int c = 0;
            
            using(Transaction t = new Transaction(doc,"Remove Empty Elevation Marks"))
            {
            	t.Start();
            	foreach(var mark in markers.ToList())
            	{
            		doc.Delete(mark.Id);
            		c++;
            	}
            	t.Commit();
            }
            
            TaskDialog.Show("Success", String.Format("Successfully removed {0} unused Elevation Markers", c.ToString()));
		}
		public void BulkReloadLinks()
		{
			UIApplication uiApp = this;
			UIDocument uidoc = uiApp.ActiveUIDocument;
			Document doc = uiApp.ActiveUIDocument.Document;
			
			List<RevitLinkType> links = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkType)).Cast<RevitLinkType>().ToList();
			
			System.Windows.Forms.OpenFileDialog theDialogRevit = new System.Windows.Forms.OpenFileDialog();
			theDialogRevit.Title = "Select Revit Project Files";
			theDialogRevit.Filter = "RVT files|*.rvt";
			theDialogRevit.FilterIndex = 1;
			theDialogRevit.Multiselect = true;
			
			if (theDialogRevit.ShowDialog() == System.Windows.Forms.DialogResult.OK)
			{
                foreach (String projectPath in theDialogRevit.FileNames)
                {
                 	FileInfo filePath = new FileInfo(projectPath);
                 	string filename = filePath.Name;
                    ModelPath mp = ModelPathUtils.ConvertUserVisiblePathToModelPath(filePath.FullName);					
		            WorksetConfiguration wc = new WorksetConfiguration(WorksetConfigurationOption.OpenAllWorksets);
		            
		            RevitLinkType link = links.FirstOrDefault(x => x.Name.Equals(filename) && x.GetParentId().IntegerValue == -1);
		            
		            if(link != null)
		            {
		            	link.LoadFrom(mp, wc);	
		            }
                }
			}
		}
	
		public void DimGrids()
		{			
			UIDocument uidoc = this.ActiveUIDocument;
			Document doc = uidoc.Document;
			
			GridSelectionFilter filter = new ThisApplication.GridSelectionFilter(doc);			
			IList<Element> grids = uidoc.Selection.PickElementsByRectangle(filter, "Pick Grid Lines");
			
			ReferenceArray refArray = new ReferenceArray();
			XYZ dir = null;
			
			foreach(Element el in grids)
			{
				Grid gr = el as Grid;
				
				if(gr == null) continue;
				if(dir == null)
				{
					Curve crv = gr.Curve;
					dir = new XYZ(0,0,1).CrossProduct((crv.GetEndPoint(0) - crv.GetEndPoint(1)));
				}
				Reference gridRef = null;
				Options opt = new Options();
				opt.ComputeReferences = true;
				opt.IncludeNonVisibleObjects = true;
				opt.View = doc.ActiveView;
				foreach (GeometryObject obj in gr.get_Geometry(opt))
				{
					if (obj is Line)
					{
						Line l = obj as Line;
						gridRef = l.Reference;
						refArray.Append(gridRef);
					}
				}
			}
			
			XYZ pickPoint = uidoc.Selection.PickPoint();
			Line line = Line.CreateBound(pickPoint, pickPoint + dir * 100);
			
			using(Transaction t = new Transaction(doc, "Make Dim"))
			{
				t.Start();
				if( !doc.IsFamilyDocument )
				{
					doc.Create.NewDimension( 
					  doc.ActiveView, line, refArray);
				}				
				t.Commit();
			}					
		}	
	public void SelectUniqueModelGroups()
		{
			UIDocument uidoc = ActiveUIDocument;
			Document doc = uidoc.Document;
		    View current = doc.ActiveView;
		    
		    List<GroupType> gtypes = new FilteredElementCollector(doc).OfClass(typeof(GroupType)).WhereElementIsElementType().Cast<GroupType>().ToList();
		    List<Group> groups = new FilteredElementCollector(doc).OfClass(typeof(Group)).WhereElementIsNotElementType().Cast<Group>().ToList();
		    		    
		    HashSet<ElementId> usedTypes = new HashSet<ElementId>();
		    HashSet<ElementId> safe = new HashSet<ElementId>();
		    
		    foreach(var g in groups)
		    {
		    	if(!usedTypes.Add(g.GroupType.Id) && usedTypes.Contains(g.GroupType.Id))
		    	{
		    		safe.Add(g.GroupType.Id);
		    	}
		    }
		    
		    var intersect = usedTypes.Except(safe);
		    
		    //TaskDialog.Show("Test", String.Format("{0} safe and {1} used and {2} total", safe.Count.ToString(), usedTypes.Count.ToString(), gtypes.Count().ToString()));
		    
		    var test = groups.Where(x => intersect.Contains(x.GroupType.Id)).Select(x => x.Id);
		    uidoc.Selection.SetElementIds(test.ToList());
		}
		public void ReportCopyMonitored()
		{
			UIDocument uidoc = ActiveUIDocument;
			Document doc = uidoc.Document;
		    View current = doc.ActiveView;

		    FilteredElementCollector collector = new FilteredElementCollector(doc);        
		    
		    IList<Element> monitored = collector.WhereElementIsNotElementType().Where(x => x.GetMonitoredLinkElementIds().Count > 0).ToList();
		    
		    TaskDialog.Show("Test", monitored.Count.ToString());			
		}
		public void ReportDesignOptions()
		{
			UIDocument uidoc = ActiveUIDocument;
			Document doc = uidoc.Document;
		    View current = doc.ActiveView;

		    FilteredElementCollector collector = new FilteredElementCollector(doc);        
		    
		    IList<Element> dosets = collector.OfCategory(BuiltInCategory.OST_DesignOptionSets).ToList();
		   
		    TaskDialog.Show("Test", dosets.Count.ToString());	
			
		}	
	public void SaveRevitFiles()
	{   
			OpenFileDialog theDialogRevit = new OpenFileDialog();
			theDialogRevit.Title = "Select Revit Project Files";
			theDialogRevit.Filter = "RVT files|*.rvt";
			theDialogRevit.FilterIndex = 1;
			theDialogRevit.InitialDirectory = @"C:\";
			theDialogRevit.Multiselect = true;
			if (theDialogRevit.ShowDialog() == DialogResult.OK)
			{
				string mpath = "";
		        string mpathOnlyFilename = "";
		        FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();
		        folderBrowserDialog1.Description = "Select Folder Where Revit Projects to be Saved in Local";
		        folderBrowserDialog1.RootFolder = Environment.SpecialFolder.MyComputer;
		        if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
		        {
		            mpath = folderBrowserDialog1.SelectedPath;
	                foreach (String projectPath in theDialogRevit.FileNames)
	                {
	                 	FileInfo filePath = new FileInfo(projectPath);
                        ModelPath mp = ModelPathUtils.ConvertUserVisiblePathToModelPath(filePath.FullName);
                        OpenOptions opt = new OpenOptions();
                        opt.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets;
                        opt.Audit = true;
                        mpathOnlyFilename = filePath.Name;
                        Document openedDoc = Application.OpenDocumentFile(mp, opt);   
                        SaveAsOptions options = new SaveAsOptions();
                        if(openedDoc.IsWorkshared)
                        {
	                        WorksharingSaveAsOptions wsOptions = new WorksharingSaveAsOptions();
	                        wsOptions.SaveAsCentral = true;
                        	options.SetWorksharingOptions(wsOptions);                          	
                        }
                        options.OverwriteExistingFile = true;
                        options.MaximumBackups = 1;
                        ModelPath modelPathout = ModelPathUtils.ConvertUserVisiblePathToModelPath(mpath + "\\" + mpathOnlyFilename);
                        openedDoc.SaveAs(modelPathout, options);
                        openedDoc.Close(false);
	                }
		        }
		}
       	}	
		public void FilledRegionPopulate()
		{
		    UIDocument uidoc = ActiveUIDocument;
		    Document doc = ActiveUIDocument.Document;
		    View current = doc.ActiveView;

		    FilteredElementCollector collector = new FilteredElementCollector(doc);

		    List<Element> filledTypes = collector.OfClass(typeof(FilledRegionType)).ToElements().ToList();
		    filledTypes = filledTypes.OrderBy(z => z.Name).ToList();

		    string s = "";
		    int num = filledTypes.Count;

		    List<CurveLoop> cloop = new List<CurveLoop>();
		    CurveLoop loop = new CurveLoop();

		    double width = 3.0;
		    double height = 1.5;
		    double co = 2.0;

		    int length = Convert.ToInt16(Math.Ceiling(Math.Sqrt(num)));
		    int x = 0;
		    int y = 0;

		    TextNoteOptions toptions = new TextNoteOptions();
		    toptions.TypeId = doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType);

		    using(Transaction t = new Transaction(doc, "FilledRegionPopulate"))
		    {
			t.Start();                    
			int counter = 0;

			List<XYZ> coordinates = new List<XYZ>{
			    new XYZ(0,0,0),
			    new XYZ(width,0,0),
			    new XYZ(width,height,0),
			    new XYZ(0,height,0),
			};

			foreach(Element ft in filledTypes)
			{
			    List<Curve> curves = new List<Curve>();

			    x = counter%length + 1;
			    y = counter/length + 1;

			    XYZ pos = new XYZ(co*x*width, -co*y*height,0);

			    Curve c = null;

			    for(int i = 0; i < coordinates.Count; i++)
			    {
				c = Line.CreateBound(coordinates[i].Add(pos),coordinates[(i+1)%(coordinates.Count)].Add(pos));
				curves.Add(c);
			    }

			    loop = CurveLoop.Create(curves);
			    cloop.Add(loop);

			    FilledRegion.Create(doc,ft.Id,current.Id,cloop);
			    cloop.Clear();
			    TextNote.Create(doc,current.Id, pos.Add(new XYZ(0,-0.2,0)), 0.1,Truncate(ft.Name, 20),toptions);
			    counter ++;
			    s += counter.ToString();
			}  
			t.Commit();         
		    }                        
		}	
		public void ChangeTextFont()
		{
			Document doc = ActiveUIDocument.Document;
			
			FilteredElementCollector coll = new FilteredElementCollector(doc);
						
			var tntypes = new FilteredElementCollector(doc).OfClass(typeof(TextNoteType)).ToList();
			
			var tetypes = new FilteredElementCollector(doc).OfClass(typeof(TextElementType)).ToList();

			using(Transaction t = new Transaction(doc, "Change Font"))
			{
				t.Start();
				
				foreach(var tt in tntypes)
				{
					TextNoteType type = tt as TextNoteType;
					if(type == null) continue;
					
					type.LookupParameter("Text Font").Set("Gill Sans MT");
				}
				foreach(var te in tetypes)
				{
					TextElementType type = te as TextElementType;
					if(type == null) continue;
					
					type.LookupParameter("Text Font").Set("Gill Sans MT");
				}				
				t.Commit();
			}
		}
		
		
		public void SelectHosted()
		{	
		    UIDocument uidoc = ActiveUIDocument;
		    Document doc = ActiveUIDocument.Document;
		    Selection selection = uidoc.Selection;
		    ICollection<ElementId> selectedIds;

		    Reference refHost = selection.PickObject(ObjectType.Element, "Pick Host Object");

		    using(Transaction t = new Transaction(doc, "Fake"))
		    {
			t.Start();
			selectedIds = doc.Delete(refHost.ElementId);
			t.RollBack();
		    }

		    selection.SetElementIds(selectedIds);
		}
		public void RenameFilledRegionByTextNode()
		{
		    UIDocument uidoc = ActiveUIDocument;
		    Document doc = ActiveUIDocument.Document;
		    View current = doc.ActiveView;

		    string s = "";
		    do
		    {
			FilteredElementCollector collector = new FilteredElementCollector(doc);            

			    TextNote text = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Pick Text").ElementId) as TextNote;

			    if(!SelectionError(text)) return;

			    s = text.Text;

			    FilledRegion region = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Pick Filled Region").ElementId) as FilledRegion;

			    if(!SelectionError(text)) return;

			    using(Transaction t = new Transaction(doc,"Rename Filled Region"))
			    {
				t.Start();
				doc.GetElement(region.GetTypeId()).Name = s;
				t.Commit();
			    }     
		    }
		    while (true);
		}
		public void ReplaceSharedParameter()
		{
		    UIDocument uidoc = ActiveUIDocument;
		    Document doc = ActiveUIDocument.Document;
		    View current = doc.ActiveView;

		    FilteredElementCollector collector = new FilteredElementCollector(doc);        

			List<SharedParameterElement> sharedParamters = collector
				.OfClass(typeof(SharedParameterElement))
				.Cast<SharedParameterElement>()
				.OrderBy(x => x.Name)
				.ToList();

			string s = sharedParamters.Count.ToString();

			StringBuilder builder = new StringBuilder();

			var bindings = doc.ParameterBindings;

			BindingMap map = doc.ParameterBindings;
			DefinitionBindingMapIterator it = map.ForwardIterator();
			it.Reset();

			Definition def = null;

			while (it.MoveNext())
			{
				sharedParamters.RemoveAll(x => x.Name.Equals(it.Key.Name));
			}

			string e = sharedParamters.Count.ToString();

			builder.Append(String.Format("Number of all Shared Parameters: {0}{1}", s, Environment.NewLine));
			builder.Append(String.Format("Number of unused Shared Parameters: {0}{1}", e, Environment.NewLine));

			foreach(SharedParameterElement param in sharedParamters)
			{
				builder.Append(param.Name + Environment.NewLine);
			}
			
			using(TaskDialog td = new TaskDialog("Save unused parameter names."))
			{
				td.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No;
				td.DefaultButton = TaskDialogResult.Yes;
				
				TaskDialogResult result = td.Show();
				if(result == TaskDialogResult.Yes)
				{
				    using(Transaction t = new Transaction(doc, "FilledRegionPopulate"))
				    {
						t.Start();                    
						SaveText(builder);
						t.Commit();         
				    } 					
				}				
			}
						
			using(TaskDialog td = new TaskDialog("Delete unused Shared Paramters?"))
			{
				td.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No;
				td.DefaultButton = TaskDialogResult.Yes;
				td.FooterText = "Create a back-up before going forward might be a smart idea.";
				TaskDialogResult result = td.Show();
				if(result == TaskDialogResult.Yes)
				{
				    using(Transaction t = new Transaction(doc, "FilledRegionPopulate"))
				    {
						t.Start();                    
						doc.Delete(sharedParamters.Select(x => x.Id).ToArray());
						t.Commit();         
				    } 	
				    TaskDialog.Show("Result", String.Format("{0} paramters have been deleted.", e));
				}				
			}                 
		}
		public void DeleteElementIDs()
		{
    			Document doc = Document;
            
    			var file = LoadText();
			
			using(StreamReader reader = new StreamReader(file))
			{
				while(true)
				{
					string line = reader.ReadLine();
					if(line == null)
					{
						break;
					}
					try
					{
						using(Transaction t = new Transaction(doc,"del"))
						{
						    FailureHandlingOptions foptions = t.GetFailureHandlingOptions();
						    FailureHandler fhandler = new FailureHandler();
						    foptions.SetFailuresPreprocessor(fhandler);
						    foptions.SetClearAfterRollback(true);
						    t.SetFailureHandlingOptions(foptions);
				    
							t.Start();
							doc.Delete(new ElementId(int.Parse(line)));
							t.Commit();
						}
					}
					catch(Exception)
					{
						
					}
				}
			}
		}   
		
		public void SaveElementIDs()
		{
			Document doc = Document;
			
			Selection selection = this.Selection;
			
			StringBuilder builder = new StringBuilder();
			
			if(selection.GetElementIds().Count > 0)
			{
				foreach(ElementId id in selection.GetElementIds())
				{
					builder.AppendLine(id.ToString());
				}
				
				SaveText(builder);
			}						
		}
		
		internal void SaveText(StringBuilder builder)
		{
			using (System.Windows.Forms.SaveFileDialog dialog = new System.Windows.Forms.SaveFileDialog()) 
			{
				dialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*"  ;
				dialog.FilterIndex = 1 ;
				dialog.RestoreDirectory = true ;

			    if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK) 
			    {
			        File.WriteAllText(dialog.FileName, builder.ToString());
			    }
			}			
		}
		
		internal string LoadText()
		{
			string filename = "";	 
		    
			using(var ofd = new System.Windows.Forms.OpenFileDialog())
			{
			    System.Windows.Forms.DialogResult result = ofd.ShowDialog();
			
			    if (result == System.Windows.Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(ofd.FileName))
			    {
			        filename = ofd.FileName;
			    }
			}

			return filename;			
		}
		
		public void DeleteElementIDs()
		{
			UIDocument uidoc = ActiveUIDocument;
			Document doc = uidoc.Document;
            
    		var file = LoadText();
			
			using(StreamReader reader = new StreamReader(file))
			{
				while(true)
				{
					string line = reader.ReadLine();
					if(line == null)
					{
						break;
					}
					try
					{
						using(Transaction t = new Transaction(doc,"del"))
						{
							FailureHandlingOptions foptions = t.GetFailureHandlingOptions();
						    FailureHandler fhandler = new FailureHandler();
						    foptions.SetFailuresPreprocessor(fhandler);
						    foptions.SetClearAfterRollback(true);
						    t.SetFailureHandlingOptions(foptions);
				    
							t.Start();
							doc.Delete(new ElementId(int.Parse(line)));
							t.Commit();
						}
					}
					catch(Exception)
					{
						
					}
				}
			}
		}   
		
		public void SaveElementIDs()
		{
			UIDocument uidoc = ActiveUIDocument;
			Document doc = uidoc.Document;
			
			Selection selection = uidoc.Selection;
			
			StringBuilder builder = new StringBuilder();
			
			if(selection.GetElementIds().Count > 0)
			{
				foreach(ElementId id in selection.GetElementIds())
				{
					builder.AppendLine(id.ToString());
				}
				
				SaveText(builder);
			}						
		}
		
		internal void SaveText(StringBuilder builder)
		{
			using (System.Windows.Forms.SaveFileDialog dialog = new System.Windows.Forms.SaveFileDialog()) 
			{
				dialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*"  ;
				dialog.FilterIndex = 1 ;
				dialog.RestoreDirectory = true ;

			    if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK) 
			    {
			        File.WriteAllText(dialog.FileName, builder.ToString());
			    }
			}			
		}
		
		internal string LoadText()
		{
			string filename = "";	 
		    
			using(var ofd = new System.Windows.Forms.OpenFileDialog())
			{
			    System.Windows.Forms.DialogResult result = ofd.ShowDialog();
			
			    if (result == System.Windows.Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(ofd.FileName))
			    {
			        filename = ofd.FileName;
			    }
			}

			return filename;			
		}
		public void PDFMultipleDocuments()
		{			
			OpenFileDialog theDialogRevit = new OpenFileDialog();
			theDialogRevit.Title = "Select Revit Project Files";
			theDialogRevit.Filter = "RVT files|*.rvt";
			theDialogRevit.FilterIndex = 1;
			theDialogRevit.Multiselect = true;
			
			OpenOptions opt = new OpenOptions();
            opt.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets;
            opt.Audit = false;
            
			if (theDialogRevit.ShowDialog() == DialogResult.OK)
			{
                foreach (String projectPath in theDialogRevit.FileNames)
                {
                 	FileInfo filePath = new FileInfo(projectPath);
                 	string filename = filePath.Name;
                    ModelPath mp = ModelPathUtils.ConvertUserVisiblePathToModelPath(filePath.FullName);					
		            WorksetConfiguration wc = new WorksetConfiguration(WorksetConfigurationOption.OpenAllWorksets);
		            
		            try{
			            Document doc = Application.OpenDocumentFile(mp,opt);
			            PrintViewSets(doc);
			            doc.Close(false);		            	
		            }
		            catch(Exception)
		            {
		            	
		            }
                }
			}
		}
		private void PrintViewSets(Document doc)
		{
			List<ViewSheetSet> viewSets = new FilteredElementCollector(doc).OfClass(typeof(ViewSheetSet)).Cast<ViewSheetSet>().ToList();			

			// No ViewSets No Game			
			if(viewSets.Count == 0) 
			{
				TaskDialog.Show("Error", String.Format("No ViewSets in {0} found.", doc.Title));
				return;
			}
			
			foreach(ViewSheetSet vset in viewSets)
			{
				using (Transaction t = new Transaction(doc, "Print Test"))
				{
					t.Start();
					doc.Print(vset.Views);				
					t.Commit();
				}				
			}			
		}
		public void CreateWorksets()
		{
			UIDocument uidoc = ActiveUIDocument;
			Document doc = uidoc.Document;
			
			string [] names = {
				"GHA-Walls",
				"GHA-Doors",
				"GHA-Floors",
				"GHA-Ceilings",
				"GHA-Furniture",
				"GHA-Stairs & Railings",
				"GHA-Structure",
				"GHA-Windows",
				"GHA-Site",
				"X-Cad",
				"X-Admin",
				"X-Links",				
			};
			
				CreateWorkset(doc, "test");
				
			foreach(string n in names)
			{
				CreateWorkset(doc, n);
			}
		}
		
		public void PopulateFamily()
		{
			UIDocument uidoc = ActiveUIDocument;
			Document doc = uidoc.Document;
			
			FamilyManager manager = doc.FamilyManager;
				
			BuiltInParameterGroup addToGroup = BuiltInParameterGroup.PG_VISIBILITY;
			ParameterType paramType = ParameterType.YesNo;
						
			Selection selection = uidoc.Selection;
			ElementId current = doc.ActiveView.Id;
			
			TextNote tbox = doc.GetElement(selection.PickObject(ObjectType.Element, "Pick Text Box").ElementId) as TextNote;
			
			XYZ pos = (tbox.get_BoundingBox(doc.ActiveView).Max + tbox.get_BoundingBox(doc.ActiveView).Min) *0.5;
			var width = tbox.Width + 0.01;
			TextNoteOptions opt = new TextNoteOptions();
		    opt.TypeId = doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType);
			
			Dictionary<string, string> signs = new Dictionary<string, string>()
			{
				{"S13", "Fire Door Keep Shut"},
				{"S14", "Fire Door Keep Locked"},
				{"S16", "Automatic Fire Door, Keep Clear at Night"},
				{"S20", "Fire Escape Keep Clear"},
				{"S22", "Fire Exit"},
				{"S23", "Slide to Open"},
				{"S25", "Push Bar to Open"},
				{"S26", "Directional Arrow(Green)"},
				{"S27", "Collection of Fire Fighting Equipment"},
				{"S28", "Fire Alarm Call Point"},
				{"S30", "Fire Hose Reel"},
				{"S31", "Fire Extinguisher"},
				{"S33", "Dry Riser"},
				{"S35", "Firemans Switch"},
				{"S38", "Fire Plan"},
				{"S39", "Directional Arrow(Red)"},
			};
			
			using(Transaction t = new Transaction(doc, "Populate Parameters"))
			{
				t.Start();		
				foreach(KeyValuePair<string, string> pair in signs)
				{					
					TextNote note = TextNote.Create(doc, current, pos, width, pair.Value, opt);					
					FamilyParameter famParam = doc.FamilyManager.AddParameter(pair.Key + " Bottom", addToGroup, paramType, true);
					Parameter param = note.LookupParameter("Visible");
					manager.AssociateElementParameterToFamilyParameter(param, famParam);
					//manager.Set(famParam, 0);
					manager.SetFormula(famParam, String.Format("and({0}, {1})", pair.Key, "Bottom Label"));					
				}
				t.Commit();
			}					
		}
		
		public void DeleteSelectedType()
		{
			UIDocument uidoc = ActiveUIDocument;
			Document doc = uidoc.Document;
			
			Selection selection = uidoc.Selection;
			
			TextNote tbox = doc.GetElement(selection.PickObject(ObjectType.Element, "Pick Text Box").ElementId) as TextNote;
			
			if (tbox == null) return;
			
			using(Transaction t = new Transaction(doc, "Delete Type"))
			{
				t.Start();
				doc.Delete(tbox.GetTypeId());				
				t.Commit();
			}
			
		}
		
		internal Workset CreateWorkset(Document document, string name)
		{
			Workset newWorkset = null;
			// Worksets can only be created in a document with worksharing enabled
			if (document.IsWorkshared)
			{
				string worksetName = "New Workset";
				// Workset name must not be in use by another workset
				if (WorksetTable.IsWorksetNameUnique(document, worksetName))
				{
					using (Transaction worksetTransaction = new Transaction(document, "Set preview view id"))
					{
						worksetTransaction.Start();
						newWorkset = Workset.Create(document, worksetName);
						worksetTransaction.Commit();
					}
				}
			}
			return newWorkset;
		}
		
		
		
		public void DeleteUnusedParameter()
		{
			UIDocument uidoc = ActiveUIDocument;
			Document doc = uidoc.Document;
		    View current = doc.ActiveView;

		    FilteredElementCollector collector = new FilteredElementCollector(doc);        
			    		    
			List<ParameterElement> parameters = collector
				.OfClass(typeof(ParameterElement))
				.Cast<ParameterElement>()
				.OrderBy(x => x.Name)
				.ToList();

			string s = parameters.Count.ToString();

			StringBuilder builder = new StringBuilder();

			var bindings = doc.ParameterBindings;

			BindingMap map = doc.ParameterBindings;
			DefinitionBindingMapIterator it = map.ForwardIterator();
			it.Reset();

			
			while (it.MoveNext())
			{
				parameters.RemoveAll(x => x.Name.Equals(it.Key.Name));
			}
			
			string e = parameters.Count.ToString();

			builder.Append(String.Format("Number of all Parameters: {0}{1}", s, Environment.NewLine));
			builder.Append(String.Format("Number of unused Parameters: {0}{1}", e, Environment.NewLine));

			foreach(ParameterElement param in parameters)
			{
				builder.Append(param.Name + " : " + param.Id + Environment.NewLine);
			}
			
			using(TaskDialog td = new TaskDialog("Save unused parameter names."))
			{
				td.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No;
				td.DefaultButton = TaskDialogResult.Yes;
				
				TaskDialogResult result = td.Show();
				if(result == TaskDialogResult.Yes)
				{               
					SaveText(builder);	
				}				
			}
			
			int deleted = 0;	
			
			using(TaskDialog td = new TaskDialog("Delete unused Paramters?"))
			{
				td.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No;
				td.DefaultButton = TaskDialogResult.Yes;
				td.FooterText = "Create a back-up before going forward might be a smart idea.";
				TaskDialogResult result = td.Show();
				if(result == TaskDialogResult.Yes)
				{
					foreach(var param in parameters)
					{
					
					    using(Transaction t = new Transaction(doc, "FilledRegionPopulate"))
					    {
		        			FailureHandlingOptions options = t.GetFailureHandlingOptions();
					    	FailureHandler failureHandler = new FailureHandler();
			                failureHandler = new FailureHandler();
			                options.SetFailuresPreprocessor(failureHandler);
			                options.SetClearAfterRollback(true);
			                t.SetFailureHandlingOptions(options);	
	                
							t.Start();                    
							doc.Delete(param.Id);
							t.Commit();         
							deleted ++;
					    } 	
					}
			    	TaskDialog.Show("Result", String.Format("{0} paramters have been deleted.", deleted.ToString()));
				}				
			}			           
		}		
		internal void SaveText(StringBuilder builder)
		{
			using (System.Windows.Forms.SaveFileDialog dialog = new System.Windows.Forms.SaveFileDialog()) 
			{
				dialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*"  ;
				dialog.FilterIndex = 1 ;
				dialog.RestoreDirectory = true ;

			    if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK) 
			    {
			        File.WriteAllText(dialog.FileName, builder.ToString());
			    }
			}			
		}
		
		public void ShowLevelInfo()
		{
			UIDocument uidoc = ActiveUIDocument;
			Document doc = uidoc.Document;
			Getinfo_Level(doc);
		}
		
		private void Getinfo_Level(Document document)
		{
	        StringBuilder levelInformation = new StringBuilder();
	        int levelNumber = 0;
	        FilteredElementCollector collector = new FilteredElementCollector(document);
	        ICollection<Element> collection = collector.OfClass(typeof(Level)).ToElements();
	        foreach (Element e in collection)
	        {
	                Level level = e as Level;
	        
	                if (null != level)
	                {
	                        // keep track of number of levels
	                        levelNumber++;
	                
	                        //get the name of the level
	                        levelInformation.Append("\nLevel Name: " + level.Name);
	
	                        //get the elevation of the level
	                        levelInformation.Append("\n\tElevation: " + level.Elevation);
	                
	                        // get the project elevation of the level
	                        levelInformation.Append("\n\tProject Elevation: " + level.ProjectElevation);
	                }
	        }
	
	        //number of total levels in current document
	        levelInformation.Append("\n\n There are " + levelNumber + " levels in the document!");
	        
	        //show the level information in the messagebox
	        TaskDialog.Show("Revit",levelInformation.ToString());
		}	
		
		public void DuplicateSheet()
		{
			Document doc = this.ActiveUIDocument.Document;
			Selection sel = this.ActiveUIDocument.Selection;			
			
			//ViewSheet duplicateSheet = SheetToDuplicate(doc);
			
			List<ViewSheet> sheetsToDuplicate = SheetToDuplicateFromSelection(this.ActiveUIDocument);
			
			int times = 0;
			//Int32.TryParse(Prompt.ShowDialog("Which Sheet Number?", "o.O"), out times);	//sheet number is unique, whereas its name is not
			using(TransactionGroup tgroup = new TransactionGroup(doc, "Duplicate Sheets"))
			{
				tgroup.Start();
				foreach(var sheet in sheetsToDuplicate)
				{	
					times ++;				
					ViewSheet newSheet = DuplicateViewSheet(doc, times.ToString(), sheet);
				}		
				tgroup.Assimilate();
			}
		}
		public void DuplicateSheetAndViews()
		{
			Document doc = this.ActiveUIDocument.Document;
			Selection sel = this.ActiveUIDocument.Selection;
			
			int counter = 0;		
			
//			ViewSheet duplicateSheet = SheetToDuplicate(doc);
			
			List<ViewSheet> sheetsToDuplicate = SheetToDuplicateFromSelection(this.ActiveUIDocument);			
			using(TransactionGroup tgroup = new TransactionGroup(doc, "Duplicate Sheets"))
			{
				tgroup.Start();
				foreach(var sheet in sheetsToDuplicate)
				{
					ViewSheet newSheet = DuplicateViewSheet(doc, "_" + counter.ToString(), sheet);
				
					var placedViews = sheet.GetAllViewports();			
					
					Dictionary<Autodesk.Revit.DB.View, XYZ> viewsOnSheet = new Dictionary<Autodesk.Revit.DB.View, XYZ>();					
											
					foreach(var view in placedViews)
					{
						Viewport vport = doc.GetElement(view) as Viewport;
						Autodesk.Revit.DB.View v = doc.GetElement(vport.ViewId) as Autodesk.Revit.DB.View;
						
						if(v.ViewType == ViewType.Legend) 
						{
							viewsOnSheet.Add(v, vport.GetBoxCenter());
						}				
						else
						{
							viewsOnSheet.Add(DuplicateView(doc, v), vport.GetBoxCenter());
						}
					}
								
					if(newSheet == null) 
					{
						TaskDialog.Show("Error", "Something Failed");
						return;
					}
					
					using(Transaction t = new Transaction(doc, "Add Views On Sheet"))
					{
						t.Start();
						foreach(var v in viewsOnSheet)
						{
							if(Viewport.CanAddViewToSheet(doc, newSheet.Id, v.Key.Id))
							{
								Viewport.Create(doc, newSheet.Id, v.Key.Id, v.Value);
							}
						}
						t.Commit();
					}
					counter ++;
				}
				tgroup.Assimilate();
			}			
		}
		internal ViewSheet DuplicateViewSheet(Document doc, string increment, ViewSheet duplicateSheet)
		{
			
			ElementId tblock = null;
			
			foreach (var element in new FilteredElementCollector(doc, duplicateSheet.Id))
			{
				if(element.Category.Name.ToString().Equals("Title Blocks"))
				{
					tblock = element.GetTypeId();
				}
			}	
			
			ViewSheet newSheet = null;
			
			using(Transaction t = new Transaction(doc, "Duplicate Sheet"))
			{
				t.Start();				
				newSheet = ViewSheet.Create(doc, tblock);
				newSheet.Name = duplicateSheet.Name;
				newSheet.SheetNumber = duplicateSheet.SheetNumber + increment;
				newSheet.LookupParameter("Sheet Group").Set(duplicateSheet.LookupParameter("Sheet Group").AsString());
				t.Commit();				
			}
			
			return newSheet;
		}
		internal ViewSheet SheetToDuplicate(Document doc)
		{	
			string sheetNumber = Prompt.ShowDialog("Which Sheet Number?", "o.O");	//sheet number is unique, whereas its name is not
			
			ViewSheet duplicateSheet = new FilteredElementCollector(doc)
				.WhereElementIsNotElementType()
				.OfClass(typeof(ViewSheet))
				.Cast<ViewSheet>()
				.First(x => x.SheetNumber.Equals(sheetNumber));
			
			return duplicateSheet;
		}
		internal List<ViewSheet> SheetToDuplicateFromSelection(UIDocument uidoc)
		{	
			var selectedSheets = uidoc.Selection.GetElementIds().Select(x => uidoc.Document.GetElement(x) as ViewSheet).ToList();
						
			return selectedSheets;
		}
		internal Autodesk.Revit.DB.View DuplicateView(Document doc, Autodesk.Revit.DB.View v)
		{
			Autodesk.Revit.DB.View duplicate = null;
			
			using(Transaction t = new Transaction(doc, "Duplicate View"))
			{
				t.Start();
				duplicate = doc.GetElement(v.Duplicate(ViewDuplicateOption.Duplicate)) as Autodesk.Revit.DB.View;
				t.Commit();
			}			
			
			return duplicate;
		}
		
		public void DuplicateViewsByParamter()
		{
			UIDocument uidoc = ActiveUIDocument;
			Document doc = uidoc.Document;
			
			string parameterName = "View Category";
			string parameterValue = "9 Originals";
			
			List<View> views = new FilteredElementCollector(doc)
				.OfCategory(BuiltInCategory.OST_Views)
				.WhereElementIsNotElementType()
				.Where(x => x.LookupParameter(parameterName).AsString() != null &&
				       x.LookupParameter(parameterName).AsString().Contains(parameterValue))
				.Cast<View>()
				.Where(x => !x.IsTemplate)
				.ToList();
						
			using(Transaction t = new Transaction(doc, "Duplicate Views"))
			{
				t.Start();
				foreach(View v in views)
				{
					v.Duplicate(ViewDuplicateOption.Duplicate);
				}
				t.Commit();
			}
		}
		
		public void DeleteViewsAndSheets()
		{
			UIDocument uidoc = ActiveUIDocument;
			Document doc = uidoc.Document;
			
			List<ElementId> viewsToDelete = new FilteredElementCollector(doc)
				.OfCategory(BuiltInCategory.OST_Views)
				.WhereElementIsNotElementType()
				.Where(x => x.LookupParameter("View Category").AsString() != null &&
				       !x.LookupParameter("View Category").AsString().Contains("9 Originals"))
				.Select(x => x.Id)
				.ToList();
			
			string [] groups = {
				"ADMIN",
				"SPLASH SCREEN"
			};
			
			List<ElementId> sheetsToDelete = new FilteredElementCollector(doc)
				.OfCategory(BuiltInCategory.OST_Sheets)
				.WhereElementIsNotElementType()
				.Where(x => x.LookupParameter("Sheet Group").AsString() != null &&
				       !groups.Contains(x.LookupParameter("Sheet Group").AsString()))
				.Select(x => x.Id)
				.ToList();
			
			using(Transaction t = new Transaction(doc, "Delete Views and Sheets"))
			{
				t.Start();
				doc.Delete(viewsToDelete);
				doc.Delete(sheetsToDelete);
				t.Commit();
			}
			
		}
		public void Overkill()
		{
			UIDocument uidoc = this.ActiveUIDocument;
			Document doc = uidoc.Document;
			
			LineSelectionFilter filter = new ThisApplication.LineSelectionFilter(doc);
			IList<Reference> references = uidoc.Selection.PickObjects(ObjectType.Element, filter, "Select multiple lines");
			
			Dictionary<Reference, Curve> detail_lines = new Dictionary<Reference, Curve>();			
			
			Dictionary<Reference, Curve> model_lines = new Dictionary<Reference, Curve>();
			
			Dictionary<Reference, Curve> room_lines = new Dictionary<Reference, Curve>();			
			
			Dictionary<Reference, Curve> area_lines = new Dictionary<Reference, Curve>();
			
			List<ElementId> survivor = new List<ElementId>();
			List<ElementId> casualty = new List<ElementId>();
			
			foreach(Reference r in references)
			{				
				if ((doc.GetElement(r).Location as LocationCurve).Curve is Line)
				{
					if(doc.GetElement(r).Name == "Detail Lines")
					{
						detail_lines.Add(r, (doc.GetElement(r).Location as LocationCurve).Curve);						
					}	
					else
					{
						if(doc.GetElement(r).Category.Name == "Lines")
						{
							model_lines.Add(r, (doc.GetElement(r).Location as LocationCurve).Curve);
						}
						else if(doc.GetElement(r).Category.Name == "<Room Separation>")
						{
							room_lines.Add(r, (doc.GetElement(r).Location as LocationCurve).Curve);
						}
						else if(doc.GetElement(r).Category.Name == "<Area Boundary>")
						{
							area_lines.Add(r, (doc.GetElement(r).Location as LocationCurve).Curve);
						}
					}
				}
			}
						
			itterate(detail_lines, survivor, casualty, doc);
			itterate(model_lines, survivor, casualty, doc);
			itterate(room_lines, survivor, casualty, doc);
			itterate(area_lines, survivor, casualty, doc);
			
			
			using (Transaction t = new Transaction(doc,"Overkill"))
            {
                t.Start();
                doc.Delete(casualty);
                t.Commit();
            }
			TaskDialog.Show("Overkill", "Number of lines deleted: " + casualty.Count);
		}
		
		private void itterate(Dictionary<Reference, Curve> lines, List<ElementId> survivor, List<ElementId> casualty, Document doc)
		{			
			int c = 0;
			
			foreach(KeyValuePair<Reference, Curve> a in lines)
			{
//				TaskDialog.Show("test", "aId " + c.ToString() + " " + a.Key.ElementId);
				
				if (casualty.Contains(a.Key.ElementId))
				{
					c++;
					continue;
				}
				
				foreach(KeyValuePair<Reference, Curve> b in lines)
				{				
					if (a.Equals(b))
					{
						c++;
						continue;
					}
					
					if(casualty.Contains(b.Key.ElementId))
					{
						c++;
						continue;
					}
					
					Tuple<KeyValuePair<Reference, Curve>, KeyValuePair<Reference, Curve>> l_tuple = Overlap(a, b, doc);
					
					if (null != l_tuple)
					{
						survivor.Add(l_tuple.Item1.Key.ElementId);
						casualty.Add(l_tuple.Item2.Key.ElementId);
						continue;
					}
					else
					{
						survivor.Add(b.Key.ElementId);
					}					
				}
			}			
			
//			TaskDialog.Show("test", "iterations " + c.ToString() + " casualties " + casualty.Count);
		}
				
		private Tuple<KeyValuePair<Reference, Curve>, KeyValuePair<Reference, Curve>> Overlap(KeyValuePair<Reference, Curve> l1, KeyValuePair<Reference, Curve> l2, Document doc)
		{						
			double precision = 0.0001;
			
			XYZ l1p1 = l1.Value.GetEndPoint(0);
			XYZ l1p2 = l1.Value.GetEndPoint(1);
			XYZ l2p1 = l2.Value.GetEndPoint(0);
			XYZ l2p2 = l2.Value.GetEndPoint(1);
			
			XYZ v1 = l1p1 - l1p2;
			XYZ v2 = l2p1 - l2p2;
			
			XYZ check = l1p1 - l2p2;
			
			XYZ cross = v1.CrossProduct(v2);
			XYZ cross_check = v1.CrossProduct(check);
			
//			TaskDialog.Show("Overkill", "Cross + cross_check: " + cross + ":" + cross_check);
			
			XYZ min1 = new XYZ(Math.Min(l1p1.X, l1p2.X), 
			                   Math.Min(l1p1.Y, l1p2.Y),
			                   Math.Min(l1p1.Z, l1p2.Z));
			
			XYZ max1 = new XYZ(Math.Max(l1p1.X, l1p2.X), 
			                   Math.Max(l1p1.Y, l1p2.Y),
			                   Math.Max(l1p1.Z, l1p2.Z));			
			
			XYZ min2 = new XYZ(Math.Min(l2p1.X, l2p2.X), 
			                   Math.Min(l2p1.Y, l2p2.Y),
			                   Math.Min(l2p1.Z, l2p2.Z));
			
			XYZ max2 = new XYZ(Math.Max(l2p1.X, l2p2.X), 
			                   Math.Max(l2p1.Y, l2p2.Y),
			                   Math.Max(l2p1.Z, l2p2.Z));
			
//			XYZ minIntersect = new XYZ(Math.Max(min1.X, min2.X), Math.Max(min1.Y, min2.Y), Math.Max(min1.Z, min2.Z));
//			XYZ maxIntersect = new XYZ(Math.Min(max1.X, max2.X), Math.Min(max1.Y, max2.Y), Math.Min(max1.Z, max2.Z));
//			
//			XYZ minEnd = new XYZ(Math.Min(min1.X, min2.X), Math.Min(min1.Y, min2.Y), Math.Min(min1.Z, min2.Z));
//			XYZ maxEnd = new XYZ(Math.Max(max1.X, max2.X), Math.Max(max1.Y, max2.Y), Math.Max(max1.Z, max2.Z));
//			
//			bool intersect = (minIntersect.X <= maxIntersect.X) && (minIntersect.Y <= maxIntersect.Y) && (minIntersect.Z <= maxIntersect.Z);
			
			bool contain = (((min1.X - min2.X) <= precision) && (max1.X - max2.X) >= -precision && (min1.Y - min2.Y) <= precision && (max1.Y - max2.Y) >= -precision && (min1.Z - min2.Z) <= precision &&  (max1.Z - max2.Z) >= -precision) ||
				(((min1.X - min2.X) >= -precision) && (max1.X - max2.X) <= precision && (min1.Y - min2.Y) >= -precision && (max1.Y - max2.Y) <= precision && (min1.Z - min2.Z) >= -precision && (max1.Z - max2.Z) <= precision);
			
//			TaskDialog.Show("test", "check " + intersect + ":" + minIntersect + ":" + maxIntersect);
								
//			TaskDialog.Show("Overkill", "contain: " + contain);
			
			if(cross.IsZeroLength() && cross_check.IsZeroLength() && contain)
			{
//				using (Transaction t = new Transaction(doc,"Overkill"))
//	            {
//	                t.Start();
//	                (doc.GetElement(l1.Key).Location as LocationCurve).Curve = Line.CreateBound(maxEnd, minEnd);
//	                t.Commit();
//	            }				
//				return Tuple.Create(l1,l2);
				
				return (v1.GetLength() > v2.GetLength()) ? Tuple.Create(l1, l2) : Tuple.Create(l2, l1);							
			}
			
			else return null;
		}
	
		private Tuple<ElementId, ElementId> OverlapId(ElementId l1, ElementId l2, Document doc)
		{						
			double precision = 0.01;			
						
			Curve c1 = (doc.GetElement(l1).Location as LocationCurve).Curve as Curve;
			Curve c2 = (doc.GetElement(l2).Location as LocationCurve).Curve as Curve;
//			
//			TaskDialog.Show("OffAxis", "here ");
						
			XYZ l1p1 = c1.GetEndPoint(0);
			XYZ l1p2 = c1.GetEndPoint(1);
			XYZ l2p1 = c2.GetEndPoint(0);
			XYZ l2p2 = c2.GetEndPoint(1);
			
			XYZ v1 = l1p1 - l1p2;
			XYZ v2 = l2p1 - l2p2;
			
			XYZ check = l1p1 - l2p2;
			
			XYZ cross = v1.CrossProduct(v2);
			XYZ cross_check = v1.CrossProduct(check);
			
//			TaskDialog.Show("Overkill", "Cross + cross_check: " + cross + ":" + cross_check);
			
			XYZ min1 = new XYZ(Math.Min(l1p1.X, l1p2.X), 
			                   Math.Min(l1p1.Y, l1p2.Y),
			                   Math.Min(l1p1.Z, l1p2.Z));
			
			XYZ max1 = new XYZ(Math.Max(l1p1.X, l1p2.X), 
			                   Math.Max(l1p1.Y, l1p2.Y),
			                   Math.Max(l1p1.Z, l1p2.Z));			
			
			XYZ min2 = new XYZ(Math.Min(l2p1.X, l2p2.X), 
			                   Math.Min(l2p1.Y, l2p2.Y),
			                   Math.Min(l2p1.Z, l2p2.Z));
			
			XYZ max2 = new XYZ(Math.Max(l2p1.X, l2p2.X), 
			                   Math.Max(l2p1.Y, l2p2.Y),
			                   Math.Max(l2p1.Z, l2p2.Z));
			
//			XYZ minIntersect = new XYZ(Math.Max(min1.X, min2.X), Math.Max(min1.Y, min2.Y), Math.Max(min1.Z, min2.Z));
//			XYZ maxIntersect = new XYZ(Math.Min(max1.X, max2.X), Math.Min(max1.Y, max2.Y), Math.Min(max1.Z, max2.Z));
//			
//			XYZ minEnd = new XYZ(Math.Min(min1.X, min2.X), Math.Min(min1.Y, min2.Y), Math.Min(min1.Z, min2.Z));
//			XYZ maxEnd = new XYZ(Math.Max(max1.X, max2.X), Math.Max(max1.Y, max2.Y), Math.Max(max1.Z, max2.Z));
//			
//			bool intersect = (minIntersect.X <= maxIntersect.X) && (minIntersect.Y <= maxIntersect.Y) && (minIntersect.Z <= maxIntersect.Z);
			
			bool contain = (((min1.X - min2.X) <= precision) && (max1.X - max2.X) >= -precision && (min1.Y - min2.Y) <= precision && (max1.Y - max2.Y) >= -precision && (min1.Z - min2.Z) <= precision &&  (max1.Z - max2.Z) >= -precision) ||
				(((min1.X - min2.X) >= -precision) && (max1.X - max2.X) <= precision && (min1.Y - min2.Y) >= -precision && (max1.Y - max2.Y) <= precision && (min1.Z - min2.Z) >= -precision && (max1.Z - max2.Z) <= precision);
			
//			TaskDialog.Show("test", "check " + intersect + ":" + minIntersect + ":" + maxIntersect);
								
//			TaskDialog.Show("Overkill", "contain: " + contain);
			
			if(cross.IsZeroLength() && cross_check.IsZeroLength() && contain)
			{
//				using (Transaction t = new Transaction(doc,"Overkill"))
//	            {
//	                t.Start();
//	                (doc.GetElement(l1.Key).Location as LocationCurve).Curve = Line.CreateBound(maxEnd, minEnd);
//	                t.Commit();
//	            }				
//				return Tuple.Create(l1,l2);
				
				return (v1.GetLength() > v2.GetLength()) ? Tuple.Create(l1, l2) : Tuple.Create(l2, l1);							
			}
			
			else return null;
		}
		/// <summary>
		/// Attempts to remove overlapping Lines. Will only remove complete overlap or if line is subset to another line
		/// </summary>
		public void OverlappingLines()
		{			
			Document doc = this.ActiveUIDocument.Document;
			
			List<ElementId> survivor = new List<ElementId>();
			List<ElementId> casualty = new List<ElementId>();
			
		    string filename = "";		    
		    
			using(var ofd = new OpenFileDialog())
			{
			    DialogResult result = ofd.ShowDialog();
			
			    if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(ofd.FileName))
			    {
			        filename = ofd.FileName;
			    }
			}
		
		    IList<ElementId> ids = new List<ElementId>();		    
			
		    using (StreamReader sr = new StreamReader(filename))
		    {
		        string line = "";
		        
		        while ((line = sr.ReadLine()) != null)
		        {
		        	var firstIndex = line.IndexOf("Model Line");
		        	
		        	if (firstIndex != line.LastIndexOf("Model Line") && firstIndex != -1)
	        	    {		        		
	        	    	string[] row = System.Text.RegularExpressions.Regex.Split(line,"id ");
	            
			            if (row.Length != 3)
			            {
			                continue;
			            }
			            						
			            string id1 = row[1].Split(' ')[0];
			            string id2 = row[2].Split(' ')[0];
			
			            ElementId elid1 = (new ElementId(int.Parse(id1)));
			            ElementId elid2 = (new ElementId(int.Parse(id2)));			            
			            
			            Tuple<ElementId, ElementId> l_tuple = OverlapId(elid1, elid2, doc);					            
						
						if (null != l_tuple)
						{
							survivor.Add(l_tuple.Item1);
							casualty.Add(l_tuple.Item2);
						}
	        	    }	       		            
		        }
		    }    
		
			
		    using (Transaction t = new Transaction(doc,"Overkill"))
            {
                t.Start();
                doc.Delete(casualty);
                t.Commit();
            }
			TaskDialog.Show("Overkill", "Number of lines deleted: " + casualty.Count);
		}	
		/// <summary>
		/// Re-centers room tags to their hosts
		/// </summary>
		public void RoomAreaTags()
		{
			Document doc = this.ActiveUIDocument.Document;
			
			List<ElementId> survivor = new List<ElementId>();
			List<ElementId> casualty = new List<ElementId>();
			
		    string filename = "";
		    
			using(var ofd = new OpenFileDialog())
			{
			    DialogResult result = ofd.ShowDialog();
			
			    if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(ofd.FileName))
			    {
			        filename = ofd.FileName;
			    }
			}
			
		    string message = "";
		    
		    IList<ElementId> ids = new List<ElementId>();		    
		    
		    bool next = false;
		    
		    List<Tuple<RoomTag, XYZ>> rtSet = new List<Tuple<RoomTag, XYZ>>();
		    List<Tuple<AreaTag, XYZ>> atSet = new List<Tuple<AreaTag, XYZ>>();
            
		    using (StreamReader sr = new StreamReader(filename))
		    {
		        string line = "";
		        
		        while ((line = sr.ReadLine()) != null)
		        {
		        	if (next)
		        	{
		        		next = false;
		        				        		
		        		string[] row = System.Text.RegularExpressions.Regex.Split(line,"id ");
		        					            			            
			            string id = row[1].Split(' ')[0];
						
			            //We found the faulty room tag!
			            Element element = doc.GetElement(new ElementId(int.Parse(id)));	
						
			            RoomTag roomTag = element as RoomTag;	
			            if(roomTag != null)
			            {			            
				            try
				            {	
				            	XYZ translation = (roomTag.Room.Location as LocationPoint).Point - (roomTag.Location as LocationPoint).Point;	
				            	rtSet.Add(new Tuple<RoomTag, XYZ>(roomTag, translation));
			            		message += String.Format("RoomTag {0} was moved to its Room Host.{1}", id, Environment.NewLine);
				            }
				            catch(Exception ex)
				            {
				            	throw ex;
				            }			            	
			            }
			            
			            AreaTag areaTag = element as AreaTag;
			            if(areaTag != null)
			            {		            
				            try
				            {	
				            	XYZ translation = (areaTag.Area.Location as LocationPoint).Point - (areaTag.Location as LocationPoint).Point;	
				            	atSet.Add(new Tuple<AreaTag, XYZ>(areaTag, translation));
			            		message += String.Format("AreaTag {0} was moved to its Room Host.{1}", id, Environment.NewLine);
				            }
				            catch(Exception ex)
				            {
				            	throw ex;
				            }				            	
			            }
		        	}
		        	else
		        	{
		        		var firstIndexRoomTag = line.IndexOf("Room Tag");
		        	    var firstIndexAreaTag = line.IndexOf("Area Tag");
		        	    
			        	if (firstIndexRoomTag != line.LastIndexOf("Room Tag") && firstIndexRoomTag != -1)
		        	    {	     		
							next = true;		        	    	
		        	    }	 
			        	else if(firstIndexAreaTag != line.LastIndexOf("Room Tag") && firstIndexAreaTag != -1)
			        	{
			        		next = true;
			        	}
		        	}		        	      		            
		        }
		        if(rtSet.Count > 0)
		        {
			        using (Transaction t = new Transaction(doc,"RoomTag_AssignLeader"))
		            {
		                t.Start();
		                foreach(Tuple<RoomTag, XYZ> tuple in rtSet)
		                {
		                	tuple.Item1.HasLeader = true;
	        				//tuple.Item1.Location.Move(tuple.Item2);	                	
		                }				                
		                t.Commit();
		            }
			        using (Transaction t = new Transaction(doc,"RoomTag_RemoveLeader"))
		            {
		                t.Start();
		                foreach(Tuple<RoomTag, XYZ> tuple in rtSet)
		                {
		                	tuple.Item1.HasLeader = false;
	        				//tuple.Item1.Location.Move(tuple.Item2);	                	
		                }				                
		                t.Commit();
		            }		        	
		        }
		        if(atSet.Count > 0)
		        {
			        using (Transaction t = new Transaction(doc,"AreaTag_AssignLeader"))
		            {
		                t.Start();
		                foreach(Tuple<AreaTag, XYZ> tuple in atSet)
		                {
		                	tuple.Item1.HasLeader = true;
	        				//tuple.Item1.Location.Move(tuple.Item2);	                	
		                }				                
		                t.Commit();
		            }
			        using (Transaction t = new Transaction(doc,"AreaTag_RemoveLeader"))
		            {
		                t.Start();
		                foreach(Tuple<AreaTag, XYZ> tuple in atSet)
		                {
		                	tuple.Item1.HasLeader = false;
	        				//tuple.Item1.Location.Move(tuple.Item2);	                	
		                }				                
		                t.Commit();
		            }			        	
		        }
		        
		        message += String.Format("{0}{1}Overall {2} Room Tags and {3} Area Tags have moved.", Environment.NewLine, Environment.NewLine, rtSet.Count, atSet.Count);
			}	    		
			
			TaskDialog.Show("RoomTag", message);
		}
		/// <summary>
		/// Removes Identical Instances Warning by deleting one of the duplicating elements
		/// </summary>
		public void IdenticalInstances()
		{
			Document doc = this.ActiveUIDocument.Document;
			Selection selection = this.ActiveUIDocument.Selection;
			
			List<ElementId> survivor = new List<ElementId>();
			List<ElementId> casualty = new List<ElementId>();
			
		    string filename = "";
		    
			using(var ofd = new OpenFileDialog())
			{
			    DialogResult result = ofd.ShowDialog();
			
			    if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(ofd.FileName))
			    {
			        filename = ofd.FileName;
			    }
			}
			
		    string message = "";
		    
		    IList<ElementId> ids = new List<ElementId>();		    
		    
		    bool next = false;
		                
		    using (StreamReader sr = new StreamReader(filename))
		    {
		        string line = "";
		        
		        while ((line = sr.ReadLine()) != null)
		        {		        	
		        	if (next)
		        	{
		        		next = false;
		        				        		
		        		string[] row = System.Text.RegularExpressions.Regex.Split(line,"id ");
						
			            string id1 = row[1].Split(' ')[0];	
			            string id2 = row[2].Split(' ')[0];	
			            			     
			            survivor.Add(new ElementId(int.Parse(id1)));
			            casualty.Add(new ElementId(int.Parse(id2)));
			            
			            message += String.Format("Element Id {0}{1}",id1, Environment.NewLine);
		        	}
		        	else
		        	{
		        		var firstIndex = line.IndexOf("identical instances");
		        	
			        	if (firstIndex != -1)
		        	    {	    	
							next = true;		        	    	
		        	    }	 
		        	}		        	      		            
		        }
		        if (casualty.Count == 0)
		        {
		        	message = "No such warning.";
		        	TaskDialog.Show("RoomTag", message);
		        	
		        	return;
		        }
		        
		        FilteredElementCollector col = new FilteredElementCollector(doc, doc.ActiveView.Id).WhereElementIsNotElementType();
            	ICollection<ElementId> allAlements = col.ToElementIds();
            	
            	//TaskDialog.Show("IdenticalInstances", casualty.Count.ToString());
            	casualty.RemoveAll(item => allAlements.Contains(item));
            	//TaskDialog.Show("IdenticalInstances", casualty.Count.ToString());            	
            	
				using (Transaction t = new Transaction(doc,"IdenticalInstances_RemoveElements"))
				{		
					FailureHandlingOptions foptions = t.GetFailureHandlingOptions();
				    FailureHandler fhandler = new FailureHandler();
				    foptions.SetFailuresPreprocessor(fhandler);
				    foptions.SetClearAfterRollback(true);
				    t.SetFailureHandlingOptions(foptions);
				    
				    
				    t.Start();	
				    
				    //selection.SetElementIds(survivor);
				    doc.Delete(casualty);
				    t.Commit();
				    	    
				    if (t.GetStatus() == TransactionStatus.Committed)
				    {	                	
						message += String.Format("{0}{1}Overall {2} Elements Deleted.", Environment.NewLine, Environment.NewLine, casualty.Count);
				    }
				    else
				    {
				    	message = "Transaction failed";
				    }
				}	       
			}	    		
			
			TaskDialog.Show("IdenticalInstances", message);
		}
		/// <summary>
		/// Attempts to remove the line component of this Warning.
		/// Will procede if no warning is raised during the execution
		/// </summary>
		public void OverlappingWallLine()
		{			
			Document doc = this.ActiveUIDocument.Document;
			
			List<ElementId> survivor = new List<ElementId>();
			List<ElementId> casualty = new List<ElementId>();
			
		    string filename = "";		    
		    
			using(var ofd = new OpenFileDialog())
			{
			    DialogResult result = ofd.ShowDialog();
			
			    if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(ofd.FileName))
			    {
			        filename = ofd.FileName;
			    }
			}
			
		    string message = "";
		    
		    IList<ElementId> ids = new List<ElementId>();		    
		    
	        int success = 0;
	        int failure = 0;
	        
		    bool next = false;
		                
		    using (StreamReader sr = new StreamReader(filename))
		    {
		        string line = "";
		        
		        while ((line = sr.ReadLine()) != null)
		        {
		        	if (next)
		        	{
		        		next = false;
		        				        		
		        		string[] row = System.Text.RegularExpressions.Regex.Split(line,"id ");
						
			            string id1 = row[1].Split(' ')[0];	
			            string id2 = row[2].Split(' ')[0];	
			            			            
			            
			            Wall wall = (doc.GetElement(new ElementId(int.Parse(id1)))) as Wall;
			            ModelLine mline = (doc.GetElement(new ElementId(int.Parse(id2)))) as ModelLine;
			                                            
                        if(wall == null || mline == null)
                        {
                        	continue;
                        }			                                                                        
			                                                                        
			            survivor.Add(new ElementId(int.Parse(id1)));
			            casualty.Add(new ElementId(int.Parse(id2)));
			            			            			                                                                  
			            message += String.Format("Element Id {0}{1}",id1, Environment.NewLine);
		        	}
		        	else
		        	{
		        		var firstIndex = line.IndexOf("A wall and a room separation line overlap.");
		        	
			        	if (firstIndex != -1)
		        	    {	    	
							next = true;		        	    	
		        	    }	 
		        	}		        	      		            
		        }
		        if (casualty.Count == 0)
		        {
		        	message = "No such warning.";
		        	TaskDialog.Show("OverlappingWallLine", message);
		        	
		        	return;
		        }
		        
		        using (Transaction t = new Transaction(doc))
	            {		     
		        	
		        	FailureHandlingOptions options = t.GetFailureHandlingOptions();
	                FailureHandler failureHandler = new FailureHandler();
	                failureHandler = new FailureHandler();
	                options.SetFailuresPreprocessor(failureHandler);
	                options.SetClearAfterRollback(true);
	                t.SetFailureHandlingOptions(options);	
	                
		        	foreach(ElementId id in casualty)
		        	{		
		                t.Start("IdenticalInstances_RemoveElements");	
		                doc.Delete(casualty);
		                t.Commit();	
		                /*
		                if (failureHandler.ErrorMessage != "")
		                {	                	
			        		message += String.Format("Error: {0} with severity {1}. Item {2} not processed. {3}",
		                	                        failureHandler.ErrorMessage, failureHandler.ErrorSeverity, 
		                	                        id.ToString(), Environment.NewLine);
		                }	       
						*/ 	
		        	}
	            }		        
			}	    		
		    message += String.Format("{0}Overall {1} Elements Deleted and {2} Failed.", Environment.NewLine, success.ToString(), failure.ToString());
			TaskDialog.Show("OverlappingWallLine", message);
		}

		public class FailureHandler : IFailuresPreprocessor
		{
			//public string ErrorMessage { set; get; }
			//public string ErrorSeverity { set; get; }
			
			public FailureHandler()
			{
				//ErrorMessage = "";
				//ErrorSeverity = "";
			}
			
			public FailureProcessingResult PreprocessFailures(
				FailuresAccessor failuresAccessor)
			{
				IList<FailureMessageAccessor> failureMessages = failuresAccessor.GetFailureMessages();
				
				if (failureMessages.Count == 0)
				{	
					return FailureProcessingResult.Continue;
				}				
				else
				{	
					foreach(FailureMessageAccessor fma in failureMessages)
					{
						FailureSeverity fsav = fma.GetSeverity();
						
						if(fsav == FailureSeverity.Warning)
						{
							failuresAccessor.DeleteWarning(fma);
						}
						else
						{
							failuresAccessor.ResolveFailure(fma);
							return FailureProcessingResult.ProceedWithCommit;
						}
					}
					return FailureProcessingResult.Continue;
				}
				/*
				foreach( FailureMessageAccessor failureMessageAccessor in failureMessages)
				{
					FailureDefinitionId id = failureMessageAccessor.GetFailureDefinitionId();
					
					try{
						ErrorMessage = failureMessageAccessor.GetDescriptionText();
					}
					catch
					{
						ErrorMessage = "Unknown Error";
					}
					
					try{
						FailureSeverity failureSeverity = failureMessageAccessor.GetSeverity();
						
						ErrorSeverity = failureSeverity.ToString();
						
						if (failureSeverity == FailureSeverity.Warning || failureSeverity == FailureSeverity.Error)
						{
							return FailureProcessingResult.ProceedWithRollBack;
						}
					}
					catch
					{						
					}
					*/
			}
		}
		public void AlignViews()
		{
			Document doc = this.ActiveUIDocument.Document;
			Selection sel = this.ActiveUIDocument.Selection;
			
			Viewport vp = doc.GetElement(sel.PickObject(ObjectType.Element, "Pick View to Align To")) as Viewport;
			
			if(vp == null) 
			{
				TaskDialog.Show("Error", "That's not a View");
				return;
			}
						
			XYZ loc = vp.GetBoxCenter();
			
			List<ViewSheet> viewSheets = new FilteredElementCollector(doc)
				.OfClass(typeof(ViewSheet))
				.WhereElementIsNotElementType()
				.Cast<ViewSheet>()
				.ToList();
			
			List<Viewport> viewPorts = new List<Viewport>();
			
			foreach(ViewSheet sheet in viewSheets)
			{
				List<Viewport> ports = sheet.GetAllViewports()
					.Select<ElementId, Viewport>(
						id => doc.GetElement(id) as Viewport)
					.ToList<Viewport>();
				
				viewPorts.AddRange(ports);
			}
			
			viewPorts = viewPorts.Where(x => x.LookupParameter("View Name").AsString().Contains("GA")).ToList();
			
			using(Transaction t = new Transaction(doc, "Align Views"))
			{
				t.Start();
				foreach(Viewport v in viewPorts)
				{
					XYZ delta = loc - v.GetBoxCenter();
					ElementTransformUtils.MoveElement(doc, v.Id, delta);
				}
				t.Commit();
			}
		}
		public void FamilyTypeCreate()
		{
			Document doc = this.ActiveUIDocument.Document;
			if (!doc.IsFamilyDocument) 
			{
				TaskDialog.Show("Error", "Only execute in family document");
				return;
			}
			
			FamilyManager famanager = doc.FamilyManager;
			FamilyType famtype = famanager.CurrentType;
			if (famtype == null)
			{
				using(Transaction t = new Transaction(doc, "Create Family Type"))
				{
					t.Start();
					famanager.NewType("Default");
					t.Commit();
				}
				
				using(Transaction t = new Transaction(doc, "Delete Family Type"))
				{
					t.Start();
					famanager.DeleteCurrentType();
					t.Commit();
				}
				
				return;					
			}
			else
			{
				TaskDialog.Show("Error", famtype.Name + ":" + famtype.ToString());
			}
		}
		public void FamilyCount()
		{
			Document doc = this.ActiveUIDocument.Document;
			
			FilteredElementCollector col = new FilteredElementCollector(doc);
			
			List<Element> elements = col
				.OfClass(typeof(Family))
				.ToList();
			
			col = new FilteredElementCollector(doc);
			
			List<FamilyInstance> instances = col
				.OfClass(typeof(FamilyInstance))
				.WhereElementIsNotElementType()
				.Cast<FamilyInstance>()
				.ToList();
						
			DataSet dataSet = new DataSet(String.Format("{0} family usage.", doc.Title));
			DataTable table = new DataTable(String.Format("Family usage {0}", DateTime.Today));
			table.Columns.Add("Family Name");
			table.Columns.Add("File Size");
			table.Columns.Add("Number of times used");
			table.Columns.Add("Element Id");
			
			
			string s = "";
			
			s += String.Format("Total number of families: {0}{1}{2}", elements.Count, Environment.NewLine, Environment.NewLine);
			
			long totalSize = 0;
			
			foreach(var family in elements)
			{	
				string count = instances.Where(x => x.Symbol.FamilyName.Equals(family.Name)).ToList().Count.ToString();	
				long size = familySize(family as Family);		
				totalSize += size;
				
				//string count = "";
				s += String.Format("{0} : {1} - used {2} times {3}", family.Name, convertSize(size), count, Environment.NewLine);
				table.Rows.Add(family.Name, convertSize(size), count, family.Id);
			}
			s += String.Format("{0}{0}Total size used {1}",Environment.NewLine, convertSize(totalSize));
			TaskDialog.Show("Number of Families", s);
			
			var builder = new StringBuilder();
			
			foreach(DataRow row in table.Rows)
			{
				builder.AppendLine(String.Join("\t",row.ItemArray));				
			}
			var file = new FileStream("C:/Users/dene/Documents/Working/macros/family size/report.txt",FileMode.Create);
			var writer = new StreamWriter(file);
			writer.Write(builder.ToString());
			writer.Flush();
			writer.Close();
			return;
		}
		private string convertSize(long length)
		{
			if(length == -1) return "Family is not editable";
			string[] sizes = { "B", "KB", "MB", "GB" };
			int order = 0;
			while (length >= 1024 && ++order < sizes.Length) {
			    length = length/1024;
			}
			string result = String.Format("{0:0.##} {1}", length, sizes[order]);
			return result;				
		}
		private long familySize(Family family)
		{
			try
			{				
				Document doc = this.ActiveUIDocument.Document.EditFamily(family);
				string str = Path.Combine(Path.GetTempPath(), doc.Title);
				SaveAsOptions saveAsOptions = new SaveAsOptions();
				saveAsOptions.OverwriteExistingFile = true;
				doc.SaveAs(str, saveAsOptions);
				long length = new FileInfo(str).Length;
				
				File.Delete(str);
				return length;
			}
			catch (Exception ex)
			{
				return -1;
			}
		}
		public void DeleteAreaLines()
		{			
			Document doc = this.ActiveUIDocument.Document;
			
			List<ElementId> survivor = new List<ElementId>();
			List<ElementId> casualty = new List<ElementId>();
			
		    string filename = "";
		    		    
			using(var ofd = new OpenFileDialog())
			{
			    DialogResult result = ofd.ShowDialog();
			
			    if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(ofd.FileName))
			    {
			        filename = ofd.FileName;
			    }
			}
			
		    string message = "";
		    
		    IList<ElementId> ids = new List<ElementId>();		    
		    
		    bool next = false;
		                
		    using (StreamReader sr = new StreamReader(filename))
		    {
		        string line = "";
		        
		        while ((line = sr.ReadLine()) != null)
		        {
		        	if (next)
		        	{
		        		next = false;
		        				        		
		        		string[] row = System.Text.RegularExpressions.Regex.Split(line,"id ");
						
			            string id1 = row[1].Split(' ')[0];		
			            
			            if(id1.ToString().Equals("3273058")) continue;
			            
			            casualty.Add(new ElementId(int.Parse(id1)));
			            			            			                                                                  
			            message += String.Format("Element Id {0}{1}",id1, Environment.NewLine);
		        	}
		        	else
		        	{
		        		var firstIndex = line.IndexOf("Area separation line is slightly off axis and may cause inaccuracies");
		        	
			        	if (firstIndex != -1)
		        	    {	    	
							next = true;		        	    	
		        	    }	 
		        	}		        	      		            
		        }
		        if (casualty.Count == 0)
		        {
		        	message = "No such warning.";
		        	TaskDialog.Show("AreaLines", message);
		        	
		        	return;
		        }
		        
		        using (Transaction t = new Transaction(doc))
	            {		                
	                t.Start("DeleteAreaLines");	
	                doc.Delete(casualty);
	                t.Commit();	
	            }		        
			}	    		
		    message += String.Format("{0}Overall {1} Elements Deleted.", Environment.NewLine, casualty.Count.ToString());
			TaskDialog.Show("DeleteAreaLines", message);
		}
		
		/// <summary>
		/// Arranges (vertically) a number of selected viewports on a sheet
		/// </summary>
		public void ArrangeViewports()
		{						
			UIDocument uidoc = this.ActiveUIDocument;
			Document doc = this.ActiveUIDocument.Document;
			
			ViewPortSelectionFilter filter = new ThisApplication.ViewPortSelectionFilter(doc);
			IList<Reference> viewports = uidoc.Selection.PickObjects(ObjectType.Element, filter, "Select multiple viewports");
			
			if (viewports.Count > 1)
			{
				//Align(doc, viewports, "Vertical");
				Distribute(doc, viewports, "Horizontal");
			}		
		}
		public void ImportedDWG()
        {
            Document doc = ActiveUIDocument.Document;
            
            FilteredElementCollector col = new FilteredElementCollector(doc)
                .OfClass(typeof(ImportInstance));
            
            IList<ImportInstance> elements = col
                .Cast<ImportInstance>()
                .Where(x => !x.IsLinked)
                .ToList();
            
            TaskDialog.Show("NumberImports", String.Format("There are {0} number of imported files", 
                                                           elements.Count().ToString()));
            string s = "";
            
            foreach(ImportInstance instance in elements)
            {
                s += instance.LookupParameter("Name").AsString() + " : " + doc.GetElement(instance.OwnerViewId).Name.ToString() + Environment.NewLine;                
            }
            TaskDialog.Show("NumberImports", String.Format("There are {0}{1}", Environment.NewLine,
                                                           s));
            
            return;
        }        
		public void OffAxis()
		{
            Document doc = this.ActiveUIDocument.Document;			
			Selection selection = this.ActiveUIDocument.Selection;
			List<ElementId> ids = selection.PickObjects(ObjectType.Element, "Pick Axis to fix").ToList().Select(x => x.ElementId).ToList();
			
			foreach(ElementId id in ids)
			{
				Fix(id);
			}
		}
		
	    private void Fix(ElementId id)
        {            
            Document doc = this.ActiveUIDocument.Document;
			Curve locationLine = (doc.GetElement(id) as Grid).Curve;
            
			/*
            // mini procedure to make sure rotation a line doesn't drag other lines with it
            if(doc.GetElement(id).Name.Equals("Model Lines"))
            {
                ElementArray empty = locationLine.get_ElementsAtJoin(0);
                using (Transaction temp = new Transaction(doc, "Temporary cut joined element"))
                {
                    temp.Start();
                    temp.Commit();
                }
            }
			*/
            double rotation = getRotation(locationLine);

            Line axis = getAxis(locationLine);

            using (Transaction t = new Transaction(doc, "Rotate off axis element."))
            {
                t.Start();
                ElementTransformUtils.RotateElement(doc, id, axis, (rotation));
                t.Commit();
            }
        }
        /// <summary>
        ///  Get Axis of rotation (Z) of line
        /// </summary>
        /// <param name="locationLine"></param>
        /// <returns></returns>
        private Line getAxis(Curve locationLine)
        {
            XYZ basePoint = locationLine.GetEndPoint(0);
            XYZ direction = (locationLine as Line).Direction;
            XYZ cross = XYZ.BasisY.CrossProduct(direction).Normalize();
            return cross.IsZeroLength() ? null : Line.CreateBound(basePoint, cross+basePoint);
        }
        /// <summary>
        ///  Get Axis of rotation (Z) of plane
        /// </summary>
        /// <param name="locationLine"></param>
        /// <returns></returns>
        private Line getAxis(Plane plane)
        {
            XYZ basePoint = plane.Origin;
            XYZ direction = plane.XVec;
            XYZ cross = XYZ.BasisY.CrossProduct(direction).Normalize();
            return cross.IsZeroLength() ? null : Line.CreateBound(basePoint, cross + basePoint);
        }
        /// <summary>
        /// Get rotation angle of line
        /// </summary>
        /// <param name="locationLine"></param>
        /// <returns></returns>
        private double getRotation(Curve locationLine)
        {
            Line line = locationLine as Line;
            double angle = XYZ.BasisY.AngleTo(line.Direction);
            if (angle > 0)
            {
                while (angle > 0.01)
                {
                    angle -= Math.PI / 4;
                }
            }
            else
            {
                while (angle < -0.01)
                {
                    angle += Math.PI / 4;
                }
            }
            return -angle;
        }
	public void SelectRoomBoundingElements()
	{
		UIDocument uidocument = this.ActiveUIDocument;
		Document doc = this.ActiveUIDocument.Document;
		Selection sel = this.ActiveUIDocument.Selection;

		ICollection<ElementId> joinedElements = new Collection<ElementId>(); // collection to store the walls joined to the selected wall

		Room room = doc.GetElement(sel.PickObject(ObjectType.Element, "Pick the Room").ElementId) as Room;

		SpatialElementBoundaryOptions sebOptions = new SpatialElementBoundaryOptions();			

		var boundaries = room.GetBoundarySegments(sebOptions);

		foreach(var boundary in boundaries)
		{
			foreach(var segment in boundary)
			{
				joinedElements.Add(segment.ElementId);
			}					
		}

		uidocument.Selection.SetElementIds(joinedElements); // select all of the joined elements
	}
	public void RenameGrids()
	{
		Document doc = this.ActiveUIDocument.Document;
		Selection sel = this.ActiveUIDocument.Selection;

		List<Grid> grids = sel.PickObjects(ObjectType.Element, "Pick Grids").Select<Reference, Grid>(
			x => doc.GetElement(x.ElementId) as Grid)
			.ToList<Grid>();

		if(grids.Count == 0) TaskDialog.Show("Error", "No grids selected");

		using(Transaction t = new Transaction(doc, "Grids Rename"))
		{
			int counter = 1;
			t.Start();
			foreach(Grid g in grids)
			{
				g.Name = String.Format("F{0}", counter.ToString());
				counter ++;
			}				
			t.Commit();
		}
	}	
        public void PurgeImportedLines()
        {
            Document doc = ActiveUIDocument.Document;
            
            string m = "";
            
            FilteredElementCollector col = new FilteredElementCollector(doc)
                .OfClass(typeof(LinePatternElement));
            
            List<ElementId> lpeIds = col.ToElementIds().Where(x => doc.GetElement(x).Name.Contains("IMPORT")).ToList();
            List<LinePatternElement> linePatterns = col.Cast<LinePatternElement>().Where(x => x.Name.Contains("IMPORT")).ToList();
            
            using (Transaction t = new Transaction(doc, "Purge imported line patterns"))
            {
                t.Start();
                foreach(LinePatternElement lpe in linePatterns)
                {
                    m += lpe.Name + Environment.NewLine;
                }
                doc.Delete(lpeIds);
                t.Commit();
            }
            
            m += Environment.NewLine + String.Format(("A total of {0} imported line patterns have been removed from this project"), linePatterns.Count.ToString());
            TaskDialog.Show("PurgeImportedLines", m);
            return;            
        }
		private void Distribute(Document doc, IList<Reference> viewports, string direction)
		{		
			Viewport [] viewportArray = null;			
			double [] centers = null;			
			
			SetValues(doc, viewports, direction, ref viewportArray, ref centers);		
			
			double distance = (centers[centers.Length - 1] - centers[0])/(centers.Length - 1);
			double delta = 0;
			XYZ translation = null;
			
			using(Transaction t = new Transaction(doc, "Arrange Viewports"))
			{				
				t.Start();
				for (int i = 0; i < centers.Length - 1; i++)
				{
					delta = (i*distance - (centers[i]-centers[0]));
					switch(direction)
					{
						case "Vertical": case "Top": case "Bottom":							
							translation = new XYZ(0, delta, 0);
							break;
						case "Horizontal": case "Left": case "Right":						
							translation = new XYZ(delta, 0, 0);
							break;							
					}
					ElementTransformUtils.MoveElement(doc, viewportArray[i].Id, translation);
				}
				t.Commit();
			}
			return;
		}
		
		private void Align(Document doc, IList<Reference> viewports, string direction)
		{		
			Viewport [] viewportArray = null;			
			double [] centers = null;
			
			SetValues(doc, viewports, direction, ref viewportArray, ref centers);
			
			double distance = (centers[centers.Length - 1] + centers[0])/2;
			double delta = 0;
			
			XYZ translation = null;
			
			using(Transaction t = new Transaction(doc, "Arrange Viewports"))
			{				
				t.Start();
				for (int i = 0; i < centers.Length; i++)
				{
					delta = distance-centers[i];
					if(direction == "Vertical" || direction == "Top" || direction == "Bottom")
					{
						translation = new XYZ(0, delta, 0);						
					}
					else 
					{						
						translation = new XYZ(delta, 0, 0);	
					}
					ElementTransformUtils.MoveElement(doc, viewportArray[i].Id, translation);
				}
				t.Commit();
			}
			return;
		}
		
		private void SetValues(Document doc, IList<Reference> viewports, string direction, ref Viewport [] viewportArray, ref double [] centers)
		{			
			switch(direction)
			{
				case "Vertical":					
					viewportArray = viewports.Select(i => doc.GetElement(i) as Viewport).OrderBy(o => o.GetBoxCenter().Y).ToArray();
					centers = viewportArray.Select(x => x.GetBoxCenter().Y).ToArray();
					break;
				case "Horizontal":					
					viewportArray = viewports.Select(i => doc.GetElement(i) as Viewport).OrderBy(o => o.GetBoxCenter().X).ToArray();
					centers = viewportArray.Select(x => x.GetBoxCenter().X).ToArray();
					break;
				case "Top":
					viewportArray = viewports.Select(i => doc.GetElement(i) as Viewport).OrderBy(o => o.GetBoxOutline().MinimumPoint.Y).ToArray();
					centers = viewportArray.Select(x => x.GetBoxOutline().MinimumPoint.Y).ToArray();					
					break;
				case "Bottom":
					viewportArray = viewports.Select(i => doc.GetElement(i) as Viewport).OrderBy(o => o.GetBoxOutline().MaximumPoint.Y).ToArray();
					centers = viewportArray.Select(x => x.GetBoxOutline().MaximumPoint.Y).ToArray();					
					break;
				case "Left":
					viewportArray = viewports.Select(i => doc.GetElement(i) as Viewport).OrderBy(o => o.GetBoxOutline().MinimumPoint.X).ToArray();
					centers = viewportArray.Select(x => x.GetBoxOutline().MinimumPoint.X).ToArray();					
					break;
				case "Right":
					viewportArray = viewports.Select(i => doc.GetElement(i) as Viewport).OrderBy(o => o.GetBoxOutline().MaximumPoint.X).ToArray();
					centers = viewportArray.Select(x => x.GetBoxOutline().MaximumPoint.X).ToArray();					
					break;
			}		
		}
//		public void AttachedMiss()
//		{
//			Document doc = this.ActiveUIDocument.Document;
//		    string filename = Path.Combine("S:/15504_KXG/00_BR-PEERSYNC/02_BIM/01-WIP/01.07-Temp/01.07.04-Warnings/KXC-A-001-A-BH-M3-BaseBuilding@big.html");
//		    string message = "";
//		    
//		    //IList<ElementId> ids = new List<ElementId>();		    
//		    Dictionary<Wall, Floor> elements = new Dictionary<Wall, Floor>();
//		    
//		    bool next = false;
//		                
//		    using (StreamReader sr = new StreamReader(filename))
//		    {
//		        string line = "";
//		        
//		        while ((line = sr.ReadLine()) != null)
//		        {
//		        	if (next)
//		        	{
//		        		next = false;
//		        				        		
//		        		string[] row = System.Text.RegularExpressions.Regex.Split(line,"id ");
//						
//			            string id1 = row[1].Split(' ')[0];	
//			            string id2 = row[2].Split(' ')[0];	
//			            			            			            
//			            Floor floor = (doc.GetElement(new ElementId(int.Parse(id1)))) as Floor;
//			            Wall wall = (doc.GetElement(new ElementId(int.Parse(id2)))) as Wall;
//			                                            
//                        if(wall == null || floor == null)
//                        {
//                        	continue;
//		        			TaskDialog.Show("AttachedButMiss", "here");
//                        }		
//
//                        elements.Add(wall, floor);
//			            			            			                                                                  
//			            message += String.Format("Element Id {0}{1}",id1, Environment.NewLine);
//		        	}
//		        	else
//		        	{
//		        		var firstIndex = line.IndexOf("Highlighted walls are attached to, but miss, the highlighted targets.");
//		        	
//			        	if (firstIndex != -1)
//		        	    {	    	
//							next = true;			        	    	
//		        	    }	 
//		        	}		        	      		            
//		        }
//		        TaskDialog.Show("AttachedButMiss", message);	
//		        return;
//		        
//		        if (elements.Count == 0)
//		        {
//		        	message = "No such warning.";
//		        	TaskDialog.Show("OverlappingWallLine", message);
//		        	
//		        	return;
//		        }
//		        
//		        using (Transaction t = new Transaction(doc))
//	            {		     
//		        	
//		        	FailureHandlingOptions options = t.GetFailureHandlingOptions();
//	                FailureHandler failureHandler = new FailureHandler();
//	                failureHandler = new FailureHandler();
//	                options.SetFailuresPreprocessor(failureHandler);
//	                options.SetClearAfterRollback(true);
//	                t.SetFailureHandlingOptions(options);	
//	                
//		        	foreach(KeyValuePair<Wall, Floor> pair in elements)
//		        	{	                
//		                t.Start("IdenticalInstances_RemoveElements");	
//		                
//		                t.Commit();	
//		                /*
//		                if (failureHandler.ErrorMessage != "")
//		                {	                	
//			        		message += String.Format("Error: {0} with severity {1}. Item {2} not processed. {3}",
//		                	                        failureHandler.ErrorMessage, failureHandler.ErrorSeverity, 
//		                	                        id.ToString(), Environment.NewLine);
//		                }	       
//						*/ 	
//		        	}
//	            }		        
//			}	    		
//		    message += String.Format("{0}Overall {1} Walls were detached.", Environment.NewLine, elements.Count.ToString());
//			TaskDialog.Show("OverlappingWallLine", message);			
//		}
		public void DuplicateMarkValues()
		{
			Document doc = this.ActiveUIDocument.Document;
			
			List<Element> elements = new List<Element>();
			
		    string filename = "";		
			
			using(var ofd = new OpenFileDialog())
			{
			    DialogResult result = ofd.ShowDialog();
			
			    if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(ofd.FileName))
			    {
			        filename = ofd.FileName;
			    }
			}
			
		    string message = "";
		    
		    IList<ElementId> ids = new List<ElementId>();		    
		    
		    bool next = false;
		                
		    using (StreamReader sr = new StreamReader(filename))
		    {
		        string line = "";
		        
		        while ((line = sr.ReadLine()) != null)
		        {
		        	if (next)
		        	{
		        		next = false;
		        				        		
		        		string[] row = System.Text.RegularExpressions.Regex.Split(line,"id ");
						
			            string id = "";		
			            		
				        for( int i = 1; i < row.Length; i++)
				        {
				        	id = row[i].Split(' ')[0];
				        	ids.Add((new ElementId(int.Parse(id))));                                                     
			            	message += String.Format("Element Id {0}{1}",id, Environment.NewLine);
				        }			            			            			            			             
		        	}
		        	else
		        	{
		        		var firstIndex = line.IndexOf("Elements have duplicate 'Mark' values.");
		        	
			        	if (firstIndex != -1)
		        	    {	    	
							next = true;		        	    	
		        	    }	 
		        	}		        	      		            
		        }
		        if (ids.Count == 0)
		        {
		        	message = "No such warning.";
		        	TaskDialog.Show("AreaLines", message);
		        	
		        	return;
		        }
		        
		        message = "";
		        using (Transaction t = new Transaction(doc))
	            {		                
	                t.Start("DuplicateMarkValues");	
	                foreach(ElementId id in ids)
	                {
	                	doc.GetElement(id).LookupParameter("Comments").Set(doc.GetElement(id).LookupParameter("Mark").AsString());
	                	doc.GetElement(id).LookupParameter("Mark").Set("");
	                }
	                t.Commit();	
	            }		        
			}	    		
		    message += String.Format("{0}The mark values of overall {1} Elements changed.", Environment.NewLine, ids.Count.ToString());
			TaskDialog.Show("DeleteAreaLines", message);			
		}
		public void SelectRoomBoundingElements()
		{
			UIDocument uidocument = this.ActiveUIDocument;
			Document doc = this.ActiveUIDocument.Document;
			Selection sel = this.ActiveUIDocument.Selection;
			
			ICollection<ElementId> joinedElements = new Collection<ElementId>(); // collection to store the walls joined to the selected wall
			
			Room room = doc.GetElement(sel.PickObject(ObjectType.Element, "Pick the Room").ElementId) as Room;
			
			SpatialElementBoundaryOptions sebOptions = new SpatialElementBoundaryOptions();			
			
			var boundaries = room.GetBoundarySegments(sebOptions);
			
			foreach(var boundary in boundaries)
			{
				foreach(var segment in boundary)
				{
					joinedElements.Add(segment.ElementId);
				}					
			}
			
			uidocument.Selection.SetElementIds(joinedElements); // select all of the joined elements
		}
		public void FindLegendsOnSheets()
		{
			UIDocument uidocument = this.ActiveUIDocument;
			Document doc = this.ActiveUIDocument.Document;
			
			string legendName = Prompt.ShowDialog("Which Legend?", "o.O");
			
			List<ViewSheet> sheetList = new FilteredElementCollector(doc)
				.WhereElementIsNotElementType()
				.OfClass(typeof(ViewSheet))
				.Cast<ViewSheet>()
				.ToList();
			
			string s = "";
			
			foreach(var sheet in sheetList)
			{
				var placedViews = sheet.GetAllPlacedViews();
				
				if(placedViews.Count == 0) continue;
				
				foreach(var view in placedViews)
				{
					Autodesk.Revit.DB.View v = doc.GetElement(view) as Autodesk.Revit.DB.View;
					if(v.ViewType != ViewType.Legend) continue;
					if(v.Name.Equals(legendName)) s += sheet.SheetNumber + " - " + sheet.Name + Environment.NewLine;
				}
			}
			
			TaskDialog.Show("Results", s);			
		}		
		internal static class Prompt
		{
		    public static string ShowDialog(string text, string caption)
		    {
		        System.Windows.Forms.Form prompt = new System.Windows.Forms.Form()
		        {
		            Width = 500,
		            Height = 150,
		            FormBorderStyle = FormBorderStyle.FixedDialog,
		            Text = caption,
		            StartPosition = FormStartPosition.CenterScreen
		        };
		        Label textLabel = new Label() { Left = 50, Top=20, Text=text };
		        System.Windows.Forms.TextBox textBox = new System.Windows.Forms.TextBox() { Left = 50, Top=40, Width=400 };
		        Button confirmation = new Button() { Text = "Ok", Left=350, Width=100, Top=70, DialogResult = DialogResult.OK };
		        confirmation.Click += (sender, e) => { prompt.Close(); };
		        prompt.Controls.Add(textBox);
		        prompt.Controls.Add(confirmation);
		        prompt.Controls.Add(textLabel);
		        prompt.AcceptButton = confirmation;
		
		        return prompt.ShowDialog() == DialogResult.OK ? textBox.Text : "";
		    }
		}
		public void RenameGrids()
		{
			Document doc = this.ActiveUIDocument.Document;
			Selection sel = this.ActiveUIDocument.Selection;
			
			List<Grid> grids = sel.PickObjects(ObjectType.Element, "Pick Grids").Select<Reference, Grid>(
				x => doc.GetElement(x.ElementId) as Grid)
				.ToList<Grid>();
			
			if(grids.Count == 0) TaskDialog.Show("Error", "No grids selected");
			
			using(Transaction t = new Transaction(doc, "Grids Rename"))
			{
				int counter = 1;
				t.Start();
				foreach(Grid g in grids)
				{
					g.Name = String.Format("F{0}", counter.ToString());
					counter ++;
				}				
				t.Commit();
			}
		}
		public void PDFMultipleDocuments()
		{			
			OpenFileDialog theDialogRevit = new OpenFileDialog();
			theDialogRevit.Title = "Select Revit Project Files";
			theDialogRevit.Filter = "RVT files|*.rvt";
			theDialogRevit.FilterIndex = 1;
			theDialogRevit.Multiselect = true;
			
			OpenOptions opt = new OpenOptions();
            opt.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets;
            opt.Audit = false;
            
			if (theDialogRevit.ShowDialog() == DialogResult.OK)
			{
                foreach (String projectPath in theDialogRevit.FileNames)
                {
                 	FileInfo filePath = new FileInfo(projectPath);
                 	string filename = filePath.Name;
                    ModelPath mp = ModelPathUtils.ConvertUserVisiblePathToModelPath(filePath.FullName);					
		            WorksetConfiguration wc = new WorksetConfiguration(WorksetConfigurationOption.OpenAllWorksets);
		            
		            try{
			            Document doc = Application.OpenDocumentFile(mp,opt);
			            PrintViewSets(doc);
			            doc.Close(false);		            	
		            }
		            catch(Exception)
		            {
		            	
		            }
                }
			}
		}
		private void PrintViewSets(Document doc)
		{
			List<ViewSheetSet> viewSets = new FilteredElementCollector(doc).OfClass(typeof(ViewSheetSet)).Cast<ViewSheetSet>().ToList();			

			// No ViewSets No Game			
			if(viewSets.Count == 0) 
			{
				TaskDialog.Show("Error", String.Format("No ViewSets in {0} found.", doc.Title));
				return;
			}
			
			foreach(ViewSheetSet vset in viewSets)
			{
				using (Transaction t = new Transaction(doc, "Print Test"))
				{
					t.Start();
					doc.Print(vset.Views);				
					t.Commit();
				}				
			}			
		}
		public void ElementToWorkset()
		{			
			UIDocument uidoc = this.ActiveUIDocument;
			Document doc = ActiveUIDocument.Document;
			string workName = "X-Admin";
			
			LineSelectionFilter filter = new ThisApplication.LineSelectionFilter(doc);
			IList<Reference> references = uidoc.Selection.PickObjects(ObjectType.Element, filter, "Select multiple lines");
			
			FilteredElementCollector collector = new FilteredElementCollector(doc);
			FilteredWorksetCollector workCollector = new FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset);
			
			Workset workset = workCollector.Single(x => x.Name.Equals(workName));
			//IList<Element> elements = collector.OfCategory(BuiltInCategory.OST_AreaSchemeLines).ToElements();
			//IList<Element> elements = collector.OfCategory(BuiltInCategory.OST_Areas).ToElements();
			IList<Element> elements = references.Select(x => doc.GetElement(x)).ToList();
			int converted = 0;
			using(Transaction t = new Transaction(doc,"ElementToWorkset"))
			{
				t.Start();
				foreach(Element el in elements)
				{
					if(el.WorksetId.IntegerValue != workset.Id.IntegerValue) //only convert if they are not in that workset
					{
						el.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM).Set(workset.Id.IntegerValue);	
						converted++;						
					}
				}
				t.Commit();
			}
			
			TaskDialog.Show("AreaBoundaries", String.Format("{0} number of Area Boundary Lines were assigned to {1} workset", converted.ToString(), workName));
		}
		public void Assign()
		{
			UIDocument uidoc = this.ActiveUIDocument;
			Document doc = ActiveUIDocument.Document;
			
			List<Element> elements = new List<Element>();
			ICollection<ElementId> selection = uidoc.Selection.GetElementIds();
			
			foreach(ElementId id in selection)
			{
				elements.Add(doc.GetElement(id));
			}
			
			FilteredWorksetCollector worksetCollector = new FilteredWorksetCollector(doc);
            		List<Workset> worksets = worksetCollector.OfKind(WorksetKind.UserWorkset).ToList();
            
			WorksetId worksetId = null;
			
            foreach (Workset workset in worksets)
            {
            	if(workset.Name.Equals("X-Admin"))
            	{
            		worksetId = workset.Id;
            		break;
            	}
            }
            TaskDialog.Show("Bzzt", String.Format("Element count: {0} {1}Workset: {2}", elements.Count.ToString(), Environment.NewLine, worksetId.IntegerValue.ToString()));
            
			
			using (Transaction t = new Transaction(doc, "Assign"))
            {
                FailureHandlingOptions foptions = t.GetFailureHandlingOptions();
                FailureHandler fhandler = new FailureHandler();
                foptions.SetFailuresPreprocessor(fhandler);
                foptions.SetClearAfterRollback(true);
                t.SetFailureHandlingOptions(foptions);

                t.Start();

                foreach (Element el in elements)
                {
                    if (el.WorksetId.IntegerValue != worksetId.IntegerValue) //only convert if they are not in that workset
                    {
                        if(el.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM).IsReadOnly)
                        {
                            //failed++;
                            continue;
                        }
                        try
                        {
                            el.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM).Set(worksetId.IntegerValue);
                            //converted++;
                        }
                        catch
                        {

                        }
                    }
                }
                t.Commit();
            }
		}
		public void Select()
		{
			UIDocument uidoc = this.ActiveUIDocument;
			Document doc = ActiveUIDocument.Document;
			
			Categories categories = doc.Settings.Categories;
			Category cat = null;
			
			foreach (Category c in categories)
			{
				if(c.Name.Equals("Lines"))
				{
					cat = c;
					break;
				}				   
			}
			
			List<ElementId> elements = new List<ElementId>();
            List<Category> subCat = getSubCategories(doc, cat);

            if (subCat != null)
            {
                foreach (Category subCategory in subCat)
                {
                    //The recursion
                    var subElem = recursiveElements(doc, subCategory);
                    elements.AddRange(subElem);
                }
            }

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            List<ElementId> thisElements = collector.OfCategoryId(cat.Id)
                .WhereElementIsNotElementType()
                .ToElementIds()
                .ToList();
            elements.AddRange(thisElements);
			
            uidoc.Selection.SetElementIds(elements);   
		}
		internal List<Category> getSubCategories(Document doc, Category cat)
        {
            List<Category> listCat = null;
            var categories = cat.SubCategories;
            if(!categories.IsEmpty)
            {
                listCat = new List<Category>();
                foreach (Category subCat in categories)
                {
                    listCat.Add(subCat);
                }
            }
            return listCat;
        }
		
		internal List<ElementId> recursiveElements(Document doc, Category cat)
        {
            List<ElementId> elements = new List<ElementId>();
            List<Category> subCat = getSubCategories(doc, cat);

            if (subCat != null)
            {
                foreach (Category subCategory in subCat)
                {
                    //The recursion
                    var subElem = recursiveElements(doc, subCategory);
                    elements.AddRange(subElem);
                }
            }

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            List<ElementId> thisElements = collector.OfCategoryId(cat.Id)
                .WhereElementIsNotElementType()
                .ToElementIds()
                .ToList();
            elements.AddRange(thisElements);

            return elements;
        }
		public void DeleteSheets()
		{
			UIDocument uidoc = this.ActiveUIDocument;
			Document doc = ActiveUIDocument.Document;
			
			FilteredElementCollector collector = new FilteredElementCollector(doc);
			
			List<ElementId> ids = collector.OfClass(typeof(ViewSheet)).ToElementIds().ToList();
			
			using(Transaction t = new Transaction(doc, "Delete All Sheets"))
			{
				t.Start();
				doc.Delete(ids);
				t.Commit();
			}
		}
		public void ChangeTypeParamter()
		{
			UIDocument uidoc = this.ActiveUIDocument;
			Document doc = ActiveUIDocument.Document;
			
			FilteredElementCollector collector = new FilteredElementCollector(doc);
			
			ICollection<Element> panels = collector.OfClass(typeof(FamilyInstance)).OfCategory(BuiltInCategory.OST_GenericModel).ToElements();
			List<Element> sorted = new List<Element>();
			string s = "";
			string n = "";
			string typeName = "";
			string prefix = "Panel Type {0}";
			string message = "";
			
			foreach(Element panel in panels)
			{
				if(null != panel.GetTypeId())
				{
					Element panelType = doc.GetElement(panel.GetTypeId());
					if(panelType.LookupParameter("Type Mark").AsString() != null &&  panelType.LookupParameter("Type Mark").AsString().Contains("GL"))
					{
						if(panel.IsValidObject && panel.LookupParameter("Glass Length") != null && panel.LookupParameter("Glass Heigth") != null)
						{
							sorted.Add(panel);
						}
					}
				}
			}
			
			sorted = sorted.OrderBy(x => x.LookupParameter("Glass Length").AsValueString()).
				ThenBy(x => x.LookupParameter("Glass Heigth").AsValueString()).ToList();
				
			int c = 0;
			
			using (Transaction t = new Transaction(doc, "Rename"))
			{
				t.Start();
				foreach(Element panel in sorted)
				{				
					s = string.Format("{0}-{1}", 
					                   panel.LookupParameter("Glass Length").AsValueString(),
					                   panel.LookupParameter("Glass Heigth").AsValueString() + Environment.NewLine);
					if(s != n)
					{
						n = s;
						c++;
						typeName = string.Format(prefix, c.ToString());
						message += typeName + Environment.NewLine;
					}
					
					SetTypeParameter(doc, panel.Id, "Comments", typeName);
				}
				t.Commit();
			}
			/*
			foreach(Element pane in panes)
			{
				s += pane.LookupParameter("Glass Length").AsDouble().ToString() + Environment.NewLine;
				
			}
			*/
			TaskDialog.Show("ChangeTypeParamter", message);
			
		}
		private void SetTypeParameter(Document doc, ElementId id, string param, string name)
		{
			doc.GetElement(id).LookupParameter(param).Set(name);
		}
		public void OpenDetachSave()
		{
			UIDocument uidoc = this.ActiveUIDocument;
			Document doc = ActiveUIDocument.Document;			
				
			string file = "";			
			
			using(var ofd = new OpenFileDialog())
			{
			    DialogResult result = ofd.ShowDialog();
			
			    if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(ofd.FileName))
			    {
			        file = ofd.FileName;
			    }
			}
			
			ModelPath path = ModelPathUtils.ConvertUserVisiblePathToModelPath(file);
				
			Document saveDoc = OpenDetached(this.Application, path);
			
			//ShowInfoOnOpenedWorksharedDocument( saveDoc );
			
			String serverPathRoot = this.Application
		      .GetRevitServerNetworkHosts().First();
		 	
			 
		    ModelPath modelPath = new ServerPath(
		      serverPathRoot,
		      "\\KGX1\\Consultants\\" + Path.GetFileName(file));
			
			SaveAsOptions options = new SaveAsOptions();
            options.OverwriteExistingFile = true;
            
		    WorksharingSaveAsOptions wsOptions
		      = new WorksharingSaveAsOptions();
		    
		    wsOptions.SaveAsCentral = true;
		    options.SetWorksharingOptions( wsOptions );
		    saveDoc.SaveAs( modelPath, options );
		 
		    ShowInfoOnOpenedWorksharedDocument( saveDoc );
		 
		    saveDoc.Close( false );			 
		}
		
		static Document OpenDetached(
	    Autodesk.Revit.ApplicationServices.Application app,
		ModelPath modelPath )
		{
			OpenOptions options = new OpenOptions();
						
			options.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets;
			
			return app.OpenDocumentFile(modelPath,options);
		}
		
		/// <summary>
		/// Show popup with info about worksets 
		/// and worksharing status
		/// </summary>
		static void ShowInfoOnOpenedWorksharedDocument(
		Document doc )
		{
		String documentName = doc.Title;
		bool isWorkshared = doc.IsWorkshared;
		
		FilteredWorksetCollector fwc
		  = new FilteredWorksetCollector( doc );
		
		fwc.OfKind( WorksetKind.UserWorkset );
		
		int wsCount = fwc.Count<Workset>();
		
		TaskDialog td = new TaskDialog(
		  "Opened document info" );
		
		td.MainInstruction
		  = "Application has opened the document "
		    + documentName;
		
		string mainContent = "Workshared: "
		  + isWorkshared
		  + ( isWorkshared
		    ? "\nassociated to: "
		      + ModelPathUtils
		        .ConvertModelPathToUserVisiblePath(
		          doc.GetWorksharingCentralModelPath() )
		    : "" )
		  + "\nWorkset count: " + wsCount + "\n"
		  + String.Join( "\n",
		    fwc.Select<Workset, String>(
		      ws => ws.Name + " - "
		        + ( ws.IsOpen ? "open" : "closed" ) ) );
		
		td.MainContent = mainContent;
		
		td.Show();
		}
		
		public void PopulateAreasToRooms()
		{			
            Document doc = ActiveUIDocument.Document;
            
            Level level = new FilteredElementCollector(doc)
            	.WhereElementIsNotElementType()
            	.OfCategory(BuiltInCategory.OST_Levels)
            	.Cast<Level>()
            	.Where(x => x.Name.Equals("L.04_A_FFL"))
            	.FirstOrDefault();            
            
            Level levelRooms = new FilteredElementCollector(doc)
            	.WhereElementIsNotElementType()
            	.OfCategory(BuiltInCategory.OST_Levels)
            	.Cast<Level>()
            	.Where(x => x.Name.Equals("L.04_C_FFL"))
            	.FirstOrDefault();
            
            //AreaScheme _scheme = document.GetElement(_area.get_Parameter(BuiltInParameter.AREA_SCHEME_ID).AsElementId()) as AreaScheme;
            //string _AreaSchemeName = _scheme.get_Parameter(BuiltInParameter.AREA_SCHEME_NAME).AsString();
                    
            List<Area> areas = new FilteredElementCollector(doc)
            	.WhereElementIsNotElementType()
            	.OfCategory(BuiltInCategory.OST_Areas)
            	.Cast<Area>()
            	.Where(x => x.Name.Contains("Bed"))
				.Where(x => x.Level.Name.Equals(level.Name))
            	.Where(x => (doc.GetElement(x.get_Parameter(BuiltInParameter.AREA_SCHEME_ID).AsElementId()) as AreaScheme)
            	       .get_Parameter(BuiltInParameter.AREA_SCHEME_NAME).AsString().Equals("Residential NSA"))
            	.ToList();
			            
            List<Room> rooms = new FilteredElementCollector(doc)
            	.WhereElementIsNotElementType()
            	.OfCategory(BuiltInCategory.OST_Rooms)
            	.Cast<Room>()
				.Where(x => x.Level.Name.Equals(levelRooms.Name))
            	.ToList();
                     
            TaskDialog.Show("Test", areas.Count.ToString() + ":" + rooms.Count.ToString());
            
            using(Transaction t = new Transaction(doc, "Populate Area information"))
            {
            	t.Start();
	            foreach(Area area in areas)
	            {
	            	if(area == null || !(area.Area > 0.1)) continue;
	            	foreach(Room room in rooms)
	            	{
	            		if(room == null) continue;
	            		
	            		if(AreaContains(area, (room.Location as LocationPoint).Point))
						{
	            			string unitType = area.Name[0] + " bedroom";
	            			string areaNumber = area.LookupParameter("Apartment Type").AsString();
	            			string number = String.Format("{0}{1}{2}", unitType, Environment.NewLine, areaNumber);
	            			room.Number = number;
	            			break;
						}   
	            	}
	            }
	            t.Commit();
            }            
		}
		
		private bool AreaContains(Area a, XYZ p1)
		{
			bool ret = false;
			var p = MaakPuntArray( a );
			PointInPoly pp = new PointInPoly();
			ret = pp.PolyGonContains( p, p1 );
			return ret;
		}
		private List<XYZ> MaakPuntArray(Area area)
		{
			SpatialElementBoundaryOptions opt
			  = new SpatialElementBoundaryOptions();
			
			opt.SpatialElementBoundaryLocation
			  = SpatialElementBoundaryLocation.Center;
			
			var boundaries = area.GetBoundarySegments(
			  opt );
			
			return MaakPuntArray( boundaries );
		}
		private List<XYZ> MaakPuntArray(IList<IList<BoundarySegment>> boundaries)
		{
			List<XYZ> puntArray = new List<XYZ>();
			foreach( var bl in boundaries )
			{
				foreach( var s in bl )
				{
					Curve c = s.GetCurve();
					AddToPunten(puntArray, c.GetEndPoint(0));
					AddToPunten(puntArray, c.GetEndPoint(1));
				}
			}
			puntArray.Add( puntArray.First() );
			return puntArray;
		}
		private void AddToPunten(List<XYZ> XYZarray,XYZ p1)
		{
			var p = XYZarray.Where(
			  c => Math.Abs( c.X - p1.X ) < 0.001
			    && Math.Abs( c.Y - p1.Y ) < 0.001 )
			  .FirstOrDefault();
			
			if( p == null )
			{
			  XYZarray.Add( p1 );
			}
		}
	}
	
	public class UVArray
  {
		
    public UV TOUV( XYZ point )
    {
      UV ret = new UV( point.X, point.Y );
      return ret;
    }
    
    List<UV> arrayPoints;
    public UVArray( List<XYZ> XYZArray )
    {
      arrayPoints = new List<UV>();
      foreach( var p in XYZArray )
      {
        arrayPoints.Add( TOUV(p) );
      }
    }

    public UV get_Item( int i )
    {
      return arrayPoints[i];
    }

    public int Size
    {
      get
      {
        return arrayPoints.Count;
      }
    }
  }

    public class PointInPoly
  {
    /// <summary>
    /// Determine the quadrant of a polygon vertex 
    /// relative to the test point.
    /// </summary>
    Quadrant GetQuadrant( UV vertex, UV p )
    {
      return ( vertex.U > p.U )
        ? ( ( vertex.V > p.V ) ? 0 : 3 )
        : ( ( vertex.V > p.V ) ? 1 : 2 );
    }

    /// <summary>
    /// Determine the X intercept of a polygon edge 
    /// with a horizontal line at the Y value of the 
    /// test point.
    /// </summary>
    double X_intercept( UV p, UV q, double y )
    {
      Debug.Assert( 0 != ( p.V - q.V ),
        "unexpected horizontal segment" );

      return q.U
        - ( ( q.V - y )
          * ( ( p.U - q.U ) / ( p.V - q.V ) ) );
    }

    void AdjustDelta(
      ref int delta,
      UV vertex,
      UV next_vertex,
      UV p )
    {
      switch( delta )
      {
        // make quadrant deltas wrap around:
        case 3: delta = -1; break;
        case -3: delta = 1; break;
        // check if went around point cw or ccw:
        case 2:
        case -2:
          if( X_intercept( vertex, next_vertex, p.V )
            > p.U )
          {
            delta = -delta;
          }
          break;
      }
    }
    
    public UV TOUV( XYZ point )
    {
      UV ret = new UV( point.X, point.Y );
      return ret;
    }

    public bool PolyGonContains( List<XYZ> xyZArray, XYZ p1 )
    {
      UVArray uva = new UVArray( xyZArray );
      return PolygonContains( uva, TOUV(p1) );
    }
	    
	
		public void WireInstanceParameters()
		{
		    UIDocument uidoc = ActiveUIDocument;
		    Document doc = ActiveUIDocument.Document;
		    
		    if(!doc.IsFamilyDocument) return;
		    
		    var wireFamily = uidoc.Selection.PickObject(ObjectType.Element, "Pick Object");
		    var fam = doc.GetElement(wireFamily) as FamilyInstance;
		    
		    SortedList<string, FamilyParameter> famParam = new SortedList<string, FamilyParameter>();
		    FamilyManager familyManager = doc.FamilyManager;
            FamilyType familyType = familyManager.CurrentType;

            string s = "";
            
            foreach (FamilyParameter fp in familyManager.Parameters)
            {
            	famParam.Add(fp.Definition.Name, fp);
            }
            
            s += Environment.NewLine;
            
        	using(Transaction t = new Transaction(doc, "Wire parameters"))
        	{
        		t.Start();
        		
	            foreach(Parameter param in fam.Parameters)
	            {
        			try{
		            	s = param.Definition.Name;
		            	FamilyParameter famParameter = null;
		            	if(famParam.TryGetValue(param.Definition.Name, out famParameter))
		        	    {
//		            		doc.FamilyManager.Set(famParameter, param.Id);
		            		doc.FamilyManager.AssociateElementParameterToFamilyParameter(param, famParameter);
		        	    }
		            	
		            }
		            catch(Exception ex)
		            {
//		            	TaskDialog.Show("Error", s + " : " +ex.Message);
		            } 
       		    }
        		t.Commit();
        	}           
		}		
		public void WireParameters()
		{
		    UIDocument uidoc = ActiveUIDocument;
		    Document doc = ActiveUIDocument.Document;
		    
		    if(!doc.IsFamilyDocument) return;
		    
		    var wireFamily = uidoc.Selection.PickObject(ObjectType.Element, "Pick Object");
		    var fam = doc.GetElement(wireFamily) as FamilyInstance;
		    
		    SortedList<string, FamilyParameter> famParam = new SortedList<string, FamilyParameter>();
		    FamilyManager familyManager = doc.FamilyManager;
            FamilyType familyType = familyManager.CurrentType;

            string s = "";
            
            foreach (FamilyParameter fp in familyManager.Parameters)
            {
            	famParam.Add(fp.Definition.Name, fp);
            }
            
            s += Environment.NewLine;
            
        	using(Transaction t = new Transaction(doc, "Wire parameters"))
        	{
        		t.Start();
        		
	            foreach(Parameter param in fam.Symbol.Parameters)
	            {
        			try{
		            	s = param.Definition.Name;
		            	FamilyParameter famParameter = null;
		            	if(famParam.TryGetValue(param.Definition.Name, out famParameter))
		        	    {
//		            		doc.FamilyManager.Set(famParameter, param.Id);
		            		doc.FamilyManager.AssociateElementParameterToFamilyParameter(param, famParameter);
		        	    }
		            	
		            }
		            catch(Exception ex)
		            {
//		            	TaskDialog.Show("Error", s + " : " +ex.Message);
		            } 
       		    }
        		t.Commit();
        	}           
		}
		public void PushParameter()
		{			
		    UIDocument uidoc = ActiveUIDocument;
		    Document doc = ActiveUIDocument.Document;
		    
		    if(!doc.IsFamilyDocument) return;
		    
		    var parameters = doc.FamilyManager.Parameters;
		    
		    List<FamilyParameter> famParameters = new List<FamilyParameter>();
		    FamilyParameter paramToPush = null;
		    string paramName = "Client";
		    
		    foreach(FamilyParameter famParam in parameters)
		    {
		    	famParameters.Add(famParam);
		    	if(famParam.Definition.Name.Equals(paramName)) paramToPush = famParam;
		    }
		    
		    if(paramToPush == null) return; //Couldn't find it
		    
		    var selection = uidoc.Selection.PickObject(ObjectType.Element, "Pick family to push to");	//Select a family
		    var famInstance = doc.GetElement(selection.ElementId) as FamilyInstance;
		    var family = famInstance.Symbol.Family;	//Get the family from the selection
		    var famDoc = doc.EditFamily(family);	//Open the family
		    
		    // Add the parameter to the family
		    try{
		    	using(Transaction ft = new Transaction(famDoc, "Push parameter"))
				{
					ft.Start();		    	
					famDoc.FamilyManager.AddParameter(paramToPush.Definition.Name, 
					                                  paramToPush.Definition.ParameterGroup,
					                                  paramToPush.Definition.ParameterType,
					                                  paramToPush.IsInstance);
							    	
					ft.Commit();
				}
		    }
		    catch(Exception){}
		   	// Load back the family	    
		    using(Transaction t = new Transaction(famDoc, "Push parameter"))
		    {
		    	t.Start();		    	
		    	family = famDoc.LoadFamily(doc, new FamilyOption());
		    	t.Commit();
		    }
		    // Associate the parameters
		    using(Transaction tw = new Transaction(doc, "Wire parameters"))
        	{
        		tw.Start();    
    			try{
	            	var famParameter = famInstance.LookupParameter(paramName);
            		doc.FamilyManager.AssociateElementParameterToFamilyParameter(famParameter, paramToPush);
	            }
	            catch(Exception ex) { } 
        		tw.Commit();
        	}   
		    
		    TaskDialog.Show("Test", "Done");
		}
		    
	    
	public void SwapViews()
		{
			UIDocument uidoc = this.ActiveUIDocument;
			Document doc = uidoc.Document;
			
			List<Tuple<View, Viewport>> result = new List<Tuple<View, Viewport>>();
			
			string level = "Ground Floor Plan";
			
			var views = new FilteredElementCollector(doc)
				.OfClass(typeof(View))
				.WhereElementIsNotElementType()
				.Cast<View>()
				.Where(v => v.LookupParameter("ViewSubgroup") != null
				       && v.LookupParameter("ViewSubgroup").HasValue
				       && v.LookupParameter("ViewSubgroup").AsString().Contains("20_100"))
				.Where(v => v.Name.Contains(level))
				.ToList();
			
//			RenameBack(doc, views);
//			return;
			
			var viewSheets = new FilteredElementCollector(doc)
				.OfClass(typeof(ViewSheet))	
				.Cast<ViewSheet>()
				.Where(v => v.Name.Contains(level))
				.ToList();
			
			var master = views.First(v => v.Name.Contains("Master"));				
			var viewsToDuplicate = views.Except(new List<View>{master});
			
			viewsToDuplicate = new List<View> {viewsToDuplicate.First(v => v.Name.Equals("Block D Ground Floor Plan - Core D1 South"))};
									
			RenameOldViews(doc, viewsToDuplicate);
			
			using(Transaction t = new Transaction(doc, "Swap Views"))
			{
				t.Start();
				foreach(var v in viewsToDuplicate)
				{
					var sheet = GetSheetViewFromViewandSheet(doc, viewSheets, v);
					if(sheet != null)
					{
						var target = SwapView(doc, master, sheet, v);	
						PropagateGrids(doc, v, target.Item1);
						CopyDetailedItem(doc, v, master);
						result.Add(target);
					}
				}
				t.Commit();
			}	
			AlignViews(doc, result);
		}		
		private void CopyDetailedItem(Document doc, View source, View target)
		{
			List<BuiltInCategory> builtInCats = new List<BuiltInCategory>();
			builtInCats.Add(BuiltInCategory.OST_Dimensions);
			builtInCats.Add(BuiltInCategory.OST_TextNotes);
			builtInCats.Add(BuiltInCategory.OST_AreaTags);
			 
			ElementMulticategoryFilter filter = new ElementMulticategoryFilter(builtInCats);
			
			var elements = new FilteredElementCollector(doc, source.Id)
				.WherePasses(filter)
				.ToElementIds();
			
			ElementTransformUtils.CopyElements(source, elements,target,null, new CopyPasteOptions());
		}
		private void AlignViews(Document doc, List<Tuple<View, Viewport>> views)
		{
			using(Transaction t = new Transaction(doc, "Align Views"))
			{
				t.Start();
				foreach(var v in views)
				{					
					XYZ loc_old = new XYZ(1.2,1,0);
					XYZ loc_new = (v.Item2.GetBoxOutline().MinimumPoint + v.Item2.GetBoxOutline().MaximumPoint) / 2;
					XYZ delta = loc_old - loc_new;
					ElementTransformUtils.MoveElement(doc, v.Item2.Id, delta);
				}
				t.Commit();
			}
		}
		private ViewSheet GetSheetViewFromViewandSheet(Document doc, List<ViewSheet> viewSheets, View view)
		{
			var sheet = viewSheets.First(vs => vs.GetAllPlacedViews().Any(v => doc.GetElement(v).Name.Equals(view.Name)));
			
			return sheet;
		}
		private void PropagateGrids(Document doc, View source, View target)
		{
			var grids = new FilteredElementCollector(doc, source.Id)
				.OfCategory(BuiltInCategory.OST_Grids)
				.Cast<Grid>()
				.ToList();
			
			var scopeBox = 	source.LookupParameter("Scope Box").AsElementId();			
			
			source.LookupParameter("Scope Box").Set(new ElementId(-1));
			source.CropBoxActive = false;			
			
			target.LookupParameter("Scope Box").Set(new ElementId(-1));
			target.CropBoxActive = false;
			
			foreach(var g in grids)
			{
				g.PropagateToViews(source, new HashSet<ElementId> {target.Id});
			}
						
			target.LookupParameter("Scope Box").Set(scopeBox);
			target.get_Parameter(BuiltInParameter.VIEWER_ANNOTATION_CROP_ACTIVE).Set(0);
			target.CropBoxActive = true;
			target.CropBoxVisible = false;
		}
		private Tuple<View, Viewport> SwapView(Document doc, View master, ViewSheet sheet, View view)
		{
			var viewport = doc.GetElement(sheet.GetAllViewports().First()) as Viewport;
			var name = view.Name;
			var scopeBox = view.LookupParameter("Scope Box").AsElementId();		
			
			var type = viewport.GetTypeId();
			
			// Duplicate the view as Dependent
			var swap = doc.GetElement(master.Duplicate(ViewDuplicateOption.AsDependent)) as View;
			swap.Name = name.Replace("_old","");
			swap.LookupParameter("Scope Box").Set(scopeBox);
			swap.CropBoxVisible = false;
			swap.CropBoxActive = true;
			swap.get_Parameter(BuiltInParameter.VIEWER_ANNOTATION_CROP_ACTIVE).Set(1);
			
			// Create the new Viewport
			var swappedView = Viewport.Create(doc, sheet.Id, swap.Id, new XYZ(0,0,0));

			// Change the Type of the Viewport			
			swappedView.ChangeTypeId(type);			
		
			// Finally delete the old viewport
			doc.Delete(viewport.Id);

			return new Tuple<View, Viewport> (swap,swappedView) ;
		}
		private void RenameBack(Document doc, IEnumerable<View> viewsToDuplicate)
		{
			using(Transaction t = new Transaction(doc, "Rename Old Views"))
			{
				t.Start();
				foreach(var v in viewsToDuplicate)
				{
					if(v.Name.Contains("_old"))
	                   v.Name = v.Name.Replace("_old", "");
				}
				t.Commit();
			}	
		}
		private void RenameOldViews(Document doc, IEnumerable<View> viewsToDuplicate)
		{
			using(Transaction t = new Transaction(doc, "Rename Old Views"))
			{
				t.Start();
				foreach(var v in viewsToDuplicate)
				{
					v.Name += "_old";					
					v.get_Parameter(BuiltInParameter.VIEWER_ANNOTATION_CROP_ACTIVE).Set(1);
				}
				t.Commit();
			}	
		}
    /// <summary>
    /// Determine whether given 2D point lies within 
    /// the polygon.
    /// 
    /// Written by Jeremy Tammik, Autodesk, 2009-09-23, 
    /// based on code that I wrote back in 1996 in C++, 
    /// which in turn was based on C code from the 
    /// article "An Incremental Angle Point in Polygon 
    /// Test" by Kevin Weiler, Autodesk, in "Graphics 
    /// Gems IV", Academic Press, 1994.
    /// 
    /// Copyright (C) 2009 by Jeremy Tammik. All 
    /// rights reserved.
    /// 
    /// This code may be freely used. Please preserve 
    /// this comment.
    /// </summary>
    public bool PolygonContains(
      UVArray polygon,
      UV point )
    {
      // initialize
      Quadrant quad = GetQuadrant(
        polygon.get_Item( 0 ), point );

      Quadrant angle = 0;

      // loop on all vertices of polygon
      Quadrant next_quad, delta;
      int n = polygon.Size;
      for( int i = 0; i < n; ++i )
      {
        UV vertex = polygon.get_Item( i );

        UV next_vertex = polygon.get_Item(
          ( i + 1 < n ) ? i + 1 : 0 );

        // calculate quadrant and delta from last quadrant

        next_quad = GetQuadrant( next_vertex, point );
        delta = next_quad - quad;

        AdjustDelta(
          ref delta, vertex, next_vertex, point );

        // add delta to total angle sum
        angle = angle + delta;

        // increment for next step
        quad = next_quad;
      }

      // complete 360 degrees (angle of + 4 or -4 ) 
      // means inside

      return ( angle == +4 ) || ( angle == -4 );

      // odd number of windings rule:
      // if (angle & 4) return INSIDE; else return OUTSIDE;
      // non-zero winding rule:
      // if (angle != 0) return INSIDE; else return OUTSIDE;
    }
  }
	
	class FamilyOption : IFamilyLoadOptions
	{
		public bool OnFamilyFound(bool familyInUse, out bool overwriteParameterValues)
		{
			overwriteParameterValues = true;
			return true;
		}

		public bool OnSharedFamilyFound(Family sharedFamily,
			bool familyInUse,
			out FamilySource source,
			out bool overwriteParameterValues )
		{
			source = FamilySource.Family;
			overwriteParameterValues = true;
			return true;
		}
	}
	
}
