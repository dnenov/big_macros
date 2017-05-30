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
		
		public void DeleteUnusedSharedParameter()
		{
			UIDocument uidoc = ActiveUIDocument;
			Document doc = uidoc.Document;
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

			Definition def = null;
			/*
			while (it.MoveNext())
			{
				parameters.RemoveAll(x => x.Name.Equals(it.Key.Name));
			}
			*/
			string e = parameters.Count.ToString();

			builder.Append(String.Format("Number of all Shared Parameters: {0}{1}", s, Environment.NewLine));
			builder.Append(String.Format("Number of unused Shared Parameters: {0}{1}", e, Environment.NewLine));

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
						doc.Delete(parameters.Select(x => x.Id).ToArray());
						t.Commit();         
				    } 	
				    TaskDialog.Show("Result", String.Format("{0} paramters have been deleted.", e));
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
	}
}
