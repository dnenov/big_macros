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
using System.IO;
using System.Linq;

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
			
		    string filename = Path.Combine("S:/15504_KXG/00_BR-PEERSYNC/02_BIM/01-WIP/01.07-Temp/01.07.04-Warnings/KXC-A-001-A-BH-M3-BaseBuilding@big.html");
		
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
		public void RoomTags()
		{
			Document doc = this.ActiveUIDocument.Document;
			
			List<ElementId> survivor = new List<ElementId>();
			List<ElementId> casualty = new List<ElementId>();
			
		    string filename = Path.Combine("S:/15504_KXG/00_BR-PEERSYNC/02_BIM/01-WIP/01.07-Temp/01.07.04-Warnings/KXC-A-001-A-BH-M3-BaseBuilding@big.html");
		    string message = "";
		    
		    IList<ElementId> ids = new List<ElementId>();		    
		    
		    bool next = false;
		    
		    List<Tuple<RoomTag, XYZ>> rtSet = new List<Tuple<RoomTag, XYZ>>();
            
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
		        	else
		        	{
		        		var firstIndex = line.IndexOf("Room Tag");
		        	
			        	if (firstIndex != line.LastIndexOf("Room Tag") && firstIndex != -1)
		        	    {	     		
							next = true;		        	    	
		        	    }	 
		        	}		        	      		            
		        }
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
		        message += String.Format("{0}{1}Overall {2} Room Tags moved.", Environment.NewLine, Environment.NewLine, rtSet.Count);
			}	    		
			
			TaskDialog.Show("RoomTag", message);
		}
		/// <summary>
		/// Removes Identical Instances Warning by deleting one of the duplicating elements
		/// </summary>
		public void IdenticalInstances()
		{
			Document doc = this.ActiveUIDocument.Document;
			
			List<ElementId> survivor = new List<ElementId>();
			List<ElementId> casualty = new List<ElementId>();
			
		    string filename = Path.Combine("S:/15504_KXG/00_BR-PEERSYNC/02_BIM/01-WIP/01.07-Temp/01.07.04-Warnings/KXC-A-001-A-BH-M3-BaseBuilding@big.html");
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
		        using (Transaction t = new Transaction(doc,"IdenticalInstances_RemoveElements"))
	            {		     
	                t.Start();	
	                doc.Delete(casualty);
	                t.Commit();
	                if (t.GetStatus() == TransactionStatus.Committed)
	                {	                	
		        		message += String.Format("{0}{1}Overall {2} Elements Deleted.", Environment.NewLine, Environment.NewLine, casualty.Count);
	                }
	                else{
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
			
		    string filename = Path.Combine("S:/15504_KXG/00_BR-PEERSYNC/02_BIM/01-WIP/01.07-Temp/01.07.04-Warnings/KXC-A-001-A-BH-M3-BaseBuilding@big.html");
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
				
				if (failureMessages.Count > 0)
				{
					return FailureProcessingResult.ProceedWithRollBack;
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
				return FailureProcessingResult.Continue;
			}
		}		
		public void DeleteAreaLines()
		{			
			Document doc = this.ActiveUIDocument.Document;
			
			List<ElementId> survivor = new List<ElementId>();
			List<ElementId> casualty = new List<ElementId>();
			
		    string filename = Path.Combine("S:/15504_KXG/00_BR-PEERSYNC/02_BIM/01-WIP/01.07-Temp/01.07.04-Warnings/KXC-A-001-A-BH-M3-BaseBuilding@big.html");
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
			
		    string filename = Path.Combine("S:/15504_KXG/00_BR-PEERSYNC/02_BIM/01-WIP/01.07-Temp/01.07.04-Warnings/KXC-A-001-A-BH-M3-BaseBuilding@big.html");
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
	}
}
