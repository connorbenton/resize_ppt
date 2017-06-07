/*
 * Created by SharpDevelop.
 * User: cbenton
 * Date: 5/26/2017
 * Time: 3:28 PM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
/*
 * Created by SharpDevelop.
 * User: cbenton
 * Date: 5/26/2017
 * Time: 11:08 AM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using Microsoft.Office.Core;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
using System.IO;
using System.Linq;


namespace testresize
{
	class Program
	{
		
		[STAThread]
		public static void Main()
		{
			
			
			// Console.WriteLine("Hello World!");
			
			// TODO: Implement Functionality Here
			
//			String strTemplate, strPic;
//			strTemplate = 
//			  "C:\\Program Files\\Microsoft Office\\Templates\\Presentation Designs\\Blends.pot";
//			strPic = "C:\\Windows\\Blue Lace 16.bmp";
//			bool bAssistantOn;
			
			
			PowerPoint.Application objApp;
			PowerPoint.Presentations objPresSet;
			PowerPoint._Presentation objPres;
			PowerPoint._Presentation objPresNew;
			PowerPoint.Slides objSlides;
			PowerPoint._Slide objSlide;
			PowerPoint.TextRange objTextRng;
			PowerPoint.Shapes objShapes;
			PowerPoint.Shape objShape;
			PowerPoint.SlideShowWindows objSSWs;
			PowerPoint.SlideShowTransition objSST;
			PowerPoint.SlideShowSettings objSSS;
			PowerPoint.SlideRange objSldRng;
			PowerPoint.ShapeRange objShpRng;
			string pptPresPath;
			string pptPresName;
			string pptCorrectedPresPath;

			
			//Pick presentation for resizing.
			OpenFileDialog openFileDialog1 = new OpenFileDialog();
			openFileDialog1.Filter = "PowerPoint Files|*.pptx;*.ppt;*.pptm";
			openFileDialog1.Title = "Select the presentation for resizing";
			
			if (openFileDialog1.ShowDialog() == DialogResult.OK)
		    	pptPresPath = openFileDialog1.FileName;
		    else
		    	pptPresPath = string.Empty;
		    	Application.Exit();
		    
		    pptPresName = Path.GetFileNameWithoutExtension(pptPresPath);
		    
		    pptCorrectedPresPath = Path.GetDirectoryName(pptPresPath) + "\\" + pptPresName + "_corrected" + Path.GetExtension(pptPresPath);
		    
		    if (File.Exists(pptCorrectedPresPath))
		    {
		    	DialogResult dialogResult1 = MessageBox.Show("A corrected version of this file appears to exist already in this folder. Do you want to proceed (this will overwrite existing corrected version)?", "Overwrite corrected file", MessageBoxButtons.YesNo);
			    if(dialogResult1 == DialogResult.No)
			    	System.Environment.Exit(0);
		    }
		    
			//Open selected presentation.
			objApp = new PowerPoint.Application();
			objApp.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
			objPresSet = objApp.Presentations;
			objPres = objPresSet.Open(pptPresPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue);
			objSlides = objPres.Slides;
			
			//Resize Section
			objPresNew = objPresSet.Add(Microsoft.Office.Core.MsoTriState.msoTrue);
			
//			try{
			
			objPresNew.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
			var seq = Enumerable.Range(1,objPres.Slides.Count).ToArray();
			objPres.Slides.Range(seq).Copy();
//			objPresNew.Slides[1].Select();
			objPresNew.Slides[1].Application.CommandBars.ExecuteMso("PasteSourceFormatting");
			Application.DoEvents();
			objPresNew.Slides[1].Delete();
//			objApp.ActiveWindow.Selection.Unselect();
			
			
			//Resize stuck shapes
			foreach (PowerPoint._Slide d in objPresNew.Slides) {
				foreach (PowerPoint.Shape e in d.Shapes) {
//					Console.WriteLine(e.Name);
					PowerPoint.Shape eOld = objPres.Slides[d.SlideNumber].Shapes[e.Name];
					if (e.Type == Microsoft.Office.Core.MsoShapeType.msoAutoShape && e.Height != eOld.Height)
					{
						e.Height = eOld.Height;
					}
				}
			}
			
			//Delete notes master
			var seq2 = Enumerable.Range(1,objPresNew.SlideMaster.Shapes.Count).ToArray();
			objPresNew.SlideMaster.Shapes.Range(seq2).Delete();
			
			//Save new pres
			objPresNew.SaveAs(pptCorrectedPresPath, PowerPoint.PpSaveAsFileType.ppSaveAsDefault, Microsoft.Office.Core.MsoTriState.msoFalse);

//			objApp.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
			
//			objPresNew.NewWindow();
//			Set pptApp = GetObject(Class:="PowerPoint.Application")
//			Set pptPres = pptApp.ActivePresentation
//			pptPath = pptPres.Path
//			newsaveOld = pptPath & "\" & fso.GetBaseOld(pptPres.Name) & "_corrected"
//			
//			With Presentations.Add
//			End With
//			
//			Set pptnewPres = pptApp.ActivePresentation
//			
//			i = 1
//			
//			Do While i <= pptPres.Slides.Count
//			
//			pptnewPres.Slides.Add Index:=pptnewPres.Slides.Count + 1, Layout:=ppLayoutBlank
//			
//			Set myRange = pptPres.Slides(i).Shapes.Range
//			Set newslide = pptnewPres.Slides(i)
//			newslide.Select
//			myRange.Copy
//			pptnewPres.Application.CommandBars.ExecuteMso ("PasteSourceFormatting")
//			DoEvents
//			i = i + 1
//			
//			Loop
//			
//			pptnewPres.SaveAs (newsaveOld)
			
//			 Console.WriteLine(pptPresPath);
//			 Console.WriteLine(pptPresName);
//			 Console.WriteLine(pptCorrectedPresPath);
//			 Console.WriteLine("Press any key to continue . . . ");
//			 Console.ReadKey(true);
//		
//			} catch(Exception exc){
//				objPresNew.Close();
//				objPres.Close();
//			}
		}
	}
}