/*
 * Created by Ranorex
 * User: i-ray
 * Date: 25-11-2019
 * Time: 04:27
 * 
 * To change this template use Tools > Options > Coding > Edit standard headers.
 */
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Threading;
using WinForms = System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using XLS=Microsoft.Office.Interop.Excel;
using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;

namespace Market_Preferences
{
	/// <summary>
	/// Creates a Ranorex user code collection. A collection is used to publish user code methods to the user code library.
	/// </summary>
	[UserCodeCollection]
	public class Bi_modal_UserGuide_UserCodeCollection
	{
		// You can use the "Insert New User Code Method" functionality from the context menu,
		// to add a new method with the attribute [UserCodeMethod].
		
		
		/// <summary>
		/// This is a placeholder text. Please describe the purpose of the
		/// user code method here. The method is published to the user code library
		/// within a user code collection.
		/// </summary>
		[UserCodeMethod]
		public static void Changing_Bimodal_userGuide()
		{
			Delay.Seconds(3);
			Ranorex.Button btnBimodal="/form[@classname='Window' and @orientation='None' and @processname='SmartFit']/container/button[8]";
			if(btnBimodal.Pressed==false)
			{
				Mouse.Click(btnBimodal);
				Delay.Seconds(3);
				Ranorex.Button btn_Save="/form[@classname='Window' and @orientation='None' and @processname='SmartFit']/button[4]";
				Mouse.Click(btn_Save);
			}
			else
			{
				Ranorex.Button btn_X_preference="/form[@classname='Window' and @orientation='None' and @processname='SmartFit']/button[@name='Close']";
				Mouse.Click(btn_X_preference);
			}
		}
		
		
		/// <summary>
		/// This is a placeholder text. Please describe the purpose of the
		/// user code method here. The method is published to the user code library
		/// within a user code collection.
		/// </summary>
		[UserCodeMethod]
		public static void Click_On_BiModaluserGuide(string _marketname,string buildname)
		{
			string	FSW_PageURL="";
			Ranorex.ContextMenu HelpMenuItems="/contextmenu[@processname='SmartFit' and @win32ownerwindowlevel='1']";
			IList<MenuItem> All_MenuItems=HelpMenuItems.FindChildren<MenuItem>();
			for(int k=0;k<=All_MenuItems.Count-1;k++)
			{
				int k1=k+1;
				Ranorex.MenuItem HelpMenuText="/contextmenu[@processname='SmartFit' and @win32ownerwindowlevel='1']/menuitem["+k1+"]";
				string txt=HelpMenuText.Text;
				if(txt !=null)
				{
					if(txt=="Bimodal Fitting Guide")
					{
						Mouse.Click(HelpMenuText);
						Delay.Seconds(6);
						IList<Ranorex.WebDocument> AllDoms=Host.Local.FindChildren<Ranorex.WebDocument>();
						foreach (WebDocument myDom in AllDoms)
						{

							if(myDom.Page.Contains("Bimodal"))
							{
								FSW_PageURL=myDom.PageUrl;
								myDom.Close();
								break;
							}
						}
						break;
					}
				}
			}
			
			
			
			Ranorex.MenuItem HelpButtonclick="/form[@title~'Smart Fit' or @title~'Solus Max' or @title~'Audigy']/contextmenu/menuitem[@automationid='MenuAutomationIds.HelpAction']";
			string Excel_Preference="";
			string 	excelFinalPath="";
			buildname=buildname.Replace(" ",string.Empty);
			buildname=buildname.ToLower();
			if(buildname.Contains("smartfit"))
			{
				excelFinalPath = @"D:\TFS\FSW\TestSuites\Market_Preferences\Mkt Preferences_Smart Fit.xlsx";
			}
			else if(buildname.Contains("audigy"))
			{
				excelFinalPath = @"D:\TFS\FSW\TestSuites\Market_Preferences\Mkt Preferences_Audigy.xlsx";
			}
			//string excelFinalPath = @"D:\TFS\FSW\TestSuites\Market_Preferences\Mkt Preferences_Smart Fit.xlsx";
			Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
			XLS.Workbook workBook = application.Workbooks.Open(excelFinalPath);
			Excel.Worksheet worksheet = (Excel.Worksheet)workBook.Worksheets[3];
			for(int i=2;i<=70;i++)
			{
				object cellValue = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[3, i]).Value;
				if(cellValue !=null)
				{
					string _Excel_Bimodal_Value=cellValue.ToString();
					if(_Excel_Bimodal_Value=="Bimodal User Guide")
					{
						for(int j=2;j<=50;j++)
						{
							string Marketname="";
							object cellValue2 = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[j, 1]).Value;
							if(cellValue2!=null)
							{
								Marketname=cellValue2.ToString();
								_marketname=_marketname.Replace(" ",string.Empty);
								Marketname=Marketname.Replace(" ",string.Empty);
								if(Marketname=="UK")
								{
									Marketname="United Kingdom";
								}
								if(Marketname=="USA")
								{
									Marketname="United States";
								}
							}
//							Marketname=Marketname.Replace(" ",string.Empty);
//							_marketname=_marketname.Replace(" ",string.Empty);
							Marketname=Marketname.Replace(" ",string.Empty).Replace("InternationalBusiness","International");
							_marketname=_marketname.Replace(" ",string.Empty).Replace("InternationalBusiness","International");
							if(_marketname!=null)
							{
								object cellValue3 = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[j,i]).Value;
								if(cellValue3 !=null )
								{
									Excel_Preference=cellValue3.ToString();
									
								}
							}
						}
					}
				}
			}
			string Excel_Preference2=Excel_Preference;
			Excel_Preference=Excel_Preference.Substring(0,Excel_Preference.LastIndexOf(":"));
			Excel_Preference=Excel_Preference.Replace("embedded",string.Empty).Replace("PDF",string.Empty).Replace(" ",string.Empty);
			
			if(FSW_PageURL.Contains(Excel_Preference))
			{
				Report.Success("Bimodal User Guide Validation is success");
			}
			else
			{
				Report.Failure("Bimodal User Guide of Excel file :"+Excel_Preference2 +" -is different from Bimodal User Guide edirection PDF :"+FSW_PageURL);
			}
			
			workBook.Close(0);
			application.Quit();
		}
	}
}
