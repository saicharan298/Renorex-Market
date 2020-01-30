/*
 * Created by Ranorex
 * User: i-ray
 * Date: 12-11-2019
 * Time: 03:00
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
using System.IO;
//using WinForms = System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;

namespace Market_Preferences
{
	/// <summary>
	/// Creates a Ranorex user code collection. A collection is used to publish user code methods to the user code library.
	/// </summary>
	[UserCodeCollection]
	public class WebUpdates_UserCodeCollection
	{
		// You can use the "Insert New User Code Method" functionality from the context menu,
		// to add a new method with the attribute [UserCodeMethod].
		public static List<string> WebElements_Count=new List<string>();
		
		public static List<string> WebElements_options=new List<string>();
		public static List<string> DefaultValues_WebElements=new List<string>();
		/// <summary>
		/// This is a placeholder text. Please describe the purpose of the
		/// user code method here. The method is published to the user code library
		/// within a user code collection.
		/// </summary>
		[UserCodeMethod]
		public static void Fetching_WebUpdates_Options(string Buildname)
		{
			Buildname=Buildname.Replace(" ",string.Empty);
			Buildname=Buildname.ToLower();
			WebElements_Count.Clear();
			WebElements_options.Clear();
			DefaultValues_WebElements.Clear();
			int checkboxCount=0;
			int comboboxCount=0;
			int checkboxCount2=0;
			
			Ranorex.Container WebCon="/form[@classname='Window' and @orientation='None' and @processname='SmartFit' or @processname='SolusMax']/container[@automationid='PART_ExtendedScrollViewer']/container";
			IList<UIAutomation> WebCon_List=WebCon.FindChildren<UIAutomation>();
			for(int i=0;i<=WebCon_List.Count-1;i++)
			{
				if(WebCon_List[i].ControlType=="CheckBox" || WebCon_List[i].ControlType=="ComboBox")
				{
					WebElements_Count.Add(WebCon_List[i].ControlType);
				}
			}
			
			for(int k=0;k<=WebElements_Count.Count-1;k++)
			{
				if(WebElements_Count[k]=="CheckBox")
				{
					int c3=checkboxCount2+1;
					Ranorex.CheckBox checkboxData="/form[@classname='Window' and @orientation='None' and @processname='SmartFit' or @processname='SolusMax']/container/container/checkbox["+c3+"]";
					//	WebElements_options.Add(checkboxData);
					string Checkbox_text=checkboxData.Text;
					if(Checkbox_text=="Check for updates on a scheduled interval")
					{
						//	WebElements_options.Add("Check Updates Interval");
						WebElements_options.Add("Check For Updates Automatically");
					}
					else if(Checkbox_text=="Always download updates when available")
					{
						WebElements_options.Add("Download updates automatically");
					}
					else
					{
						WebElements_options.Add(Checkbox_text);
					}
					checkboxCount2++;
				}
				else if(WebElements_Count[k]=="ComboBox")
				{
					
					if(Buildname.Contains("smartfit") || Buildname.Contains("audigy"))
					{
						WebElements_options.Add("Check Updates Interval");
					}
					else if(Buildname.Contains("solusmax"))
					{
						WebElements_options.Add("Week");
					}
					//	WebElements_options.Add("Check Updates Interval");
				}
			}
			for(int j=0;j<=WebElements_Count.Count-1;j++)
			{
				if(WebElements_Count[j]=="CheckBox")
				{
					int c1=checkboxCount+1;
					Ranorex.CheckBox txtChecked="/form[@classname='Window' and @orientation='None' and @processname='SmartFit' or @processname='SolusMax']/container/container/checkbox["+c1+"]";
					if(txtChecked.Checked==true)
					{
						DefaultValues_WebElements.Add("Yes");
					}
					else
					{
						DefaultValues_WebElements.Add("No");
					}
					checkboxCount++;
				}
				else if(WebElements_Count[j]=="ComboBox")
				{
					int c2=comboboxCount+1;
					Ranorex.ComboBox txtCombText="/form[@classname='Window' and @orientation='None' and @processname='SmartFit' or @processname='SolusMax']/container/container/combobox["+c2+"]";
					DefaultValues_WebElements.Add(txtCombText.Text);
					comboboxCount++;
				}
			}
			
		}
		
		
		/// <summary>
		/// This is a placeholder text. Please describe the purpose of the
		/// user code method here. The method is published to the user code library
		/// within a user code collection.
		/// </summary>
		[UserCodeMethod]
		public static void WebupdatesExcel_Comaparison(string _marketname,string buildname)
		{
			int cellindex=0;
			string Excel_Preference="";
			buildname=buildname.Replace(" ",string.Empty);
			buildname=buildname.ToLower();
			string excelFinalPath="";
			if(buildname.Contains("smartfit"))
			{
				excelFinalPath = @"D:\TFS\FSW\TestSuites\Market_Preferences\Mkt Preferences_Smart Fit.xlsx";
			}
			else if(buildname.Contains("solusmax"))
			{
				excelFinalPath = @"D:\TFS\FSW\TestSuites\Market_Preferences\Mkt Preferences_Solus Max.xlsx";
			}
			else if(buildname.Contains("audigy"))
			{
				excelFinalPath = @"D:\TFS\FSW\TestSuites\Market_Preferences\Mkt Preferences_Audigy.xlsx";
			}
			//	string excelFinalPath = @"D:\TFS\FSW\TestSuites\Market_Preferences\Mkt Preferences_Smart Fit.xlsx";
			Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
			XLS.Workbook workBook = application.Workbooks.Open(excelFinalPath);
			Excel.Worksheet worksheet = (Excel.Worksheet)workBook.Worksheets[3];
			for(int i=0;i<=WebElements_options.Count-1;i++)
			{
				for(int j=30;j<=70;j++)
				{
					object cellValue = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[3, j]).Value;
					if(cellValue !=null)
					{
						string k1=cellValue.ToString();
						k1=k1.ToLower();
						string fswprefernce=WebElements_options[i];
						fswprefernce=fswprefernce.ToLower();
						k1=k1.Replace(" ",string.Empty);
						k1=Regex.Replace(k1, @"\([^)]*\)", "");
						fswprefernce=fswprefernce.Replace(" ",string.Empty);
						if(k1==fswprefernce)
						{
							cellindex=j;
							for(int k=2;k<=50;k++)
							{
								string Marketname="";
								object cellValue2 = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[k, 1]).Value;
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
									if(_marketname=="Audigy")
									{
										_marketname="United States";
									}
//									if(_marketname=="International Business")
//									{
//										_marketname="International";
//									}
								}
								Marketname=Marketname.Replace(" ",string.Empty).Replace("InternationalBusiness","International");
								_marketname=_marketname.Replace(" ",string.Empty).Replace("InternationalBusiness","International");
								if(_marketname==Marketname)
								{
									object cellValue3 = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[k,cellindex]).Value;
									Excel_Preference=cellValue3.ToString();
								}
							}
							string fswPreferences=DefaultValues_WebElements[i].Replace(" ",string.Empty);
							fswPreferences=fswPreferences.ToLower();
							fswPreferences=fswPreferences.Replace("weeks","week").Replace("-",string.Empty);
							Excel_Preference=Excel_Preference.ToLower();
							Excel_Preference=Excel_Preference.Replace(" ",string.Empty).Replace("weeks","week").Replace("-",string.Empty);
							if(Excel_Preference==fswPreferences)
							{
								Report.Success("WebUpdates of :"+WebElements_options[i]+" is :"+DefaultValues_WebElements[i] +"- Verified succesfully");
							}
							else if(Excel_Preference=="n/a")
							{
								if(DefaultValues_WebElements[i] != null)
								{
									Report.Warn("WebUpdates Feature of:"+DefaultValues_WebElements[i]+" -is not Available in Excel sheet");
								}
								
							}
							else
							{
								Report.Failure("WebUpdates  in Excel of:"+Excel_Preference+" -is different from FSW :"+DefaultValues_WebElements[i]+" -of Preference"+WebElements_options[i]);
							}
						}
					}
				}
			}
			workBook.Close(0);
			application.Quit();
			
			//var status = TestReport.CurrentTestSuiteActivity.Status;
			
		}
	}
}
