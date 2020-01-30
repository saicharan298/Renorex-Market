/*
 * Created by Ranorex
 * User: i-ray
 * Date: 07-11-2019
 * Time: 05:00
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
using System.Linq;
using System.Diagnostics;
using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;

namespace Market_Preferences
{
	/// <summary>
	/// Creates a Ranorex user code collection. A collection is used to publish user code methods to the user code library.
	/// </summary>
	[UserCodeCollection]
	public class Machine_Preference_UserCode
	{
		// You can use the "Insert New User Code Method" functionality from the context menu,
		// to add a new method with the attribute [UserCodeMethod].
		public static	List<string> FSW_MachinePreference_List=new List<string>();
		public static	List<string> DefaultValues_MachinePreference_List=new List<string>();
		public static	List<string> ElementList_Count=new List<string>();
		/// <summary>
		/// This is a placeholder text. Please describe the purpose of the
		/// user code method here. The method is published to the user code library
		/// within a user code collection.
		/// </summary>
		[UserCodeMethod]
		public static void Fetching_Machine_Preference()
		{
			int _ComboBoxCount=0;
			FSW_MachinePreference_List.Clear();
			DefaultValues_MachinePreference_List.Clear();
			Ranorex.Container Container_Machine_Prefernces="/form[@classname='Window' and @orientation='None' and @processname='SmartFit' or @processname='SolusMax']/container[@automationid='PART_ExtendedScrollViewer']";
			IList<UIAutomation> Machine_Preferncest_list=Container_Machine_Prefernces.FindChildren<UIAutomation>();
			for(int i=0;i<=Machine_Preferncest_list.Count-4;i++)
			{
				string name=Machine_Preferncest_list[i].Name;
				if(name !=null)
				{
					//	fswprefernce=fswprefernce.Replace(" ",string.Empty);
					//name=name.Replace(":",string.Empty);
					name=name.Replace(":",string.Empty);
					if(!(name ==null || name =="Test" || name =="Parameters" || name =="Yes"))
					{
						//	FSW_MachinePreference_List.Add(Machine_Preferncest_list[i].Name);
						if(name=="Pediatric Default Target Rule")
						{
							FSW_MachinePreference_List.Add("Default Pediatric Fitting Rule");
						}
						else if(name=="Default Experience")
						{
							FSW_MachinePreference_List.Add("Default Experience Level");
						}
						else if(name=="Show GN Online Services System Tray Icon")
						{
							FSW_MachinePreference_List.Add("System Tray Visibility");
						}
						else if(name=="Programming Interface")
						{
							FSW_MachinePreference_List.Add("Default Programming Interface");
						}
						else
						{
							FSW_MachinePreference_List.Add(name);
						}
					}
				}
			}
			ElementList_Count.Clear();
			for(int j=0;j<=Machine_Preferncest_list.Count-4;j++)
			{
				string elementname=Machine_Preferncest_list[j].ControlType;
				if(elementname=="ComboBox" || elementname=="CheckBox" || elementname=="Button")
				{
					if(elementname=="Button")
					{
						if(Machine_Preferncest_list[j].Name=="Yes")
						{
							ElementList_Count.Add(elementname);
						}
					}
					else
					{
						ElementList_Count.Add(elementname);
					}
				}
			}
			
			for(int k=0;k<=ElementList_Count.Count-1;k++)
			{
				if(ElementList_Count[k]=="ComboBox")
				{
					int m=_ComboBoxCount+1;
					//	int m=
					Ranorex.ComboBox Comb_text="/form[@classname='Window' and @orientation='None' and @processname='SmartFit' or @processname='SolusMax']/container/combobox["+m+"]";
					string defvalue=Comb_text.Text;
					if(defvalue=="DSLv5 - Pediatric")
					{
						DefaultValues_MachinePreference_List.Add("DSLv5b - Pediatric");
					}
					else
					{
						DefaultValues_MachinePreference_List.Add(defvalue);
					}
					_ComboBoxCount++;
				}
				else if(ElementList_Count[k]=="CheckBox" || ElementList_Count[k]=="Button")
				{
					if(ElementList_Count[k]=="CheckBox")
					{
						Ranorex.CheckBox checkedCondition="/form[@classname='Window' and @orientation='None' and @processname='SmartFit' or @processname='SolusMax']/container/checkbox[@automationid='PART_UsePinCodeCheckBox']";
						if(checkedCondition.Checked==true)
						{
							DefaultValues_MachinePreference_List.Add("Yes");
						}
						else
						{
							DefaultValues_MachinePreference_List.Add("No");
						}
					}
					else if(ElementList_Count[k]=="Button")
					{
						Ranorex.Button btnpressedCondition="/form[@classname='Window' and @orientation='None' and @processname='SmartFit' or @processname='SolusMax']/container/button[@automationid='PART_SegmentedToggleButton']";
						if(btnpressedCondition.Pressed==true)
						{
							DefaultValues_MachinePreference_List.Add("Yes");
						}
						else
						{
							DefaultValues_MachinePreference_List.Add("No");
						}
					}
				}
			}
			
		}
		
		
		/// <summary>
		/// This is a placeholder text. Please describe the purpose of the
		/// user code method here. The method is published to the user code library
		/// within a user code collection.
		/// </summary>
		[UserCodeMethod]
		public static void Excel_Machine_Preference_Values(string _marketname,string buildname)
		{
			
			int cellindex=0;
			string Excel_Preference="";
			string excelFinalPath="";
			buildname=buildname.Replace(" ",string.Empty);
			buildname=buildname.ToLower();
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
			for(int i=0;i<=FSW_MachinePreference_List.Count-1;i++)
			{
				for(int j=20;j<=60;j++)
				{
					object cellValue = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[3, j]).Value;
					if(cellValue !=null)
					{
						string k1=cellValue.ToString();
						k1=k1.ToLower();
						string fswprefernce=FSW_MachinePreference_List[i];
						fswprefernce=fswprefernce.ToLower();
						k1=k1.Replace(" ",string.Empty);
						//	text = Regex.Replace(text, @"\([^)]*\)", "");
						k1=Regex.Replace(k1, @"\([^)]*\)", "");
						fswprefernce=fswprefernce.Replace(" ",string.Empty);
						if(k1.Contains(fswprefernce))
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
									//	string name=C_List[i].Name.Replace(":",string.Empty)
									break;
									
								}
							}
							Excel_Preference=Excel_Preference.Replace(" ",string.Empty).Replace("User",string.Empty).Replace("-",string.Empty);
							//	text = Regex.Replace(text, @"\([^)]*\)", "");
							//	Excel_Preference=Regex.Replace(Excel_Preference,@"\([^)]*\)", "");
							string fswPreferences=DefaultValues_MachinePreference_List[i].Replace(" ",string.Empty).Replace("User",string.Empty).Replace("-",string.Empty);
							
							Excel_Preference=Excel_Preference.ToLower();
							Excel_Preference=Excel_Preference.Replace("experiencednonlinear","experiencenonlinear");
							if(Excel_Preference=="comfor")
							{
								Excel_Preference="comfort";
							}
							if(Excel_Preference=="trialpeiod(4weeks)")
							{
								Excel_Preference="trialperiod(4weeks)";
							}
							fswPreferences=fswPreferences.ToLower();
							fswPreferences=fswPreferences.Replace("2cccoupler","2cc");
							Excel_Preference=Excel_Preference.Replace("2cccoupler","2cc");
							if(Excel_Preference==fswPreferences)
							{
								Report.Success("Machine Prefernce of -"+FSW_MachinePreference_List[i]+" is :"+DefaultValues_MachinePreference_List[i] +"- Verified succesfully");
							}
//							else if(Excel_Preference=="notapplicable")
//							{
//								Report.Success("Machine Preference of -"+FSW_MachinePreference_List[i] +" :is not applicable in Excel file");
//							}
							else
							{
								Report.Failure("Machine Preference in Excel :"+Excel_Preference+" -is different from FSW :"+DefaultValues_MachinePreference_List[i]+" -of Preference"+FSW_MachinePreference_List[i]);
							}
						}
					}
					//
				}
			}
			
			workBook.Close(0);
			application.Quit();
			
		}
		
	}
}
