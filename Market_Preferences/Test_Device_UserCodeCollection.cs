/*
 * Created by Ranorex
 * User: i-ray
 * Date: 11-11-2019
 * Time: 06:14
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
	public class Test_Device_UserCodeCollection
	{
		// You can use the "Insert New User Code Method" functionality from the context menu,
		// to add a new method with the attribute [UserCodeMethod].
		public static List<string> TestDeviceOption_list=new List<string>();
		public static List<string> TestDeviceDefault_list=new List<string>();
		/// <summary>
		/// This is a placeholder text. Please describe the purpose of the
		/// user code method here. The method is published to the user code library
		/// within a user code collection.
		/// </summary>
		[UserCodeMethod]
		public static void Fetching_Test_Device_Options()
		{
			TestDeviceOption_list.Clear();
			TestDeviceDefault_list.Clear();
			Ranorex.Container TestDevice_Elements="/form[@classname='Window' and @orientation='None' and @processname='SmartFit' or @processname='SolusMax']/container[@automationid='PART_ExtendedScrollViewer']/container";
			IList<UIAutomation> Elements_List=TestDevice_Elements.FindChildren<UIAutomation>();
			for(int i=0;i<=Elements_List.Count-1;i++)
			{
				if(Elements_List[i].Name=="P1" || Elements_List[i].Name=="P2" || Elements_List[i].Name=="P3" || Elements_List[i].Name=="P4")
				{
					string _testdevice_Option=Elements_List[i].Name;
					_testdevice_Option=_testdevice_Option.Replace("P","Program ");
					TestDeviceOption_list.Add(_testdevice_Option);
				}
			}
			IList<ComboBox> Combobox_List=TestDevice_Elements.FindChildren<ComboBox>();
			for(int j=0;j<=Combobox_List.Count-1;j++)
			{
				int _combobox_number=j+1;
				Ranorex.ComboBox Comb_Text="/form[@classname='Window' and @orientation='None' and @processname='SmartFit' or @processname='SolusMax']/container/container/combobox["+_combobox_number+"]";
				string _combobox_text=Comb_Text.Text;
				_combobox_text=_combobox_text.Replace("Full On Gain","FOG").Replace("Reference Test Gain","RTG").Replace("Telecoil","T-Coil");
				TestDeviceDefault_list.Add(_combobox_text);
			}
		}
		
		
		/// <summary>
		/// This is a placeholder text. Please describe the purpose of the
		/// user code method here. The method is published to the user code library
		/// within a user code collection.
		/// </summary>
		[UserCodeMethod]
		public static void TestDevice_ExcelComparison(string _marketname,string buildname)
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
			for(int i=0;i<=TestDeviceOption_list.Count-1;i++)
			{
				for(int j=30;j<=70;j++)
				{
					object cellValue = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[3, j]).Value;
					if(cellValue !=null)
					{
						string k1=cellValue.ToString();
						k1=k1.ToLower();
						string fswprefernce=TestDeviceOption_list[i];
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
							string fswPreferences=TestDeviceDefault_list[i].Replace(" ",string.Empty);
							Excel_Preference=Excel_Preference.Replace("Full On Gain","FOG").Replace("Reference Test Gain","RTG").Replace("Telecoil","T-Coil");
							fswPreferences=fswPreferences.ToLower();
							Excel_Preference=Excel_Preference.ToLower();
							Excel_Preference=Excel_Preference.Replace(" ",string.Empty);
							if(Excel_Preference=="rtg")
							{
								Excel_Preference=Excel_Preference+"(ansi)";
							}
							if(Excel_Preference==fswPreferences)
							{
								Report.Success("Test Device of:"+TestDeviceOption_list[i]+" is :"+TestDeviceDefault_list[i] +"- Verified succesfully");
							}
							else
							{
								Report.Failure("Test Device in Excel of:"+Excel_Preference+" -is different from FSW :"+TestDeviceDefault_list[i]+" -of Preference"+TestDeviceOption_list[i]);
							}
						}
						
						
					}
					
				}
			}
			workBook.Close(0);
			application.Quit();
			
		}
		
	}
}
