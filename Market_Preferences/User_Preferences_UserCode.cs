/*
 * Created by Ranorex
 * User: i-ray
 * Date: 25-10-2019
 * Time: 04:47
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
//using WinForms = System.Windows.Forms;
//using XLS = Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using XLS=Microsoft.Office.Interop.Excel;
using System.IO;
//using WinForms = System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using System.Linq;
using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;

namespace Market_Preferences
{
	/// <summary>
	/// Creates a Ranorex user code collection. A collection is used to publish user code methods to the user code library.
	/// </summary>
	[UserCodeCollection]
	public class User_Preferences_UserCode
	{
		// You can use the "Insert New User Code Method" functionality from the context menu,
		// to add a new method with the attribute [UserCodeMethod].
		public static	List<string> FSW_UserPreference_List=new List<string>();
		public static	List<string> FSW_UserPreference_Default_List=new List<string>();
		
		public static	List<string> Excel_UserPreference_List=new List<string>();
		
		/// <summary>
		/// This is a placeholder text. Please describe the purpose of the
		/// user code method here. The method is published to the user code library
		/// within a user code collection.
		/// </summary>
		[UserCodeMethod]
		public static void Default_Language_Comparison(string _marketname,string buildname)
		{
			int cellindex=0;
			string Excel_languageCode="";
			XmlDocument doc = new XmlDocument();
			string excelFinalPath="";
			buildname=buildname.Replace(" ",string.Empty);
			buildname=buildname.ToLower();
			if(buildname.Contains("smartfit"))
			{
				doc.Load(@"C:\Program Files (x86)\ReSound\SmartFit\Data\Market\market.pref");
			}
			else if(buildname.Contains("solusmax"))
			{
				doc.Load(@"C:\Program Files (x86)\Beltone\SolusMax\Data\Market\market.pref");
			}
			else if(buildname.Contains("audigy"))
			{
				doc.Load(@"C:\Program Files (x86)\Audigy\Audigy2\Data\Market\market.pref");
			}
			
			XmlNodeList nodelist = doc.GetElementsByTagName("Pref");
			string XMl_Language_code=nodelist[0].Attributes[2].InnerText;
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
			//string excelFinalPath = @"D:\TFS\FSW\TestSuites\Market_Preferences\Mkt Preferences_Smart Fit.xlsx";
			Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
			XLS.Workbook workBook = application.Workbooks.Open(excelFinalPath);
			Excel.Worksheet worksheet = (Excel.Worksheet)workBook.Worksheets[3];
			for(int i=3;i<=30;i++)
			{
				
				object cellValue = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[3, i]).Value;
				string k=cellValue.ToString();
				if(k=="Default Language XLF File reference" || k=="Default Language XLF File Reference")
				{
					cellindex=i;
				}

			}
			
			for(int j=2;j<=50;j++)
			{
				string Marketname="";
				object cellValue2 = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[j, 1]).Value;
				if(cellValue2!=null )
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
				}
				Marketname=Marketname.Replace(" ",string.Empty).Replace("InternationalBusiness","International");
				_marketname=_marketname.Replace(" ",string.Empty).Replace("InternationalBusiness","International");
				if(_marketname==Marketname)
				{
					object cellValue3 = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[j,cellindex]).Value;
					Excel_languageCode=cellValue3.ToString();
					break;
				}
				
			}
			if(XMl_Language_code==Excel_languageCode)
			{
				Report.Success("Default language :"+Excel_languageCode+" is Verified");
			}
			else
			{
				Report.Failure("Default language in Fitting Software is "+XMl_Language_code+" and Default language in Excel file is :"+Excel_languageCode);
			}
			workBook.Close(0);
			application.Quit();
			
			
		}
		/// <summary>
		/// This is a placeholder text. Please describe the purpose of the
		/// user code method here. The method is published to the user code library
		/// within a user code collection.
		/// </summary>
		[UserCodeMethod]
		public static void Fetching_UserPreference(string buildname)
		{
			FSW_UserPreference_List.Clear();
			buildname=buildname.Replace(" ",string.Empty);
			buildname=buildname.ToLower();
			if(buildname.Contains("smartfit"))
			{
				
				Ranorex.Container Container_UserPrefernces="/form[@classname='Window' and @orientation='None' and @processname='SmartFit' or @processname='SolusMax']/container[@automationid='PART_ExtendedScrollViewer']";
				IList<Container> txtContainer=Container_UserPrefernces.FindChildren<Container>();

				IList<UIAutomation> C_List=Container_UserPrefernces.FindChildren<UIAutomation>();
				for(int i=0;i<=C_List.Count-1;i++)
				{
					if(C_List[i].ControlType=="Group")
					{
						if(C_List[i].Name !="Default Environmental Programs:")
						{
							string fswPreference=C_List[i].Name;
							if(fswPreference=="Remote Fine-tuning")
							{
								fswPreference="Is RFT Visible in Preference Window";
								FSW_UserPreference_List.Add(fswPreference);
							}
							else if(fswPreference=="Remote Hearing Aid Update")
							{
								fswPreference="Is RFU Visible in Preference Window";
								FSW_UserPreference_List.Add(fswPreference);
							}
							else
							{
								FSW_UserPreference_List.Add(fswPreference);
							}

						}
						for(int j=0;j<=txtContainer.Count-1;j++)
						{
							if(C_List[i].Name==txtContainer[j].Caption)
							{
								string containerPreference="";
								int k=j+1;
								if(k==1)
								{
									Ranorex.Container ContainetTxT="/form[@classname='Window' and @orientation='None' and @processname='SmartFit']/container[@automationid='PART_ExtendedScrollViewer']/container["+k+"]";
									IList<Text> innerText=ContainetTxT.FindChildren<Text>();
									
									for(int r=0;r<=innerText.Count-2;r++)
									{
										containerPreference=innerText[r].Caption;
										if(containerPreference=="1." || containerPreference=="2." || containerPreference=="3." || containerPreference=="4.")
										{
											containerPreference=containerPreference.Replace(".",string.Empty);
											containerPreference="Default Program " +containerPreference;
											FSW_UserPreference_List.Add(containerPreference);
										}
										else
										{
											FSW_UserPreference_List.Add(containerPreference);
										}
										
									}
									
									
								}
								else if(k==2)
								{
									Ranorex.Container ContainetTxT="/form[@classname='Window' and @orientation='None' and @processname='SmartFit']/container[@automationid='PART_ExtendedScrollViewer']/container["+k+"]";
									IList<Text> innerText=ContainetTxT.FindChildren<Text>();
									for(int r=0;r<=innerText.Count-2;r++)
									{
										containerPreference=innerText[r].Caption;
										if(containerPreference=="Enable Remote Fine-tuning")
										{
											containerPreference="Is RFT Enabled";
											FSW_UserPreference_List.Add(containerPreference);
										}
										else if(containerPreference=="Default Patient Setting")
										{
											containerPreference="RFT Default Patient Setting";
											FSW_UserPreference_List.Add(containerPreference);
										}
										else
										{
											FSW_UserPreference_List.Add(containerPreference);
										}
										
									}
									
								}
								else if(k==3)
								{
									Ranorex.Container ContainetTxT="/form[@classname='Window' and @orientation='None' and @processname='SmartFit']/container[@automationid='PART_ExtendedScrollViewer']/container["+k+"]";
									IList<Text> innerText=ContainetTxT.FindChildren<Text>();
									for(int r=0;r<=innerText.Count-2;r++)
									{
										containerPreference=innerText[r].Caption;
										if(containerPreference=="Enable Remote Hearing Aid Update")
										{
											containerPreference="Is RFU Enabled";
											FSW_UserPreference_List.Add(containerPreference);
										}
										else if(containerPreference=="Enable Remote Hearing Aid Update")
										{
											containerPreference="Is RFU Enabled";
											FSW_UserPreference_List.Add(containerPreference);
										}
										else if(containerPreference=="Default Patient Setting")
										{
											containerPreference="RFU Default Patient Setting";
											FSW_UserPreference_List.Add(containerPreference);
										}
										else
										{
											FSW_UserPreference_List.Add(containerPreference);
										}
										
									}
									
								}
								else
								{
									Ranorex.Container ContainetTxT="/form[@classname='Window' and @orientation='None' and @processname='SmartFit']/container[@automationid='PART_ExtendedScrollViewer']/container["+k+"]";
									IList<Text> innerText=ContainetTxT.FindChildren<Text>();
									for(int r=0;r<=innerText.Count-2;r++)
									{
										containerPreference=innerText[r].Caption;
										FSW_UserPreference_List.Add(containerPreference);
									}
								}
							}
							
						}
					}
					if(C_List[i].ControlType=="Text")
					{
						if(C_List[i].Name!="Default Language:")
						{

							
							string name=C_List[i].Name.Replace(":",string.Empty).Replace("*",string.Empty);
							if(name=="Fine-tuning sessions launch to")
							{
								FSW_UserPreference_List.Add("Launch screen for Return Visits");
							}
							else if(name=="AutoRelate on Save")
							{
								FSW_UserPreference_List.Add("Prompt for AutoRelate on Save");
							}
							else if(name=="When navigating to Fitting screen")
							{
								FSW_UserPreference_List.Add("When Navigating to Fit screen");
							}
							else if(name=="Default Gain Level %")
							{
								FSW_UserPreference_List.Add("Gain Percentage of Target");
							}
							else if(name=="Mute instrument when Fitting")
							{
								FSW_UserPreference_List.Add("Mute HI when Fitting");
							}
							else
							{
								FSW_UserPreference_List.Add(name);
							}
						}
						
					}
				}

				
			}
			else if(buildname.Contains("solusmax") || buildname.Contains("audigy"))
			{
				Ranorex.Container Container_UserPrefernces="/form[@classname='Window' and @orientation='None' and @processname='SmartFit' or @processname='SolusMax']/container[@automationid='PART_ExtendedScrollViewer']";
				IList<Container> txtContainer=Container_UserPrefernces.FindChildren<Container>();
				IList<UIAutomation> C_List=Container_UserPrefernces.FindChildren<UIAutomation>();
				for(int i=0;i<=C_List.Count-1;i++)
				{
					if(C_List[i].ControlType=="Group")
					{
						if(C_List[i].Name !="Default Environmental Programs:")
						{
							string fswPreference=C_List[i].Name;
							if(fswPreference=="Remote Fine-tuning")
							{
								fswPreference="Is RFT Visible in Preference Window";
								FSW_UserPreference_List.Add(fswPreference);
							}
							else if(fswPreference=="Remote Hearing Aid Update")
							{
								fswPreference="Is RFU Visible in Preference Window";
								FSW_UserPreference_List.Add(fswPreference);
							}
							else
							{
								FSW_UserPreference_List.Add(fswPreference);
							}

						}
						for(int j=0;j<=txtContainer.Count-1;j++)
						{
							if(C_List[i].Name==txtContainer[j].Caption)
							{
								string containerPreference="";
								int k=j+1;
								if(k==1)
								{
									Ranorex.Container ContainetTxT="/form[@classname='Window' and @orientation='None' and @processname='SmartFit' or @processname='SolusMax']/container[@automationid='PART_ExtendedScrollViewer']/container["+k+"]";
									IList<Text> innerText=ContainetTxT.FindChildren<Text>();
									
									for(int r=0;r<=innerText.Count-2;r++)
									{
										containerPreference=innerText[r].Caption;
										if(containerPreference=="1." || containerPreference=="2." || containerPreference=="3." || containerPreference=="4.")
										{
											containerPreference=containerPreference.Replace(".",string.Empty);
											containerPreference="Default Program " +containerPreference;
											FSW_UserPreference_List.Add(containerPreference);
										}
										else
										{
											FSW_UserPreference_List.Add(containerPreference);
										}
										
									}
									
									
								}
								else if(k==2)
								{
									Ranorex.Container ContainetTxT="/form[@classname='Window' and @orientation='None' and @processname='SmartFit' or @processname='SolusMax']/container[@automationid='PART_ExtendedScrollViewer']/container["+k+"]";
									IList<Text> innerText=ContainetTxT.FindChildren<Text>();
									for(int r=0;r<=innerText.Count-2;r++)
									{
										containerPreference=innerText[r].Caption;
										if(containerPreference=="Enable Remote Fine-tuning")
										{
											containerPreference="Is RFT Enabled";
											FSW_UserPreference_List.Add(containerPreference);
										}
										else if(containerPreference=="Default Patient Setting")
										{
											containerPreference="RFT Default Patient Setting";
											FSW_UserPreference_List.Add(containerPreference);
										}
										else
										{
											FSW_UserPreference_List.Add(containerPreference);
										}
										
									}
									
								}

								else
								{
									Ranorex.Container ContainetTxT="/form[@classname='Window' and @orientation='None' and @processname='SmartFit' or @processname='SolusMax']/container[@automationid='PART_ExtendedScrollViewer']/container["+k+"]";
									IList<Text> innerText=ContainetTxT.FindChildren<Text>();
									for(int r=0;r<=innerText.Count-2;r++)
									{
										containerPreference=innerText[r].Caption;
										FSW_UserPreference_List.Add(containerPreference);
									}
								}
							}
							
						}
					}
					if(C_List[i].ControlType=="Text")
					{
						if(C_List[i].Name!="Default Language:")
						{

							
							string name=C_List[i].Name.Replace(":",string.Empty).Replace("*",string.Empty);
							if(name=="Fine-tuning sessions launch to")
							{
								FSW_UserPreference_List.Add("Launch screen for Return Visits");
							}
							else if(name=="AutoRelate on Save")
							{
								FSW_UserPreference_List.Add("Prompt for AutoRelate on Save");
							}
							else if(name=="When navigating to Fitting screen")
							{
								FSW_UserPreference_List.Add("When Navigating to Fit screen");
							}
							else if(name=="Default Gain Level %")
							{
								FSW_UserPreference_List.Add("Gain Percentage of Target");
							}
							else if(name=="Mute instrument when Fitting")
							{
								FSW_UserPreference_List.Add("Mute HI when Fitting");
							}
							else
							{
								FSW_UserPreference_List.Add(name);
							}
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
		public static void Fetching_UserPreference_DefaultValues(string buildname)
		{
			FSW_UserPreference_Default_List.Clear();
			int _comboboxCount=0;
			int _buttonCount=0;
			int _ContainerCount=0;
			buildname=buildname.Replace(" ",string.Empty);
			buildname=buildname.ToLower();
			if(buildname.Contains("smartfit"))
			{
				Ranorex.Container Container_UserPrefernces="/form[@classname='Window' and @orientation='None' and @processname='SmartFit' or @processname='SolusMax']/container[@automationid='PART_ExtendedScrollViewer']";
				IList<UIAutomation> C_List=Container_UserPrefernces.FindChildren<UIAutomation>();
				for(int i=0;i<=C_List.Count-4;i++)
				{
					if(C_List[i].ControlType=="ComboBox")
					{
						int ComboboxNumber=_comboboxCount+1;
						if(ComboboxNumber==1)
						{
							Ranorex.ComboBox seletedcombobox="/form[@classname='Window' and @orientation='None' and @processname='SmartFit' or @processname='SolusMax']/container/combobox["+ComboboxNumber+"]";
							_comboboxCount++;
						}
						else
						{
							Ranorex.ComboBox seletedcombobox="/form[@classname='Window' and @orientation='None' and @processname='SmartFit' or @processname='SolusMax']/container/combobox["+ComboboxNumber+"]";
							if(seletedcombobox.Text=="Fitting Screen:  Gain Adjustment")
							{
								FSW_UserPreference_Default_List.Add("Fit Screen:  Gain Adjustment");
							}
							else if(seletedcombobox.Text=="Enter in Simulation Mode")
							{
								FSW_UserPreference_Default_List.Add("Connect Manually");
							}
							else
							{
								FSW_UserPreference_Default_List.Add(seletedcombobox.Text);
							}
							_comboboxCount++;
						}
					}
					
					if(C_List[i].ControlType=="Group")
					{
						int ContainerNumber=_ContainerCount+1;
						Ranorex.Container ContainetTxT="/form[@classname='Window' and @orientation='None' and @processname='SmartFit' or @processname='SolusMax']/container[@automationid='PART_ExtendedScrollViewer']/container["+ContainerNumber+"]";
						IList<ComboBox> Container_Combobox=ContainetTxT.FindChildren<ComboBox>();
						IList<Button> Container_Button=ContainetTxT.FindChildren<Button>();
						if(C_List[i].Name !="Default Environmental Programs:")
						{
							FSW_UserPreference_Default_List.Add("Yes");
						}
						for(int j=0;j<=Container_Combobox.Count-1;j++)
						{
							
							int j1=j+1;
							Ranorex.ComboBox seletedcombobox="/form[@classname='Window' and @orientation='None' and @processname='SmartFit' or @processname='SolusMax']/container/container["+ContainerNumber+"]/combobox["+j1+"]";
							FSW_UserPreference_Default_List.Add(seletedcombobox.Text);
							
						}
						for(int k=0;k<=Container_Button.Count-1;k++)
						{
							int k1=k+1;
							if(k1==1)
							{
								Ranorex.Button selectedButton="/form[@classname='Window' and @orientation='None' and @processname='SmartFit' or @processname='SolusMax']/container/container["+ContainerNumber+"]/button["+k1+"]";
								//	FSW_UserPreference_Default_List.Add(selectedButton.Text);
								if(selectedButton.Pressed==true)
								{
									FSW_UserPreference_Default_List.Add("Yes");
								}
								else
								{
									FSW_UserPreference_Default_List.Add("No");
								}
							}
							if(k1==2)
							{
								Ranorex.Button selectedButton="/form[@classname='Window' and @orientation='None' and @processname='SmartFit' or @processname='SolusMax']/container/container["+ContainerNumber+"]/button["+k1+"]";
								
								if(selectedButton.Pressed==true)
								{
									FSW_UserPreference_Default_List.Add("Yes");
								}
								else
								{
									FSW_UserPreference_Default_List.Add("No");
								}
							}
						}
						
						
						
						_ContainerCount++;
						
					}
					if(C_List[i].ControlType=="Button")
					{
						int ButtonNumber=_buttonCount+1;
						Ranorex.Button seletedButton="/form[@classname='Window' and @orientation='None' and @processname='SmartFit' or @processname='SolusMax']/container/button["+ButtonNumber+"]";


						if(seletedButton.Pressed==true)
						{
							FSW_UserPreference_Default_List.Add("Yes");
						}
						else
						{
							FSW_UserPreference_Default_List.Add("No");
						}
						_buttonCount++;
					}
					
				}
			}
			else if(buildname.Contains("solusmax") || buildname.Contains("audigy"))
			{
				Ranorex.Container Container_UserPrefernces="/form[@classname='Window' and @orientation='None' and @processname='SmartFit' or @processname='SolusMax']/container[@automationid='PART_ExtendedScrollViewer']";
				IList<UIAutomation> C_List=Container_UserPrefernces.FindChildren<UIAutomation>();
				for(int i=0;i<=C_List.Count-4;i++)
				{
					if(C_List[i].ControlType=="ComboBox")
					{
						int ComboboxNumber=_comboboxCount+1;
						if(ComboboxNumber==1)
						{
							Ranorex.ComboBox seletedcombobox="/form[@classname='Window' and @orientation='None' and @processname='SmartFit' or @processname='SolusMax']/container/combobox["+ComboboxNumber+"]";
							_comboboxCount++;
						}
						else
						{
							Ranorex.ComboBox seletedcombobox="/form[@classname='Window' and @orientation='None' and @processname='SmartFit' or @processname='SolusMax']/container/combobox["+ComboboxNumber+"]";
							if(seletedcombobox.Text=="Fitting Screen:  Gain Adjustment")
							{
								FSW_UserPreference_Default_List.Add("Fit Screen:  Gain Adjustment");
							}
							else if(seletedcombobox.Text=="Enter in Simulation Mode")
							{
								FSW_UserPreference_Default_List.Add("Connect Manually");
							}
							else
							{
								FSW_UserPreference_Default_List.Add(seletedcombobox.Text);
							}
							_comboboxCount++;
						}
					}
					
					if(C_List[i].ControlType=="Group")
					{
						int ContainerNumber=_ContainerCount+1;
						Ranorex.Container ContainetTxT="/form[@classname='Window' and @orientation='None' and @processname='SmartFit' or @processname='SolusMax']/container[@automationid='PART_ExtendedScrollViewer']/container["+ContainerNumber+"]";
						IList<ComboBox> Container_Combobox=ContainetTxT.FindChildren<ComboBox>();
						IList<Button> Container_Button=ContainetTxT.FindChildren<Button>();
						if(C_List[i].Name !="Default Environmental Programs:")
						{
							FSW_UserPreference_Default_List.Add("Yes");
						}
						for(int j=0;j<=Container_Combobox.Count-1;j++)
						{
							
							int j1=j+1;
							Ranorex.ComboBox seletedcombobox="/form[@classname='Window' and @orientation='None' and @processname='SmartFit' or @processname='SolusMax']/container/container["+ContainerNumber+"]/combobox["+j1+"]";
							FSW_UserPreference_Default_List.Add(seletedcombobox.Text);
							
						}
						for(int k=0;k<=Container_Button.Count-1;k++)
						{
							int k1=k+1;
							if(k1==1)
							{
								Ranorex.Button selectedButton="/form[@classname='Window' and @orientation='None' and @processname='SmartFit' or @processname='SolusMax']/container/container["+ContainerNumber+"]/button["+k1+"]";
								//	FSW_UserPreference_Default_List.Add(selectedButton.Text);
								if(selectedButton.Pressed==true)
								{
									FSW_UserPreference_Default_List.Add("Yes");
								}
								else
								{
									FSW_UserPreference_Default_List.Add("No");
								}
							}
							if(k1==2)
							{
								Ranorex.Button selectedButton="/form[@classname='Window' and @orientation='None' and @processname='SmartFit' or @processname='SolusMax']/container/container["+ContainerNumber+"]/button["+k1+"]";
								
								if(selectedButton.Pressed==true)
								{
									FSW_UserPreference_Default_List.Add("Yes");
								}
								else
								{
									FSW_UserPreference_Default_List.Add("No");
								}
							}
						}
						
						
						
						_ContainerCount++;
						
					}
					if(C_List[i].ControlType=="Button")
					{
						int ButtonNumber=_buttonCount+1;
						Ranorex.Button seletedButton="/form[@classname='Window' and @orientation='None' and @processname='SmartFit' or @processname='SolusMax']/container/button["+ButtonNumber+"]";


						if(seletedButton.Pressed==true)
						{
							FSW_UserPreference_Default_List.Add("Yes");
						}
						else
						{
							FSW_UserPreference_Default_List.Add("No");
						}
						_buttonCount++;
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
		public static void To_Find_Market_Preference_is_present_in_excel_or_Not()
		{
			Excel_UserPreference_List.Clear();
			string excelFinalPath = @"D:\TFS\FSW\TestSuites\Market_Preferences\Mkt Preferences_Smart Fit_1.6_Release.3 (inc. Aventa).xlsx";
			Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
			XLS.Workbook workBook = application.Workbooks.Open(excelFinalPath);
			Excel.Worksheet worksheet = (Excel.Worksheet)workBook.Worksheets[3];

			for(int j=3;j<=30;j++)
			{
				object cellValue = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[3, j]).Value;
				string k1=cellValue.ToString();
				k1=k1.Replace("[AllowRFT]",string.Empty).Replace("[GloballyON]",string.Empty).Replace("[DefaultSetting]",string.Empty).Replace("\n",string.Empty);
				k1=k1.ToLower();
				Excel_UserPreference_List.Add(k1);
			}
			
			Delay.Seconds(2);
			for(int i=0;i<=FSW_UserPreference_List.Count-1;i++)
			{
				int count=0;
				for(int k=0;k<=Excel_UserPreference_List.Count-1;k++)
				{
					
					string j=FSW_UserPreference_List[i];
					j=j.ToLower();
					string j2=Excel_UserPreference_List[k];
					j2=j2.Replace(" ",string.Empty);
					j=j.Replace(" ",string.Empty);
					if(j2.Contains(j))
					{
						
						count =2;
					}
				}
				if(count ==2)
				{
					
				}
				else if(count == 0)
				{
					Report.Failure("The Market Feature is not available in Excel :"+FSW_UserPreference_List[i]);
				}
			}
			
			
		}
		/// <summary>
		/// This is a placeholder text. Please describe the purpose of the
		/// user code method here. The method is published to the user code library
		/// within a user code collection.
		/// </summary>
		[UserCodeMethod]
		public static void Excel_UsePreference_Values(string _marketname,string buildname)
		{
			buildname=buildname.Replace(" ",string.Empty);
			buildname=buildname.ToLower();
			string excelFinalPath="";
			int cellindex=0;
			string Excel_Preference="";
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
			for(int i=0;i<=FSW_UserPreference_List.Count-1;i++)
			{
				for(int j=3;j<=30;j++)
				{
					object cellValue = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[3, j]).Value;
					string k1=cellValue.ToString();
					k1=k1.Replace("[AllowRFT]",string.Empty).Replace("[GloballyON]",string.Empty).Replace("[DefaultSetting]",string.Empty).Replace("\n",string.Empty);
					k1=k1.ToLower();
					string fswprefernce=FSW_UserPreference_List[i];
					fswprefernce=fswprefernce.ToLower();
					k1=k1.Replace(" ",string.Empty);
					fswprefernce=fswprefernce.Replace(" ",string.Empty);
					
					if(k1.Contains(fswprefernce))
					{
						cellindex=j;
						string	k2="Gain Percentage of Target";
						k2=k2.ToLower();
						k2=k2.Replace(" ",string.Empty);
						if(k1==k2)
						{
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
									
									double gain=Convert.ToDouble(cellValue3);
									gain=gain*100;
									Excel_Preference=gain.ToString();
									
								}

							}
							if(Excel_Preference.ToLower()==FSW_UserPreference_Default_List[i].ToLower())
							{
								Report.Success("Market Prefernce "+FSW_UserPreference_List[i] +" Verified succesfully");
							}
							else
							{
								Report.Failure("Market Preference in Excel :"+Excel_Preference+" is different from FSW :"+FSW_UserPreference_Default_List[i]+" of Preference :"+FSW_UserPreference_List[i]);
							}
						}
						else
						{
							
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
									break;
								}

							}
							Excel_Preference=Excel_Preference.Replace(" ",string.Empty).Replace("-",string.Empty).Replace("seconds","sec");
							string fswPreferences=FSW_UserPreference_Default_List[i].Replace(" ",string.Empty).Replace("-",string.Empty);
							Excel_Preference=Excel_Preference.ToLower();
							fswPreferences=fswPreferences.ToLower();
							Excel_Preference=Excel_Preference.Replace("acousticphone","acoustictelephone");
							fswPreferences=fswPreferences.Replace("acousticphone","acoustictelephone");
							if(Excel_Preference==fswPreferences)
							{
								Report.Success("User Prefernce of :"+FSW_UserPreference_List[i]+" is :"+FSW_UserPreference_List[i] +" Verified succesfully");
							}

							else
							{
								Report.Failure("User Preference in Excel of:"+Excel_Preference+" -is different from FSW :"+FSW_UserPreference_Default_List[i]+" -of Preference"+FSW_UserPreference_List[i]);
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
