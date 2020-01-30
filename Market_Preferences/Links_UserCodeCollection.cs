/*
 * Created by Ranorex
 * User: i-ray
 * Date: 18-11-2019
 * Time: 23:21
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
using System.Collections;
using System.Diagnostics;
using System.IO;
using System.Web;
using Ranorex;
using System.Web;
using Ranorex.Core;
//using OpenQA.Selenium
using Ranorex.Core.Testing;
using Excel = Microsoft.Office.Interop.Excel;
using XLS=Microsoft.Office.Interop.Excel;
using System.IO;
//using WinForms = System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
namespace Market_Preferences
{
	/// <summary>
	/// Creates a Ranorex user code collection. A collection is used to publish user code methods to the user code library.
	/// </summary>
	[UserCodeCollection]
	public class Links_UserCodeCollection
	{
		// You can use the "Insert New User Code Method" functionality from the context menu,
		// to add a new method with the attribute [UserCodeMethod].
		public static	Market_Preferences.Market_PreferencesRepository repo=Market_Preferences.Market_PreferencesRepository.Instance;
		/// <summary>
		/// This is a placeholder text. Please describe the purpose of the
		/// user code method here. The method is published to the user code library
		/// within a user code collection.
		/// </summary>
		[UserCodeMethod]
		public static void E_Commerces_Link(string _marketname)
		{
			Kill_All_Open_Browsers();
			string FSW_PageURL="";
			string Excel_PageURL="";
			//	int _Row_index=0;
			Ranorex.MenuItem HelpButtonclick="/form[@title~'Smart Fit' or @title~'Solus Max']/contextmenu/menuitem[@automationid='MenuAutomationIds.HelpAction']";
			string Excel_Preference="";
			string excelFinalPath = @"D:\TFS\FSW\TestSuites\Market_Preferences\Mkt Preferences_Smart Fit_1.6_Release.3 (inc. Aventa).xlsx";
			Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
			XLS.Workbook workBook = application.Workbooks.Open(excelFinalPath);
			Excel.Worksheet worksheet = (Excel.Worksheet)workBook.Worksheets[3];
			for(int i=2;i<=70;i++)
			{
				object cellValue = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[3, i]).Value;
				if(cellValue !=null)
				{
					string _Excel_E_Commerces_Value=cellValue.ToString();
					if(_Excel_E_Commerces_Value=="e-Commerce Link")
					{
						for(int j=2;j<=50;j++)
						{
							string Marketname="";
							object cellValue2 = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[j, 1]).Value;
							if(cellValue2!=null)
							{
								Marketname=cellValue2.ToString();
							}
							if(_marketname==Marketname)
							{
								object cellValue3 = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[j,i]).Value;
								Excel_Preference=cellValue3.ToString();
								Mouse.Click(HelpButtonclick);
								if(Excel_Preference.StartsWith("Yes"))
								{
									int p=i+1;
									object cellValue4 = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[j,p]).Value;
									Excel_PageURL=cellValue4.ToString();
									Ranorex.ContextMenu HelpMenuItems="/contextmenu[@processname='SmartFit' and @win32ownerwindowlevel='1']";
									IList<MenuItem> All_MenuItems=HelpMenuItems.FindChildren<MenuItem>();
									int count=0;
									for(int k=0;k<=All_MenuItems.Count-1;k++)
									{
										int k1=k+1;
										Ranorex.MenuItem HelpMenuText="/contextmenu[@processname='SmartFit' and @win32ownerwindowlevel='1']/menuitem["+k1+"]";
										string txt=HelpMenuText.Text;
										if(txt !=null)
										{
											if(txt=="ReSound eCommerce")
											{
												if(HelpMenuText.Visible==true)
												{
													Mouse.Click(HelpMenuText);
													Delay.Seconds(6);
													//		IList<Ranorex.WebDocument> AllDoms = Host.Local.Find<Ranorex.WebDocument>("/dom");
													IList<Ranorex.WebDocument> AllDoms=Host.Local.FindChildren<Ranorex.WebDocument>();
													foreach (WebDocument myDom in AllDoms)
													{
														if(	myDom.Browser.Title.Contains("ReSound") || myDom.Browser.Title.Contains("myGN") || myDom.Domain.Contains("resound"))
														{
															FSW_PageURL=myDom.PageUrl;
															myDom.Close();
															break;
														}
														
													}
													FSW_PageURL=FSW_PageURL.Replace("/",string.Empty).Replace("https:",string.Empty).Replace("http:",string.Empty).Replace("www.",string.Empty);
													Excel_PageURL=Excel_PageURL.Replace("/",string.Empty).Replace("https:",string.Empty).Replace("http:",string.Empty).Replace("www.",string.Empty);
													//		FSW_PageURL=FSW_PageURL.TrimEnd(FSW_PageURL[FSW_PageURL.Length-1]);
													if(FSW_PageURL==Excel_PageURL)
													{
														Report.Success("Validation of ReSound eCommerce link :"+FSW_PageURL+" is verified Successfully ");
													}
													else
													{
														Report.Failure("ReSound eCommerce link of Excel file :"+Excel_PageURL+" is different from FSW Redirection Link :"+FSW_PageURL);
													}
												}
//												else
//												{
//													Report.Failure("The Button :"+txt +" is not Visible in Help menu");
//												}
												count =2;
												break;
											}
											
										}
										
										
									}
									
									if(count==2)
									{
										
									}
									else if(count == 0)
									{
										Report.Failure("ReSound eCommerce is not Present in Help menu");
									}
								}
								
							}
						}
					}
				}
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
		public static void E_Commerces_Link2(string _marketname)
		{
			Kill_All_Open_Browsers();
			string FSW_PageURL="";
			string Excel_PageURL="";
			string browser_URL="";
			//	int _Row_index=0;
			Ranorex.MenuItem HelpButtonclick="/form[@title~'Smart Fit' or @title~'Solus Max' or @name='Smart Launcher']/contextmenu/menuitem[@automationid='MenuAutomationIds.HelpAction']";
			//    "/form[@name='Smart Launcher' and @classname='Window']/contextmenu/menuitem[@automationid='MenuAutomationIds.HelpAction']"
			string Excel_Preference="";
			string excelFinalPath = @"D:\TFS\FSW\TestSuites\Market_Preferences\Mkt Preferences_Smart Fit.xlsx";
			Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
			XLS.Workbook workBook = application.Workbooks.Open(excelFinalPath);
			Excel.Worksheet worksheet = (Excel.Worksheet)workBook.Worksheets[3];
			for(int i=2;i<=70;i++)
			{
				object cellValue = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[3, i]).Value;
				if(cellValue !=null)
				{
					string _Excel_E_Commerces_Value=cellValue.ToString();
					if(_Excel_E_Commerces_Value=="e-Commerce Link")
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
							if(_marketname==Marketname)
							{
								object cellValue3 = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[j,i]).Value;
								Excel_Preference=cellValue3.ToString();
								//		Mouse.Click(HelpButtonclick);
								if(Excel_Preference.StartsWith("Yes"))
								{
									int p=i+1;
									object cellValue4 = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[j,p]).Value;
									Excel_PageURL=cellValue4.ToString();
									System.Diagnostics.Process.Start("iexplore", Excel_PageURL);
									Delay.Seconds(8);
									IList<Ranorex.WebDocument> AllDoms2=Host.Local.FindChildren<Ranorex.WebDocument>();
									foreach (WebDocument myDom in AllDoms2)
									{
										if(	myDom.Browser.Title.Contains("ReSound") || myDom.Browser.Title.Contains("myGN") || myDom.Domain.Contains("resound"))
										{
											browser_URL=myDom.PageUrl;
											myDom.Close();
											break;
										}
										
									}
									
									//	Mouse.Click(HelpButtonclick);
									Delay.Seconds(5);
									Kill_All_Open_Browsers();
									Mouse.Click(HelpButtonclick);
									Ranorex.ContextMenu HelpMenuItems="/contextmenu[@processname='SmartFit' and @win32ownerwindowlevel='1']";
									IList<MenuItem> All_MenuItems=HelpMenuItems.FindChildren<MenuItem>();
									int count=0;
									for(int k=0;k<=All_MenuItems.Count-1;k++)
									{
										int k1=k+1;
										Ranorex.MenuItem HelpMenuText="/contextmenu[@processname='SmartFit' and @win32ownerwindowlevel='1']/menuitem["+k1+"]";
										string txt=HelpMenuText.Text;
										if(txt !=null)
										{
											if(txt=="ReSound eCommerce")
											{
												if(HelpMenuText.Visible==true)
												{
													Mouse.Click(HelpMenuText);
													Delay.Seconds(6);
													//		IList<Ranorex.WebDocument> AllDoms = Host.Local.Find<Ranorex.WebDocument>("/dom");
													IList<Ranorex.WebDocument> AllDoms=Host.Local.FindChildren<Ranorex.WebDocument>();
													foreach (WebDocument myDom in AllDoms)
													{
														if(	myDom.Browser.Title.Contains("ReSound") || myDom.Browser.Title.Contains("myGN") || myDom.Domain.Contains("resound"))
														{
															FSW_PageURL=myDom.PageUrl;
															myDom.Close();
															break;
														}
														
													}
													if(FSW_PageURL==browser_URL)
													{
														Report.Success("Validation of ReSound eCommerce link :"+FSW_PageURL+" is verified Successfully ");
													}
//													else
//													{
//														Report.Failure("ReSound eCommerce link of Excel file :"+browser_URL+" is different from FSW Redirection Link :"+FSW_PageURL);
//													}
												}
												count =2;
												break;
											}
										}
									}
									if(count==2)
									{
										
									}
									else if(count == 0)
									{
										Report.Failure("ReSound eCommerce is not Present in Help menu");
									}
								}
							}
						}
					}
				}
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
		public static void Resound_WebSite(string _marketname)
		{
			Kill_All_Open_Browsers();
			string FSW_PageURL="";
			Ranorex.MenuItem HelpButtonclick="/form[@title~'Smart Fit' or @title~'Solus Max']/contextmenu/menuitem[@automationid='MenuAutomationIds.HelpAction']";
			string Excel_Preference="";
			string excelFinalPath = @"D:\TFS\FSW\TestSuites\Market_Preferences\Mkt Preferences_Smart Fit.xlsx";
			Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
			XLS.Workbook workBook = application.Workbooks.Open(excelFinalPath);
			Excel.Worksheet worksheet = (Excel.Worksheet)workBook.Worksheets[3];
			for(int i=2;i<=70;i++)
			{
				object cellValue = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[3, i]).Value;
				if(cellValue !=null)
				{
					string _Excel_Website_Value=cellValue.ToString();
					if(_Excel_Website_Value=="Website")
					{
						for(int j=2;j<=50;j++)
						{
							string Marketname="";
							object cellValue2 = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[j, 1]).Value;
							if(cellValue2!=null)
							{
								Marketname=cellValue2.ToString();
							}
							if(_marketname==Marketname)
							{
								Mouse.Click(HelpButtonclick);
								object cellValue3 = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[j,i]).Value;
								Excel_Preference=cellValue3.ToString();
								
								Ranorex.ContextMenu HelpMenuItems="/contextmenu[@processname='SmartFit' and @win32ownerwindowlevel='1']";
								IList<MenuItem> All_MenuItems=HelpMenuItems.FindChildren<MenuItem>();
								for(int k=0;k<=All_MenuItems.Count-1;k++)
								{
									int k1=k+1;
									Ranorex.MenuItem HelpMenuText="/contextmenu[@processname='SmartFit' and @win32ownerwindowlevel='1']/menuitem["+k1+"]";
									string txt=HelpMenuText.Text;
									if(txt !=null)
									{
										if(txt=="ReSound Website")
										{
											Mouse.Click(HelpMenuText);
											Delay.Seconds(6);
											IList<Ranorex.WebDocument> AllDoms=Host.Local.FindChildren<Ranorex.WebDocument>();
											foreach (WebDocument myDom in AllDoms)
											{
												if(	myDom.Browser.Title.Contains("ReSound") || myDom.Domain.Contains("gnhearing") || myDom.Domain.Contains("resound"))
												{
													FSW_PageURL=myDom.PageUrl;
													myDom.Close();
													break;
												}
												
											}
											//	FSW_PageURL=FSW_PageURL.TrimEnd(FSW_PageURL[FSW_PageURL.Length-1]);
											string 	FSW_PageURL2=FSW_PageURL;
											FSW_PageURL=FSW_PageURL.Replace("/",string.Empty).Replace("https:",string.Empty).Replace("http:",string.Empty).Replace("en-AU",string.Empty).Replace("de-at",string.Empty).Replace("da",string.Empty).Replace("fr",string.Empty).Replace("fr-fr",string.Empty).Replace("-",string.Empty).Replace("pt-BR ",string.Empty).Replace("nb-no",string.Empty).Replace("es-es",string.Empty);
											Excel_Preference=Excel_Preference.Replace("/",string.Empty).Replace("https:",string.Empty).Replace("http:",string.Empty).Replace("en-AU",string.Empty).Replace("de-at",string.Empty).Replace(".fi",".com").Replace("-",string.Empty).Replace(".fr",".com").Replace("nb-no",string.Empty).Replace(".no",".com").Replace("es-es",string.Empty).Replace(".es",".com");
											if(FSW_PageURL==Excel_Preference)
											{
												Report.Success("Validation of ReSound Website link :"+FSW_PageURL2+" is verified Successfully ");
											}
											else
											{
												Report.Failure("ReSound Website link link of Excel file :"+Excel_Preference+" is different from FSW Redirection Link :"+FSW_PageURL2);
											}
											break;
										}
									}
								}
								
								
							}
						}
					}
				}
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
		public static void Resound_WebSite2(string _marketname)
		{
			Kill_All_Open_Browsers();
			string FSW_PageURL="";
			string browser_URL="";
			Ranorex.MenuItem HelpButtonclick="/form[@title~'Smart Fit' or @title~'Solus Max' or @name='Smart Launcher']/contextmenu/menuitem[@automationid='MenuAutomationIds.HelpAction']";
			string Excel_Preference="";
			string excelFinalPath = @"D:\TFS\FSW\TestSuites\Market_Preferences\Mkt Preferences_Smart Fit.xlsx";
			Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
			XLS.Workbook workBook = application.Workbooks.Open(excelFinalPath);
			Excel.Worksheet worksheet = (Excel.Worksheet)workBook.Worksheets[3];
			for(int i=2;i<=70;i++)
			{
				object cellValue = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[3, i]).Value;
				if(cellValue !=null)
				{
					string _Excel_Website_Value=cellValue.ToString();
					if(_Excel_Website_Value=="Website")
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
							if(_marketname==Marketname)
							{
								object cellValue3 = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[j,i]).Value;
								Excel_Preference=cellValue3.ToString();
								
								System.Diagnostics.Process.Start("iexplore", Excel_Preference);
								Delay.Seconds(8);
								IList<Ranorex.WebDocument> AllDoms2=Host.Local.FindChildren<Ranorex.WebDocument>();
								foreach (WebDocument myDom in AllDoms2)
								{
									//	if(	myDom.Browser.Title.Contains("ReSound") || myDom.Browser.Title.Contains("myGN") || myDom.Domain.Contains("resound"))
									//	{
									browser_URL=myDom.PageUrl;
									myDom.Close();
									break;
									//	}
									
								}
								
								Delay.Seconds(3);
								Kill_All_Open_Browsers();
								Mouse.Click(HelpButtonclick);
								Ranorex.ContextMenu HelpMenuItems="/contextmenu[@processname='SmartFit' and @win32ownerwindowlevel='1']";
								IList<MenuItem> All_MenuItems=HelpMenuItems.FindChildren<MenuItem>();
								for(int k=0;k<=All_MenuItems.Count-1;k++)
								{
									int k1=k+1;
									Ranorex.MenuItem HelpMenuText="/contextmenu[@processname='SmartFit' and @win32ownerwindowlevel='1']/menuitem["+k1+"]";
									string txt=HelpMenuText.Text;
									if(txt !=null)
									{
										if(txt=="ReSound Website")
										{
											Mouse.Click(HelpMenuText);
											Delay.Seconds(8);
											IList<Ranorex.WebDocument> AllDoms=Host.Local.FindChildren<Ranorex.WebDocument>();
											foreach (WebDocument myDom in AllDoms)
											{
												//	if(	myDom.Browser.Title.Contains("ReSound") || myDom.Domain.Contains("gnhearing") || myDom.Domain.Contains("resound"))
												//	{
												FSW_PageURL=myDom.PageUrl;
												myDom.Close();
												break;
												//	}
												
											}
											FSW_PageURL=FSW_PageURL.Replace("https","http");
											browser_URL=browser_URL.Replace("https","http");
											if(FSW_PageURL==browser_URL)
											{
												Report.Success("Validation of ReSound Website link :"+FSW_PageURL+" is verified Successfully ");
											}
											else
											{
												Report.Failure("ReSound Website link link of Excel file :"+browser_URL+" is different from FSW Redirection Link :"+FSW_PageURL);
											}
											break;
										}
									}
								}
							}
						}
					}
				}
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
		public static void SolusMax_WebSite(string _marketname)
		{
			Kill_All_Open_Browsers();
			string FSW_PageURL="";
			string browser_URL="";
			Ranorex.MenuItem HelpButtonclick="/form[@title~'Smart Fit' or @title~'Solus Max' or @name='Smart Launcher']/contextmenu/menuitem[@automationid='MenuAutomationIds.HelpAction']";
			string Excel_Preference="";
			string excelFinalPath = @"D:\TFS\FSW\TestSuites\Market_Preferences\Mkt Preferences_Solus Max.xlsx";
			Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
			XLS.Workbook workBook = application.Workbooks.Open(excelFinalPath);
			Excel.Worksheet worksheet = (Excel.Worksheet)workBook.Worksheets[3];
			for(int i=2;i<=70;i++)
			{
				object cellValue = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[3, i]).Value;
				if(cellValue !=null)
				{
					string _Excel_Website_Value=cellValue.ToString();
					if(_Excel_Website_Value=="Beltone Website URL")
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
//								if(_marketname=="International Business")
//								{
//									_marketname="International";
//								}
							}
							Marketname=Marketname.Replace(" ",string.Empty).Replace("InternationalBusiness","International");
							_marketname=_marketname.Replace(" ",string.Empty).Replace("InternationalBusiness","International");
							if(_marketname==Marketname)
							{
								object cellValue3 = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[j,i]).Value;
								Excel_Preference=cellValue3.ToString();
								
								System.Diagnostics.Process.Start("iexplore", Excel_Preference);
								Delay.Seconds(8);
								IList<Ranorex.WebDocument> AllDoms2=Host.Local.FindChildren<Ranorex.WebDocument>();
								foreach (WebDocument myDom in AllDoms2)
								{
									//	if(	myDom.Browser.Title.Contains("ReSound") || myDom.Browser.Title.Contains("myGN") || myDom.Domain.Contains("resound"))
									//	{
									browser_URL=myDom.PageUrl;
									myDom.Close();
									break;
									//	}
									
								}
								
								Delay.Seconds(3);
								Kill_All_Open_Browsers();
								Mouse.Click(HelpButtonclick);
								Ranorex.ContextMenu HelpMenuItems="/contextmenu[@processname='SmartFit' or @processname='SolusMax' and @win32ownerwindowlevel='1']";
								IList<MenuItem> All_MenuItems=HelpMenuItems.FindChildren<MenuItem>();
								for(int k=0;k<=All_MenuItems.Count-1;k++)
								{
									int k1=k+1;
									Ranorex.MenuItem HelpMenuText="/contextmenu[@processname='SmartFit' or @processname='SolusMax' and @win32ownerwindowlevel='1']/menuitem["+k1+"]";
									string txt=HelpMenuText.Text;
									if(txt !=null)
									{
										if(txt=="Beltone Website")
										{
											Mouse.Click(HelpMenuText);
											Delay.Seconds(8);
											IList<Ranorex.WebDocument> AllDoms=Host.Local.FindChildren<Ranorex.WebDocument>();
											foreach (WebDocument myDom in AllDoms)
											{
												//	if(	myDom.Browser.Title.Contains("ReSound") || myDom.Domain.Contains("gnhearing") || myDom.Domain.Contains("resound"))
												//	{
												FSW_PageURL=myDom.PageUrl;
												myDom.Close();
												break;
												//	}
												
											}
											FSW_PageURL=FSW_PageURL.Replace("https","http");
											browser_URL=browser_URL.Replace("https","http");
											if(FSW_PageURL==browser_URL)
											{
												Report.Success("Validation of Beltone Website link :"+FSW_PageURL+" is verified Successfully ");
											}
											else
											{
												Report.Failure("Beltone Website link link of Excel file :"+browser_URL+" is different from FSW Redirection Link :"+FSW_PageURL);
											}
											break;
										}
									}
								}
							}
						}
					}
				}
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
		public static void Audigy_Website(string _marketname)
		{
			Kill_All_Open_Browsers();
			string FSW_PageURL="";
			string browser_URL="";
			Ranorex.MenuItem HelpButtonclick="/form[@title~'Smart Fit' or @title~'Solus Max' or @title~'Audigy' or @name='Smart Launcher']/contextmenu/menuitem[@automationid='MenuAutomationIds.HelpAction']";
			string Excel_Preference="";
			string excelFinalPath = @"D:\TFS\FSW\TestSuites\Market_Preferences\Mkt Preferences_Audigy.xlsx";
			Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
			XLS.Workbook workBook = application.Workbooks.Open(excelFinalPath);
			Excel.Worksheet worksheet = (Excel.Worksheet)workBook.Worksheets[3];
			for(int i=2;i<=70;i++)
			{
				object cellValue = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[3, i]).Value;
				if(cellValue !=null)
				{
					string _Excel_Website_Value=cellValue.ToString();
					if(_Excel_Website_Value=="Website")
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
								if(_marketname=="Audigy")
								{
									_marketname="United States";
								}

							}
							Marketname=Marketname.Replace(" ",string.Empty).Replace("InternationalBusiness","International");
							_marketname=_marketname.Replace(" ",string.Empty).Replace("InternationalBusiness","International");
							if(_marketname==Marketname)
							{
								object cellValue3 = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[j,i]).Value;
								Excel_Preference=cellValue3.ToString();
								
								System.Diagnostics.Process.Start("iexplore", Excel_Preference);
								Delay.Seconds(8);
								IList<Ranorex.WebDocument> AllDoms2=Host.Local.FindChildren<Ranorex.WebDocument>();
								foreach (WebDocument myDom in AllDoms2)
								{
									
									browser_URL=myDom.PageUrl;
									myDom.Close();
									break;
									
								}
								
								Delay.Seconds(3);
								Kill_All_Open_Browsers();
								Mouse.Click(HelpButtonclick);
								Ranorex.ContextMenu HelpMenuItems="/contextmenu[@processname='SmartFit' or @processname='SolusMax' and @win32ownerwindowlevel='1']";
								IList<MenuItem> All_MenuItems=HelpMenuItems.FindChildren<MenuItem>();
								for(int k=0;k<=All_MenuItems.Count-1;k++)
								{
									int k1=k+1;
									Ranorex.MenuItem HelpMenuText="/contextmenu[@processname='SmartFit' or @processname='SolusMax' and @win32ownerwindowlevel='1']/menuitem["+k1+"]";
									string txt=HelpMenuText.Text;
									if(txt !=null)
									{
										if(txt=="Audigy Website")
										{
											Mouse.Click(HelpMenuText);
											Delay.Seconds(8);
											IList<Ranorex.WebDocument> AllDoms=Host.Local.FindChildren<Ranorex.WebDocument>();
											foreach (WebDocument myDom in AllDoms)
											{
												
												FSW_PageURL=myDom.PageUrl;
												myDom.Close();
												break;
												
												
											}
											FSW_PageURL=FSW_PageURL.Replace("https","http");
											browser_URL=browser_URL.Replace("https","http");
											if(FSW_PageURL==browser_URL)
											{
												Report.Success("Validation of Audigy Website link :"+FSW_PageURL+" is verified Successfully ");
											}
											else
											{
												Report.Failure("Audigy Website link link of Excel file :"+browser_URL+" is different from FSW Redirection Link :"+FSW_PageURL);
											}
											break;
										}
									}
								}
							}
						}
					}
				}
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
		public static void Click_on_Fitting(string buildname)
		{
			buildname=buildname.Replace(" ",string.Empty);
			buildname=buildname.ToLower();
			if(buildname.Contains("audigy"))
			{
				Ranorex.RadioButton Fitting="/form[@title~'Smart Fit' or @title~'Solus Max' or @title~'Audigy']/?/?/?/radiobutton[@automationid='NavigationAutomationIds.MainAutomationIds.Fitting']";
				Mouse.Click(Fitting);
			}
			else
			{
				Ranorex.RadioButton Fitting="/form[@title~'Smart Fit' or @title~'Solus Max' or @title~'Audigy']/list[2]/?/?/radiobutton[@automationid='NavigationAutomationIds.MainAutomationIds.Fitting']";
				Mouse.Click(Fitting);
			}
		}
		/// <summary>
		/// This is a placeholder text. Please describe the purpose of the
		/// user code method here. The method is published to the user code library
		/// within a user code collection.
		/// </summary>
		[UserCodeMethod]
		public static void Gain_Handlers_method(string _marketname,string buildname)
		{
			Ranorex.List GainHadlersList="/form[@classname='Window' and @orientation='None' and @processname='SmartFit' or @processname='SolusMax']/container[@automationid='PART_ExtendedScrollViewer']/list[3]";
			IList<ListItem> txtList=GainHadlersList.FindChildren<ListItem>();
			int gainlist=txtList.Count;
			Ranorex.Text txtGain_Handler="/form[@name~'Smart Fit' or @name~'Solus Max' or @title~'Audigy' and @classname='Window']/container[@automationid='PART_ExtendedScrollViewer']/list[3]/listitem["+gainlist+"]/radiobutton[1]/text";
			//	Ranorex.Text txtGain_Handler="/form[@name='ReSound Smart Fit 1.6' and @classname='Window']/container[@automationid='PART_ExtendedScrollViewer']/list[3]/listitem["+gainlist+"]/radiobutton[1]/text";
			string FSW_Gain_HighestGainHandle=txtGain_Handler.TextValue;
			Ranorex.MenuItem HelpButtonclick="/form[@title~'Smart Fit' or @title~'Solus Max' or @title~'Audigy']/contextmenu/menuitem[@automationid='MenuAutomationIds.HelpAction']";
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
			//string excelFinalPath = @"D:\TFS\FSW\TestSuites\Market_Preferences\Mkt Preferences_Smart Fit.xlsx";
			Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
			XLS.Workbook workBook = application.Workbooks.Open(excelFinalPath);
			Excel.Worksheet worksheet = (Excel.Worksheet)workBook.Worksheets[3];
			for(int i=2;i<=70;i++)
			{
				object cellValue = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[3, i]).Value;
				if(cellValue !=null)
				{
					string _Excel_GainHandle_Value=cellValue.ToString();
					if(buildname.Contains("smartfit") ||buildname.Contains("audigy"))
					{
						if(_Excel_GainHandle_Value=="Enable Expanded Gain Handles")
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
									if(_marketname=="Audigy")
									{
										_marketname="United States";
									}
								}
								Marketname=Marketname.Replace(" ",string.Empty).Replace("InternationalBusiness","International");
								_marketname=_marketname.Replace(" ",string.Empty).Replace("InternationalBusiness","International");
								if(_marketname!=null)
								{
									if(_marketname==Marketname)
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
					else if(buildname.Contains("solusmax"))
					{
						if(_Excel_GainHandle_Value=="Expanded Gain Handles Availability")
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
								Marketname=Marketname.Replace(" ",string.Empty).Replace("InternationalBusiness","International");
								_marketname=_marketname.Replace(" ",string.Empty).Replace("InternationalBusiness","International");
								Marketname=Marketname.Replace(" ",string.Empty);
								_marketname=_marketname.Replace(" ",string.Empty);
								if(_marketname!=null)
								{
									if(_marketname==Marketname)
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
					
					
				}
			}
			
			
			if(Excel_Preference.Contains(FSW_Gain_HighestGainHandle))
			{
				Report.Success("Validation of Gain Handlers is verfied Successfully");
				
			}
			else if(Excel_Preference.Contains("Yes:  Show Expanded Handles"))
			{
				Report.Success("Validation of Gain Handlers is verfied Successfully");
			}
			//	!( x == 3 || x == 4)
			else if(FSW_Gain_HighestGainHandle != "17")
			{
				//  "No:  Hide Option (Standard Handles Only)"
				Excel_Preference=Excel_Preference.Replace(" ",string.Empty);
				if(Excel_Preference.Contains("No:HideOption(StandardHandlesOnly)"))
				{
					Report.Success("Gain Handlers of Excel sheet is No: Hide Option (Standard Handles Only)");
				}
				else
				{
					Report.Failure("Gain Handlers of Excel file :"+Excel_Preference+ " -is different from Gain Handlers in FSW :"+FSW_Gain_HighestGainHandle);
				}
			}
			else if(Excel_Preference.Contains("No:  Hide Option (Standard Handles Only)") || Excel_Preference.Contains("No:   Hide Option (Standard Handles Only)"))
			{
				Report.Warn("Gain Handlers of Excel sheet is No: Hide Option (Standard Handles Only)");
			}
			else
			{
				Report.Failure("Gain Handlers of Excel file :"+Excel_Preference+ " -is different from Gain Handlers in FSW :"+FSW_Gain_HighestGainHandle);
			}
		}
		/// <summary>
		/// This is a placeholder text. Please describe the purpose of the
		/// user code method here. The method is published to the user code library
		/// within a user code collection.
		/// </summary>
		[UserCodeMethod]
		public static void ReSound_eCademy(string _marketname)
		{
			Kill_All_Open_Browsers();
			string FSW_PageURL="";
			Ranorex.MenuItem HelpButtonclick="/form[@title~'Smart Fit' or @title~'Solus Max']/contextmenu/menuitem[@automationid='MenuAutomationIds.HelpAction']";
			string Excel_Preference="";
			string excelFinalPath = @"D:\TFS\FSW\TestSuites\Market_Preferences\Mkt Preferences_Smart Fit.xlsx";
			Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
			XLS.Workbook workBook = application.Workbooks.Open(excelFinalPath);
			Excel.Worksheet worksheet = (Excel.Worksheet)workBook.Worksheets[3];
			for(int i=2;i<=70;i++)
			{
				object cellValue = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[3, i]).Value;
				if(cellValue !=null)
				{
					string _Excel_Website_Value=cellValue.ToString();
					if(_Excel_Website_Value=="eCademy redirect URL")
					{
						for(int j=2;j<=50;j++)
						{
							string Marketname="";
							object cellValue2 = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[j, 1]).Value;
							if(cellValue2!=null)
							{
								Marketname=cellValue2.ToString();
							}
							if(_marketname!=null)
							{
								if(_marketname==Marketname)
								{
									Mouse.Click(HelpButtonclick);
									Delay.Seconds(2);
									object cellValue3 = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[j,i]).Value2.ToString();
									if(cellValue3 !=null )
									{
										Excel_Preference=cellValue3.ToString();
										Ranorex.ContextMenu HelpMenuItems="/contextmenu[@processname='SmartFit' and @win32ownerwindowlevel='1']";
										IList<MenuItem> All_MenuItems=HelpMenuItems.FindChildren<MenuItem>();
										for(int k=0;k<=All_MenuItems.Count-1;k++)
										{
											int k1=k+1;
											Ranorex.MenuItem HelpMenuText="/contextmenu[@processname='SmartFit' and @win32ownerwindowlevel='1']/menuitem["+k1+"]";
											string txt=HelpMenuText.Text;
											if(txt !=null)
											{
												if(txt=="ReSound eCademy")
												{
													Mouse.Click(HelpMenuText);
													Delay.Seconds(6);
													IList<Ranorex.WebDocument> AllDoms=Host.Local.FindChildren<Ranorex.WebDocument>();
													foreach (WebDocument myDom in AllDoms)
													{
														//		if(	myDom.Browser.Title.Contains("eCademy"))
														//		{
														FSW_PageURL=myDom.PageUrl;
														myDom.Close();
														break;
														//	}
														
													}
													string 	FSW_PageURL2=FSW_PageURL;
													FSW_PageURL=FSW_PageURL.Replace("/",string.Empty).Replace("https:",string.Empty).Replace("http:",string.Empty).Replace("-AUe",string.Empty).Replace("comen","come").Replace("pt-BR",string.Empty);
													Excel_Preference=Excel_Preference.Replace("/",string.Empty).Replace("https:",string.Empty).Replace("http:",string.Empty).Replace("pt-BR",string.Empty);
													FSW_PageURL=FSW_PageURL.ToLower();
													FSW_PageURL=FSW_PageURL.Replace("fi-fi",string.Empty);
													Excel_Preference=Excel_Preference.ToLower();
													if(FSW_PageURL==Excel_Preference)
													{
														Report.Success("Validation of ReSound eCademy link :"+FSW_PageURL2+" is verified Successfully ");
													}
													else
													{
														Report.Failure("ReSound eCademy link of Excel file :"+Excel_Preference+" is different from FSW Redirection Link :"+FSW_PageURL2);
													}
													break;
												}
											}
										}
									}
								}
							}
						}
					}
				}
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
		public static void ReSound_eCademy2(string _marketname)
		{
			Kill_All_Open_Browsers();
			string FSW_PageURL="";
			string browser_URL="";
			Ranorex.MenuItem HelpButtonclick="/form[@title~'Smart Fit' or @title~'Solus Max' or @name='Smart Launcher']/contextmenu/menuitem[@automationid='MenuAutomationIds.HelpAction']";
			string Excel_Preference="";
			string excelFinalPath = @"D:\TFS\FSW\TestSuites\Market_Preferences\Mkt Preferences_Smart Fit.xlsx";
			Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
			XLS.Workbook workBook = application.Workbooks.Open(excelFinalPath);
			Excel.Worksheet worksheet = (Excel.Worksheet)workBook.Worksheets[3];
			for(int i=2;i<=70;i++)
			{
				object cellValue = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[3, i]).Value;
				if(cellValue !=null)
				{
					string _Excel_Website_Value=cellValue.ToString();
					if(_Excel_Website_Value=="eCademy redirect URL")
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
								if(_marketname==Marketname)
								{
									object cellValue3 = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[j,i]).Value2.ToString();
									if(cellValue3 !=null )
									{
										Excel_Preference=cellValue3.ToString();
										System.Diagnostics.Process.Start("iexplore", Excel_Preference);
										Delay.Seconds(8);
										IList<Ranorex.WebDocument> AllDoms2=Host.Local.FindChildren<Ranorex.WebDocument>();
										foreach (WebDocument myDom in AllDoms2)
										{
											//	if(	myDom.Browser.Title.Contains("ReSound") || myDom.Browser.Title.Contains("myGN") || myDom.Domain.Contains("resound"))
											//	{
											browser_URL=myDom.PageUrl;
											myDom.Close();
											break;
											//	}
											
										}
										
										Delay.Seconds(3);
										Kill_All_Open_Browsers();
										Mouse.Click(HelpButtonclick);
										Ranorex.ContextMenu HelpMenuItems="/contextmenu[@processname='SmartFit' and @win32ownerwindowlevel='1']";
										IList<MenuItem> All_MenuItems=HelpMenuItems.FindChildren<MenuItem>();
										for(int k=0;k<=All_MenuItems.Count-1;k++)
										{
											int k1=k+1;
											Ranorex.MenuItem HelpMenuText="/contextmenu[@processname='SmartFit' and @win32ownerwindowlevel='1']/menuitem["+k1+"]";
											string txt=HelpMenuText.Text;
											if(txt !=null)
											{
												if(txt=="ReSound eCademy")
												{
													Mouse.Click(HelpMenuText);
													Delay.Seconds(6);
													IList<Ranorex.WebDocument> AllDoms=Host.Local.FindChildren<Ranorex.WebDocument>();
													foreach (WebDocument myDom in AllDoms)
													{
														//		if(	myDom.Browser.Title.Contains("eCademy"))
														//		{
														FSW_PageURL=myDom.PageUrl;
														myDom.Close();
														break;
														//	}
														
													}
													
													if(FSW_PageURL==browser_URL)
													{
														Report.Success("Validation of ReSound eCademy link :"+FSW_PageURL+" is verified Successfully ");
													}
													else
													{
														Report.Failure("ReSound eCademy link of Excel file :"+browser_URL+" is different from FSW Redirection Link :"+FSW_PageURL);
													}
													break;
												}
											}
										}
										
										
									}
								}
							}
						}
					}
				}
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
		public static void UserGuide_and_CableGuide(string _marketname)
		{
			Kill_All_Open_Browsers();
			string FSW_PageURL="";
			string 	FSW_PageURL2="";
			Ranorex.MenuItem HelpButtonclick="/form[@title~'Smart Fit' or @title~'Solus Max']/contextmenu/menuitem[@automationid='MenuAutomationIds.HelpAction']";
			string Excel_Preference="";
			string Cable_Preference="";
			string excelFinalPath = @"D:\TFS\FSW\TestSuites\Market_Preferences\Mkt Preferences_Smart Fit.xlsx";
			Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
			XLS.Workbook workBook = application.Workbooks.Open(excelFinalPath);
			Excel.Worksheet worksheet = (Excel.Worksheet)workBook.Worksheets[3];
			int j1=0;
			for(int i=2;i<=70;i++)
			{
				object cellValue = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[3, i]).Value;
				if(cellValue !=null)
				{
					string _Excel_UserGuide_Value=cellValue.ToString();
					if(_Excel_UserGuide_Value=="User Guide URL")
					{
						for(int j=2;j<=50;j++)
						{
							string Marketname="";
							object cellValue2 = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[j, 1]).Value;
							if(cellValue2!=null)
							{
								Marketname=cellValue2.ToString();
							}
							if(_marketname!=null)
							{
								if(_marketname==Marketname)
								{
									j1=j;
									Mouse.Click(HelpButtonclick);
									Delay.Seconds(2);
									object cellValue3 = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[j,i]).Value2.ToString();
									if(cellValue3 !=null )
									{
										Excel_Preference=cellValue3.ToString();
										Ranorex.ContextMenu HelpMenuItems="/contextmenu[@processname='SmartFit' and @win32ownerwindowlevel='1']";
										IList<MenuItem> All_MenuItems=HelpMenuItems.FindChildren<MenuItem>();
										for(int k=0;k<=All_MenuItems.Count-1;k++)
										{
											int k1=k+1;
											Ranorex.MenuItem HelpMenuText="/contextmenu[@processname='SmartFit' and @win32ownerwindowlevel='1']/menuitem["+k1+"]";
											string txt=HelpMenuText.Text;
											if(txt !=null)
											{
												if(txt=="Cable Guide")
												{
													Mouse.Click(HelpMenuText);
													Delay.Seconds(6);
													IList<Ranorex.WebDocument> AllDoms=Host.Local.FindChildren<Ranorex.WebDocument>();
													foreach (WebDocument myDom in AllDoms)
													{
//														if(	myDom.Browser.Title.Contains("eCademy"))
//														{
														FSW_PageURL=myDom.PageUrl;
														myDom.Close();
														break;
//														}
														
													}
													FSW_PageURL2=FSW_PageURL;
													FSW_PageURL=FSW_PageURL.Replace("/",string.Empty).Replace("https:",string.Empty).Replace("http:",string.Empty).Replace("-AUe",string.Empty).Replace("comen","come").Replace("pt-BR",string.Empty);
													Excel_Preference=Excel_Preference.Replace("/",string.Empty).Replace("https:",string.Empty).Replace("http:",string.Empty);
													if(FSW_PageURL==Excel_Preference)
													{
														Report.Success("Validation of User Guide link :"+FSW_PageURL2+" is verified Successfully ");
													}
													else
													{
														Report.Failure("ReSound User Guide link of Excel file :"+Excel_Preference+" is different from FSW Redirection Link :"+FSW_PageURL2);
													}
													break;
												}
											}
										}
									}
								}
							}
						}
					}
					
				}
			}
			
			for(int k=2;k<=70;k++)
			{
				object cellValue = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[3, k]).Value;
				if(cellValue !=null)
				{
					string _Excel_CableGuide_Value=cellValue.ToString();
					if(_Excel_CableGuide_Value=="Cable Guide URL")
					{
						object cellValue3 = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[j1,k]).Value2.ToString();
						if(cellValue3 !=null )
						{
							Cable_Preference=cellValue3.ToString();
							Cable_Preference=Cable_Preference.Replace("/",string.Empty).Replace("https:",string.Empty).Replace("http:",string.Empty);
							if(FSW_PageURL==Cable_Preference)
							{
								Report.Success("Validation of Cable Guide link :"+FSW_PageURL2+" is verified Successfully ");
							}
							else
							{
								Report.Failure("ReSound Cable Guide link of Excel file :"+Cable_Preference+" is different from FSW Redirection Link :"+FSW_PageURL2);
							}
							break;
						}
					}
				}
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
		public static void UserGuide_and_CableGuide2(string _marketname)
		{
			Kill_All_Open_Browsers();
			string FSW_PageURL="";
			string browser_URL="";
			string browser_URL2="";
			//	string FSW_PageURL="";
			string 	FSW_PageURL2="";
			Ranorex.MenuItem HelpButtonclick="/form[@title~'Smart Fit' or @title~'Solus Max' or @name='Smart Launcher']/contextmenu/menuitem[@automationid='MenuAutomationIds.HelpAction']";
			string Excel_Preference="";
			string Cable_Preference="";
			
			string excelFinalPath = @"D:\TFS\FSW\TestSuites\Market_Preferences\Mkt Preferences_Smart Fit.xlsx";
			Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
			XLS.Workbook workBook = application.Workbooks.Open(excelFinalPath);
			Excel.Worksheet worksheet = (Excel.Worksheet)workBook.Worksheets[3];
			int j1=0;
			for(int i=2;i<=70;i++)
			{
				object cellValue = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[3, i]).Value;
				if(cellValue !=null)
				{
					string _Excel_UserGuide_Value=cellValue.ToString();
					if(_Excel_UserGuide_Value=="User Guide URL")
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
								if(_marketname==Marketname)
								{
									j1=j;
									object cellValue3 = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[j,i]).Value2.ToString();
									if(cellValue3 !=null )
									{
										Excel_Preference=cellValue3.ToString();
										System.Diagnostics.Process.Start("iexplore", Excel_Preference);
										Delay.Seconds(8);
										IList<Ranorex.WebDocument> AllDoms2=Host.Local.FindChildren<Ranorex.WebDocument>();
										foreach (WebDocument myDom in AllDoms2)
										{
//											if(	myDom.Browser.Title.Contains("ReSound") || myDom.Browser.Title.Contains("myGN") || myDom.Domain.Contains("resound"))
//											{
											browser_URL=myDom.PageUrl;
											myDom.Close();
											break;
											//			}
											
										}
										
										Delay.Seconds(3);
										Kill_All_Open_Browsers();
										Mouse.Click(HelpButtonclick);
										
										Ranorex.ContextMenu HelpMenuItems="/contextmenu[@processname='SmartFit' and @win32ownerwindowlevel='1']";
										IList<MenuItem> All_MenuItems=HelpMenuItems.FindChildren<MenuItem>();
										for(int k=0;k<=All_MenuItems.Count-1;k++)
										{
											int k1=k+1;
											Ranorex.MenuItem HelpMenuText="/contextmenu[@processname='SmartFit' and @win32ownerwindowlevel='1']/menuitem["+k1+"]";
											string txt=HelpMenuText.Text;
											if(txt !=null)
											{
												if(txt=="Cable Guide")
												{
													Mouse.Click(HelpMenuText);
													Delay.Seconds(8);
													IList<Ranorex.WebDocument> AllDoms=Host.Local.FindChildren<Ranorex.WebDocument>();
													foreach (WebDocument myDom in AllDoms)
													{
//														if(	myDom.Browser.Title.Contains("eCademy"))
//														{
														FSW_PageURL=myDom.PageUrl;
														myDom.Close();
														break;
//														}
														
													}
													if(FSW_PageURL==browser_URL)
													{
														Report.Success("Validation of User Guide link :"+FSW_PageURL+" is verified Successfully ");
													}
													else
													{
														Report.Failure("ReSound User Guide link of Excel file :"+browser_URL+" is different from FSW Redirection Link :"+FSW_PageURL);
													}
													break;
												}
											}
										}
									}
								}
							}
						}
					}
				}
			}
			
			for(int k=2;k<=70;k++)
			{
				object cellValue = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[3, k]).Value;
				if(cellValue !=null)
				{
					string _Excel_CableGuide_Value=cellValue.ToString();
					if(_Excel_CableGuide_Value=="Cable Guide URL")
					{
						object cellValue3 = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[j1,k]).Value2.ToString();
						if(cellValue3 !=null )
						{
							Cable_Preference=cellValue3.ToString();
							System.Diagnostics.Process.Start("iexplore", Cable_Preference);
							Delay.Seconds(8);
							IList<Ranorex.WebDocument> AllDoms3=Host.Local.FindChildren<Ranorex.WebDocument>();
							foreach (WebDocument myDom in AllDoms3)
							{
//								if(	myDom.Browser.Title.Contains("ReSound") || myDom.Browser.Title.Contains("myGN") || myDom.Domain.Contains("resound"))
//								{
								browser_URL2=myDom.PageUrl;
								myDom.Close();
								break;
								//	}
								
							}
							Delay.Seconds(2);
							if(FSW_PageURL==browser_URL2)
							{
								Report.Success("Validation of Cable Guide link :"+FSW_PageURL+" is verified Successfully ");
							}
							else
							{
								Report.Failure("ReSound Cable Guide link of Excel file :"+browser_URL2+" is different from FSW Redirection Link :"+FSW_PageURL);
							}
							break;
							
						}
					}
				}
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
		public static void Getting_URL()
		{
			IList<Ranorex.WebDocument> AllDoms = Host.Local.Find<Ranorex.WebDocument>("/dom");
			foreach (WebDocument myDom in AllDoms)
			{
				myDom.Close();
			}
		}
		/// <summary>
		/// This is a placeholder text. Please describe the purpose of the
		/// user code method here. The method is published to the user code library
		/// within a user code collection.
		/// </summary>
		[UserCodeMethod]
		public static void Kill_All_Open_Browsers()
		{

			Process[] AllProcesses = Process.GetProcesses();
			foreach (var process in AllProcesses)
			{
				if (process.MainWindowTitle != "")
				{
					string s = process.ProcessName.ToLower();
					if (s == "iexplore" || s == "iexplorer" || s == "chrome" || s == "firefox" || s == "MicrosoftEdge" || s == "MicrosoftEdgeCP")
						process.Kill();
				}
			}
			
			
			Process[] Edge = Process.GetProcessesByName("MicrosoftEdge");
			foreach (Process Item in Edge)
			{
				try
				{
					Item.Kill();
					Item.WaitForExit(3000);
				}
				catch (Exception)
				{
				}
			}
			
			Delay.Seconds(5);
		}
	}
}
