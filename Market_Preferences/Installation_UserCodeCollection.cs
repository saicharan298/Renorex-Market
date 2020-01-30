/*
 * Created by Ranorex
 * User: i-ray
 * Date: 18-11-2019
 * Time: 05:43
 * 
 * To change this template use Tools > Options > Coding > Edit standard headers.
 */
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Threading;
using System.Diagnostics;
using WinForms = System.Windows.Forms;
//using WinForms = System.Windows.Forms;
using XLS = Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
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
	public class Installation_UserCodeCollection
	{
		// You can use the "Insert New User Code Method" functionality from the context menu,
		// to add a new method with the attribute [UserCodeMethod].
		
		/// <summary>
		/// This is a placeholder text. Please describe the purpose of the
		/// user code method here. The method is published to the user code library
		/// within a user code collection.
		/// </summary>
		[UserCodeMethod]
		public static void Install_FSW(string Brandname)
		{
			string Brandname2=Brandname.Replace(" ",string.Empty);
			if(Brandname2.Contains("SmartFit"))
			{
				Host.Local.RunApplication("D:\\TFS\\FSW\\TestSuites\\Market_Preferences\\Builds\\"+Brandname+"\\Setup.exe", "", "", false);
				Delay.Milliseconds(0);
			}
			else if(Brandname2.Contains("SolusMax"))
			{
				Host.Local.RunApplication("D:\\TFS\\FSW\\TestSuites\\Market_Preferences\\Builds\\"+Brandname+"\\Setup.exe", "", "", false);
				Delay.Milliseconds(0);
			}
			else if(Brandname2.Contains("Audigy"))
			{
				Host.Local.RunApplication("D:\\TFS\\FSW\\TestSuites\\Market_Preferences\\Builds\\"+Brandname+"\\Setup.exe", "", "", false);
				Delay.Milliseconds(0);
			}
			
		}
		
		/// <summary>
		/// This is a placeholder text. Please describe the purpose of the
		/// user code method here. The method is published to the user code library
		/// within a user code collection.
		/// </summary>
		[UserCodeMethod]
		public static void If_Application_Is_Not_unInstall(string UninstallBuildname)
		{
			try
			{
				Ranorex.Text App_uninstall="/form[@title~'Smart Fit' or @title~'Solus Max' or @title~'Audigy']/text[@windowtext~'A newer or same version of this application is already installed on this computer. If you wish to install this version, please uninstall the installed version first.']";
				if(App_uninstall.Visible==true)
				{
					Ranorex.Button btnOk="/form[@title~'Smart Fit' or @title~'Solus Max' or @title~'Audigy']/button[@text='OK']";
					Mouse.Click(btnOk);
					Delay.Seconds(5);
					Process p = new Process();
					Delay.Seconds(3);
					p.StartInfo.FileName = @"D:\TFS\FSW\TestSuites\Market_Preferences\Builds\"+UninstallBuildname+"\\Setup.exe";
					//	p.StartInfo.FileName = @"~\Desktop\Market_Preferences\Builds\"+UninstallBuildname+"\\Setup.exe";
					p.StartInfo.Arguments = "/x /v/qn";
					p.Start();
					Delay.Seconds(5);
					
				}
				
			}
			catch
			{
				Ranorex.Button cancelbtn="/form[@title~'Smart Fit' or @title~'Solus Max' or @title~'Audigy']/button[@text='Cancel']";
				Mouse.Click(cancelbtn);
				Delay.Seconds(3);
			}
		}
		[UserCodeMethod]
		public static void EnterDevicename(string devicename)
		{
			Ranorex.Text DeviceselectionTextbox="/form[@name~'Smart Fit' or @name~'Launcher' or @name~'Solus Max' or @name~'Interton Fitting' or @name~'Audigy' or @name~'Costco' or @classname='Window']/text";
			//if the Device name Ends with 'P',it will delete the Last three character of the Device name and Enters in Textbox
			if(devicename.EndsWith("P"))
			{
				devicename=devicename.Substring(0,devicename.Length-3);
				DeviceselectionTextbox.TextValue=devicename;
			}
			else
			{
				DeviceselectionTextbox.TextValue=devicename;
			}
			
		}
		
		[UserCodeMethod]
		public static void Select_Device(string devicename)
		{
			Ranorex.Text DeviceselectionTextbox="/form[@name~'Smart Fit' or @name~'Launcher' or @name~'Solus Max' or @name~'Interton Fitting' or @name~'Audigy' or @name~'Costco' or @classname='Window']/text";
			Ranorex.Container DeviceContainer="/contextmenu[@processname~'SmartFit' or @processname~'SolusMax' or @processname~'Audigy' or @processname~'Interton']/list/container/container";
			IList<ListItem> txtlistitem=DeviceContainer.FindChildren<ListItem>();
			for(int i=0;i<=txtlistitem.Count-1;i++)
			{
				int j=i+1;
				Ranorex.ListItem device="/contextmenu[@processname~'SmartFit' or @processname~'SolusMax' or @processname~'Audigy' or @processname~'Interton']/list/container/container/listitem["+j+"]";
				string k=device.Text;
				if(devicename==k||k==DeviceselectionTextbox.TextValue)
				{
					Mouse.Click(device);
					break;
				}

			}
			
		}
		/// <summary>
		/// This is a placeholder text. Please describe the purpose of the
		/// user code method here. The method is published to the user code library
		/// within a user code collection.
		/// </summary>
		[UserCodeMethod]
		public static void ExtraNext_Button()
		{
			Delay.Seconds(4);
			try
			{
				Ranorex.Button btnNext="/form[@title~'Smart Fit' or @title~'Solus Max']/button[@text='&Next >']";
				if(btnNext.Visible==true)
				{
					Mouse.Click(btnNext);
				}
			}
			catch
			{
				
			}
		}
		
		
		/// <summary>
		/// This is a placeholder text. Please describe the purpose of the
		/// user code method here. The method is published to the user code library
		/// within a user code collection.
		/// </summary>
		[UserCodeMethod]
		public static void Uninstall(string UninstallBuildname)
		{
			Process p = new Process();
			Delay.Seconds(3);
			p.StartInfo.FileName = @"D:\TFS\FSW\TestSuites\Market_Preferences\Builds\"+UninstallBuildname+"\\Setup.exe";
			//	p.StartInfo.FileName = @"~\Desktop\Market_Preferences\Builds\"+UninstallBuildname+"\\Setup.exe";
			p.StartInfo.Arguments = "/x /v/qn";
			p.Start();
			Delay.Seconds(5);
		}
		
		/// <summary>
		/// This is a placeholder text. Please describe the purpose of the
		/// user code method here. The method is published to the user code library
		/// within a user code collection.
		/// </summary>
		[UserCodeMethod]
		public static void Deleting_Market_Pref_file(string buildname)
		{
			buildname=buildname.Replace(" ",string.Empty);
			buildname=buildname.ToLower();
			if(buildname.Contains("smartfit"))
			{
				if(System.IO.Directory.Exists(@"C:\Program Files (x86)\ReSound\SmartFit\Data\Market\"))
				{
					try
					{
						System.IO.Directory.Delete(@"C:\Program Files (x86)\ReSound\SmartFit\Data\Market", true);
					}
					catch
					{
						
					}
				}
				try
				{
				string MarketPref_Path=@"C:\Users\i-ray\AppData\Roaming\ReSound\SmartFit\User.pref";
				File.Delete(MarketPref_Path);
				}
				catch
				{
					
				}
			}
			else if(buildname.Contains("solusmax"))
			{
				if(System.IO.Directory.Exists(@"C:\Program Files (x86)\Beltone\SolusMax\Data\Market\"))
				{
					try
					{
						System.IO.Directory.Delete(@"C:\Program Files (x86)\Beltone\SolusMax\Data\Market", true);
					}
					catch
					{
						
					}
				}
				try{
				string MarketPref_Path=@"C:\Users\i-ray\AppData\Roaming\Beltone\SolusMax\User.pref";
				File.Delete(MarketPref_Path);
				}
				catch
				{
					
				}
				
				
				
			}
			
			
		}
		
		/// <summary>
		/// This is a placeholder text. Please describe the purpose of the
		/// user code method here. The method is published to the user code library
		/// within a user code collection.
		/// </summary>
		[UserCodeMethod]
		public static void kill_All_Open_ExcelSheets()
		{
			foreach (Process clsProcess in Process.GetProcesses())
				if (clsProcess.ProcessName.Equals("EXCEL"))  //Process Excel?
					clsProcess.Kill();
		}
		/// <summary>
		/// This is a placeholder text. Please describe the purpose of the
		/// user code method here. The method is published to the user code library
		/// within a user code collection.
		/// </summary>
		[UserCodeMethod]
		public static void DeleteFiles(string DeleteBuildname)
		{
			DeleteBuildname=DeleteBuildname.Replace(" ",string.Empty);
			if(DeleteBuildname.Contains("SmartFit"))
			{
				String sPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
				String sPath2 = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
				string fileName = @"Beltone\";
				string filename2=@"Temp";
				string path=Path.Combine(sPath,fileName);
				if(System.IO.Directory.Exists(path))

				{
					
					
					System.IO.Directory.Delete(path,true);
				}
				
				
				string path2=Path.Combine(sPath2,filename2);
				if(System.IO.Directory.Exists(path2))
				{
					try
					{

						System.IO.Directory.Delete(path2, true);
						
					}

					catch
					{
						
					}
				}
				
				if(System.IO.Directory.Exists(@"C:\ProgramData\GN\"))
				{
					
					try
					{
						System.IO.Directory.Delete(@"C:\ProgramData\GN", true);
						
					}

					catch
					{
						
					}
				}
				if(System.IO.Directory.Exists(@"C:\ProgramData\FLEXnet\"))
				{
					
					try
					{
						System.IO.Directory.Delete(@"C:\ProgramData\FLEXnet", true);
						
					}

					catch
					{
						
					}
				}
				if(System.IO.Directory.Exists(@"C:\ProgramData\ReSound\Aventa\"))
				{
					
					try
					{
						System.IO.Directory.Delete(@"C:\ProgramData\ReSound\Aventa", true);
						
					}

					catch
					{
						
					}
				}
				
				
				
				if(System.IO.Directory.Exists(@"C:\Program Files (x86)\ReSound\Common\"))
				{
					try
					{
						System.IO.Directory.Delete(@"C:\Program Files (x86)\ReSound\Common", true);
					}
					catch
					{
						
					}
				}
				if(System.IO.Directory.Exists(@"C:\Program Files (x86)\ReSound\Merlin\"))
				{
					try
					{
						System.IO.Directory.Delete(@"C:\Program Files (x86)\ReSound\Merlin", true);
					}
					catch
					{
						
					}
				}
				
				
				
				if(System.IO.Directory.Exists(@"C:\ProgramData\ReSound\Fuse2\"))
				{
					
					System.IO.Directory.Delete(@"C:\ProgramData\ReSound\Fuse2", true);
				}
				if(System.IO.Directory.Exists(@"C:\ProgramData\ReSound\SmartFit\"))
				{
					
					System.IO.Directory.Delete(@"C:\ProgramData\ReSound\SmartFit", true);
				}
				

				if(System.IO.Directory.Exists(@"C:\Program Files (x86)\ReSound\SmartFit\"))
				{
					
					try
					{
						System.IO.Directory.Delete(@"C:\Program Files (x86)\ReSound\SmartFit", true);
						
					}

					catch
					{
						
					}
				}
			}
			else if(DeleteBuildname.Contains("SolusMax"))
			{
				
				String sPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
				String sPath2 = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
				string fileName = @"Beltone\";
				string filename2=@"Temp";
				string path=Path.Combine(sPath,fileName);
				if(System.IO.Directory.Exists(path))

				{
					
					
					System.IO.Directory.Delete(path,true);
				}
				
				
				string path2=Path.Combine(sPath2,filename2);
				if(System.IO.Directory.Exists(path2))
				{
					try
					{

						System.IO.Directory.Delete(path2, true);
						
					}

					catch
					{
						
					}
				}
				if(System.IO.Directory.Exists(@"C:\ProgramData\GN\"))
				{
					
					try
					{
						System.IO.Directory.Delete(@"C:\ProgramData\GN", true);
						
					}

					catch
					{
						
					}
				}
				if(System.IO.Directory.Exists(@"C:\ProgramData\Beltone\"))
				{
					
					try
					{
						System.IO.Directory.Delete(@"C:\ProgramData\Beltone", true);
						
					}

					catch
					{
						
					}
				}
				
				
				
				
				
				if(System.IO.Directory.Exists(@"C:\Program Files (x86)\Beltone\Common\"))
				{
					
					try
					{
						System.IO.Directory.Delete(@"C:\Program Files (x86)\Beltone\Common", true);
						
					}
					
					catch
					{
						
					}
				}
				if(System.IO.Directory.Exists(@"C:\Program Files (x86)\Beltone\SolusMax\"))
				{
					
					try
					{
						System.IO.Directory.Delete(@"C:\Program Files (x86)\Beltone\SolusMax", true);
						
					}
					
					catch
					{
						
					}
				}
				

			}
			else if(DeleteBuildname.Contains("Audigy"))
			{
				String sPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
				String sPath2 = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
				string fileName = @"Audigy\";
				string filename2=@"Temp";
				string path=Path.Combine(sPath,fileName);
				if(System.IO.Directory.Exists(path))

				{
					
					
					System.IO.Directory.Delete(path,true);
				}
				
				
				string path2=Path.Combine(sPath2,filename2);
				if(System.IO.Directory.Exists(path2))
				{
					try
					{

						System.IO.Directory.Delete(path2, true);
						
					}

					catch
					{
						
					}
				}
				if(System.IO.Directory.Exists(@"C:\ProgramData\GN\"))
				{
					
					try
					{
						System.IO.Directory.Delete(@"C:\ProgramData\GN", true);
						
					}

					catch
					{
						
					}
				}
				if(System.IO.Directory.Exists(@"C:\ProgramData\Audigy\"))
				{
					
					try
					{
						System.IO.Directory.Delete(@"C:\ProgramData\Audigy", true);
						
					}

					catch
					{
						
					}
				}
				
				
				
				
				
				if(System.IO.Directory.Exists(@"C:\Program Files (x86)\Audigy\Common\"))
				{
					
					try
					{
						System.IO.Directory.Delete(@"C:\Program Files (x86)\Audigy\Common", true);
						
					}
					
					catch
					{
						
					}
				}
				if(System.IO.Directory.Exists(@"C:\Program Files (x86)\Audigy\Audigy2\"))
				{
					
					try
					{
						System.IO.Directory.Delete(@"C:\Program Files (x86)\Audigy\Audigy2", true);
						
					}
					
					catch
					{
						
					}
				}
			}
			
			
			
		}
	}
}
