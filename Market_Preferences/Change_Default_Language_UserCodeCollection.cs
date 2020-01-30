/*
 * Created by Ranorex
 * User: i-ray
 * Date: 13-11-2019
 * Time: 00:22
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

using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;

namespace Market_Preferences
{
	/// <summary>
	/// Creates a Ranorex user code collection. A collection is used to publish user code methods to the user code library.
	/// </summary>
	[UserCodeCollection]
	public class Change_Default_Language_UserCodeCollection
	{
		
		
		/// <summary>
		/// This is a placeholder text. Please describe the purpose of the
		/// user code method here. The method is published to the user code library
		/// within a user code collection.
		/// </summary>
		[UserCodeMethod]
		public static void Mouse_Clicks()
		{
			Ranorex.ComboBox Defaultlanguage_Comb="/form[@classname='Window' and @orientation='None' and @processname='SmartFit']/container/combobox[1]";
			Mouse.Click(Defaultlanguage_Comb);
			
	//		Ranorex.Container txtContainer="/form[@classname='Popup' and @orientation='None' and @processname='SmartFit']/container[@classname='ScrollViewer']";
			//    		IList<ListItem> listCombo=txtContainer.FindChildren<ListItem>();
			//    		for(int i=0;i<=listCombo.Count-1;i++)
			//    		{
//
			//    		}
//			Ranorex.ListItem EnglishCombo="/form[@classname='Popup' and @orientation='None' and @processname='SmartFit']/container/listitem[1]";
//			if(EnglishCombo.Visible==true)
//			{
//				
//			}
//			else
//			{
//				
//			}
			
		}
		
		
		/// <summary>
		/// This is a placeholder text. Please describe the purpose of the
		/// user code method here. The method is published to the user code library
		/// within a user code collection.
		/// </summary>
		[UserCodeMethod]
		public static void EnglishCombo_click()
		{
			Ranorex.ListItem EnglishList="/form[@classname='Popup' and @orientation='None' and @processname='SmartFit']/container/listitem[1]";
			Mouse.Click(EnglishList);
		}
		
		/// <summary>
		/// This is a placeholder text. Please describe the purpose of the
		/// user code method here. The method is published to the user code library
		/// within a user code collection.
		/// </summary>
		[UserCodeMethod]
		public static void Save_button_Click()
		{
			Ranorex.Button saveButton="/form[@classname='Window' and @orientation='None' and @processname='SmartFit']/button[4]";
			Mouse.Click(saveButton);
		}
		
		
		/// <summary>
		/// This is a placeholder text. Please describe the purpose of the
		/// user code method here. The method is published to the user code library
		/// within a user code collection.
		/// </summary>
		[UserCodeMethod]
		public static void X_Click()
		{
			Ranorex.Button X_btn="/form[@classname='Window' and @orientation='None' and @processname='SmartFit']/button[@name='Close']";
			Mouse.Click(X_btn);
		}
		
		
		/// <summary>
		/// This is a placeholder text. Please describe the purpose of the
		/// user code method here. The method is published to the user code library
		/// within a user code collection.
		/// </summary>
		[UserCodeMethod]
		public static void Exit_withOut_Saving_Button()
		{
			Delay.Seconds(4);
		//	Ranorex.Button btnExit="/form[@classname='Window' and @orientation='None' and @processname='SmartFit']/button[@automationid='PART_PositiveButton']";
		Ranorex.Button btnExit="/form[@classname='Window' and @orientation='None' and @processname='SmartFit']/button[2]";
			Mouse.Click(btnExit);
		}
		/// <summary>
		/// This is a placeholder text. Please describe the purpose of the
		/// user code method here. The method is published to the user code library
		/// within a user code collection.
		/// </summary>
		[UserCodeMethod]
		public static string  Language_Check1(string returnval)
		{
			string DefaultSelected="";
			Ranorex.ComboBox txtCombobox="/form[@classname='Window' and @orientation='None' and @processname='SmartFit']/container/combobox[1]";
			string _Value=txtCombobox.Text;
			if(_Value=="English")
			{
				DefaultSelected="Yes";
			}
			else
			{
				DefaultSelected="No";
			}
			return DefaultSelected;
			Delay.Seconds(2);
		}
		// You can use the "Insert New User Code Method" functionality from the context menu,
		// to add a new method with the attribute [UserCodeMethod].
		[UserCodeMethod]
		public static void EnterDevicename(string devicename)
		{
			Ranorex.Text DeviceselectionTextbox="/form[@name~'Smart Fit' or @name~'Launcher' or @name~'Solus Max' or @name~'Interton Fitting' or @name~'Audigy' or @name~'Costco'  or @processname='SmartFit' or  @classname='Window'  or @processname='SolusMax']/text";
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
			Ranorex.Text DeviceselectionTextbox="/form[@name~'Smart Fit' or @name~'Launcher' or @name~'Solus Max' or @name~'Interton Fitting' or @name~'Audigy' or @name~'Costco' or @processname='SmartFit'  or @processname='SolusMax' or @classname='Window']/text";
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
		
		
		
		
	}
}
