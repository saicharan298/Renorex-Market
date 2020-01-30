/*
 * Created by Ranorex
 * User: i-ray
 * Date: 03-12-2019
 * Time: 02:21
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
	public class Simulation_UserCodeCollection
	{
		// You can use the "Insert New User Code Method" functionality from the context menu,
		// to add a new method with the attribute [UserCodeMethod].
		[UserCodeMethod]
		public static void EnterDevicename(string devicename)
		{
			Ranorex.Text DeviceselectionTextbox="/form[@name~'Smart Fit' or @name~'Launcher' or @name~'Solus Max' or @name~'Interton Fitting' or @name~'Audigy' or @name~'Costco' or @processname='SmartFit' and @classname='Window']/text";
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
			Ranorex.Text DeviceselectionTextbox="/form[@name~'Smart Fit' or @name~'Launcher' or @name~'Solus Max' or @name~'Interton Fitting' or @name~'Audigy' or @processname='SmartFit' or @name~'Costco' and @classname='Window']/text";
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
