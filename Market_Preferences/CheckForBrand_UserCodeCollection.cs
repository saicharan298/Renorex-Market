/*
 * Created by Ranorex
 * User: i-ray
 * Date: 10-12-2019
 * Time: 04:15
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
    public class CheckForBrand_UserCodeCollection
    {
        // You can use the "Insert New User Code Method" functionality from the context menu,
        // to add a new method with the attribute [UserCodeMethod].
        
        /// <summary>
        /// This is a placeholder text. Please describe the purpose of the
        /// user code method here. The method is published to the user code library
        /// within a user code collection.
        /// </summary>
        [UserCodeMethod]
        public static string BrandCheck(string buildname)
        {
        	string _Brand="";
        	buildname=buildname.Replace(" ",string.Empty);
			buildname=buildname.ToLower();
			if(buildname.Contains("smartfit"))
			{
				_Brand="SmartFit";
			}
			else if(buildname.Contains("solusmax"))
			{
				_Brand="SolusMax";
			}
			else if(buildname.Contains("audigy"))
			{
				_Brand="Audigy";
			}
			return _Brand;
        }
    }
}
