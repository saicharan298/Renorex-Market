﻿///////////////////////////////////////////////////////////////////////////////
//
// This file was automatically generated by RANOREX.
// DO NOT MODIFY THIS FILE! It is regenerated by the designer.
// All your modifications will be lost!
// http://www.ranorex.com
//
///////////////////////////////////////////////////////////////////////////////

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
using Ranorex.Core.Repository;

namespace Market_Preferences
{
#pragma warning disable 0436 //(CS0436) The type 'type' in 'assembly' conflicts with the imported type 'type2' in 'assembly'. Using the type defined in 'assembly'.
    /// <summary>
    ///The Validate_SolusMax_Links recording.
    /// </summary>
    [TestModule("3851f4a4-69d8-45d8-86df-16a8ad4e8ad8", ModuleType.Recording, 1)]
    public partial class Validate_SolusMax_Links : ITestModule
    {
        /// <summary>
        /// Holds an instance of the Market_PreferencesRepository repository.
        /// </summary>
        public static Market_PreferencesRepository repo = Market_PreferencesRepository.Instance;

        static Validate_SolusMax_Links instance = new Validate_SolusMax_Links();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Validate_SolusMax_Links()
        {
            txtMarketname = "";
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Validate_SolusMax_Links Instance
        {
            get { return instance; }
        }

#region Variables

        string _txtMarketname;

        /// <summary>
        /// Gets or sets the value of variable txtMarketname.
        /// </summary>
        [TestVariable("3cdd1221-3904-47f0-8d13-29a707bf5174")]
        public string txtMarketname
        {
            get { return _txtMarketname; }
            set { _txtMarketname = value; }
        }

#endregion

        /// <summary>
        /// Starts the replay of the static recording <see cref="Instance"/>.
        /// </summary>
        [System.CodeDom.Compiler.GeneratedCode("Ranorex", global::Ranorex.Core.Constants.CodeGenVersion)]
        public static void Start()
        {
            TestModuleRunner.Run(Instance);
        }

        /// <summary>
        /// Performs the playback of actions in this recording.
        /// </summary>
        /// <remarks>You should not call this method directly, instead pass the module
        /// instance to the <see cref="TestModuleRunner.Run(ITestModule)"/> method
        /// that will in turn invoke this method.</remarks>
        [System.CodeDom.Compiler.GeneratedCode("Ranorex", global::Ranorex.Core.Constants.CodeGenVersion)]
        void ITestModule.Run()
        {
            Mouse.DefaultMoveTime = 300;
            Keyboard.DefaultKeyPressTime = 100;
            Delay.SpeedFactor = 1.00;

            Init();

            Links_UserCodeCollection.Kill_All_Open_Browsers();
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Delay", "Waiting for 3s.", new RecordItemIndex(1));
            Delay.Duration(3000, false);
            
            Links_UserCodeCollection.SolusMax_WebSite(txtMarketname);
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Delay", "Waiting for 3s.", new RecordItemIndex(3));
            Delay.Duration(3000, false);
            
            Links_UserCodeCollection.Kill_All_Open_Browsers();
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
