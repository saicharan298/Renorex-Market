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
    ///The WebUpdates_Excel_Comaparison recording.
    /// </summary>
    [TestModule("c45ae338-536b-4537-9bc2-26d47db81d78", ModuleType.Recording, 1)]
    public partial class WebUpdates_Excel_Comaparison : ITestModule
    {
        /// <summary>
        /// Holds an instance of the Market_PreferencesRepository repository.
        /// </summary>
        public static Market_PreferencesRepository repo = Market_PreferencesRepository.Instance;

        static WebUpdates_Excel_Comaparison instance = new WebUpdates_Excel_Comaparison();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public WebUpdates_Excel_Comaparison()
        {
            txtmarket = "";
            txtBuildname = "";
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static WebUpdates_Excel_Comaparison Instance
        {
            get { return instance; }
        }

#region Variables

        string _txtmarket;

        /// <summary>
        /// Gets or sets the value of variable txtmarket.
        /// </summary>
        [TestVariable("536c7c8f-838c-492d-bb6a-ca9fab0bf28c")]
        public string txtmarket
        {
            get { return _txtmarket; }
            set { _txtmarket = value; }
        }

        string _txtBuildname;

        /// <summary>
        /// Gets or sets the value of variable txtBuildname.
        /// </summary>
        [TestVariable("7d01bb1c-891d-4a68-907a-a3963bcf4771")]
        public string txtBuildname
        {
            get { return _txtBuildname; }
            set { _txtBuildname = value; }
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

            WebUpdates_UserCodeCollection.WebupdatesExcel_Comaparison(txtmarket, txtBuildname);
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
