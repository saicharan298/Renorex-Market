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
    ///The Validating_excel_and_FSW recording.
    /// </summary>
    [TestModule("4ecf5a59-5469-4377-8a6d-352385865f9b", ModuleType.Recording, 1)]
    public partial class Validating_excel_and_FSW : ITestModule
    {
        /// <summary>
        /// Holds an instance of the Market_PreferencesRepository repository.
        /// </summary>
        public static Market_PreferencesRepository repo = Market_PreferencesRepository.Instance;

        static Validating_excel_and_FSW instance = new Validating_excel_and_FSW();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Validating_excel_and_FSW()
        {
            Bimodal_MarketName = "";
            txtBrandname = "";
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Validating_excel_and_FSW Instance
        {
            get { return instance; }
        }

#region Variables

        string _Bimodal_MarketName;

        /// <summary>
        /// Gets or sets the value of variable Bimodal_MarketName.
        /// </summary>
        [TestVariable("a7b1e72b-6838-4da9-b17f-fd41a94dcca6")]
        public string Bimodal_MarketName
        {
            get { return _Bimodal_MarketName; }
            set { _Bimodal_MarketName = value; }
        }

        string _txtBrandname;

        /// <summary>
        /// Gets or sets the value of variable txtBrandname.
        /// </summary>
        [TestVariable("60945c49-27b2-4b7e-9445-165af33c0153")]
        public string txtBrandname
        {
            get { return _txtBrandname; }
            set { _txtBrandname = value; }
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

            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'SmartFit15InstallShieldWizard.MenuAutomationIdsHelpAction' at Center.", repo.SmartFit15InstallShieldWizard.MenuAutomationIdsHelpActionInfo, new RecordItemIndex(0));
            repo.SmartFit15InstallShieldWizard.MenuAutomationIdsHelpAction.Click();
            Delay.Milliseconds(200);
            
            Bi_modal_UserGuide_UserCodeCollection.Click_On_BiModaluserGuide(Bimodal_MarketName, txtBrandname);
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}