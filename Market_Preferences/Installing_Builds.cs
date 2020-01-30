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
    ///The Installing_Builds recording.
    /// </summary>
    [TestModule("3cdd62d9-e2c5-4349-9ba7-d4cdab007f0f", ModuleType.Recording, 1)]
    public partial class Installing_Builds : ITestModule
    {
        /// <summary>
        /// Holds an instance of the Market_PreferencesRepository repository.
        /// </summary>
        public static Market_PreferencesRepository repo = Market_PreferencesRepository.Instance;

        static Installing_Builds instance = new Installing_Builds();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Installing_Builds()
        {
            txtInstallation_brandname = "";
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Installing_Builds Instance
        {
            get { return instance; }
        }

#region Variables

        string _txtInstallation_brandname;

        /// <summary>
        /// Gets or sets the value of variable txtInstallation_brandname.
        /// </summary>
        [TestVariable("fc0bf9e9-01d2-4730-8375-2b4b02899c1c")]
        public string txtInstallation_brandname
        {
            get { return _txtInstallation_brandname; }
            set { _txtInstallation_brandname = value; }
        }

        /// <summary>
        /// Gets or sets the value of variable TxtMkt.
        /// </summary>
        [TestVariable("ffe224a2-607e-44cd-80df-a0f4401ec54c")]
        public string TxtMkt
        {
            get { return repo.TxtMkt; }
            set { repo.TxtMkt = value; }
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

            Report.Log(ReportLevel.Info, "Delay", "Waiting for 10s.", new RecordItemIndex(0));
            Delay.Duration(10000, false);
            
            Installation_UserCodeCollection.Install_FSW(txtInstallation_brandname);
            Delay.Milliseconds(0);
            
            Installation_UserCodeCollection.If_Application_Is_Not_unInstall(txtInstallation_brandname);
            Delay.Milliseconds(0);
            
            Installation_UserCodeCollection.Install_FSW(txtInstallation_brandname);
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'SmartFit15InstallShieldWizard.ButtonOK' at Center.", repo.SmartFit15InstallShieldWizard.ButtonOKInfo, new RecordItemIndex(4));
            repo.SmartFit15InstallShieldWizard.ButtonOK.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Delay", "Waiting for 3s.", new RecordItemIndex(5));
            Delay.Duration(3000, false);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'SmartFit15InstallShieldWizard.ButtonNext' at Center.", repo.SmartFit15InstallShieldWizard.ButtonNextInfo, new RecordItemIndex(6));
            repo.SmartFit15InstallShieldWizard.ButtonNext.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Delay", "Waiting for 2s.", new RecordItemIndex(7));
            Delay.Duration(2000, false);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'SmartFit15InstallShieldWizard.ButtonNext' at Center.", repo.SmartFit15InstallShieldWizard.ButtonNextInfo, new RecordItemIndex(8));
            repo.SmartFit15InstallShieldWizard.ButtonNext.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Delay", "Waiting for 2s.", new RecordItemIndex(9));
            Delay.Duration(2000, false);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'SmartFit15InstallShieldWizard.ButtonNext' at Center.", repo.SmartFit15InstallShieldWizard.ButtonNextInfo, new RecordItemIndex(10));
            repo.SmartFit15InstallShieldWizard.ButtonNext.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Delay", "Waiting for 2s.", new RecordItemIndex(11));
            Delay.Duration(2000, false);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'SmartFit15InstallShieldWizard.IAcceptTheTermsInTheLicenseAgreem' at Center.", repo.SmartFit15InstallShieldWizard.IAcceptTheTermsInTheLicenseAgreemInfo, new RecordItemIndex(12));
            repo.SmartFit15InstallShieldWizard.IAcceptTheTermsInTheLicenseAgreem.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'SmartFit15InstallShieldWizard.ButtonNext' at Center.", repo.SmartFit15InstallShieldWizard.ButtonNextInfo, new RecordItemIndex(13));
            repo.SmartFit15InstallShieldWizard.ButtonNext.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Delay", "Waiting for 3s.", new RecordItemIndex(14));
            Delay.Duration(3000, false);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'SmartFit15InstallShieldWizard.CheckBoxStandalone' at Center.", repo.SmartFit15InstallShieldWizard.CheckBoxStandaloneInfo, new RecordItemIndex(15));
            repo.SmartFit15InstallShieldWizard.CheckBoxStandalone.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'SmartFit15InstallShieldWizard.ComboBox17451' at Center.", repo.SmartFit15InstallShieldWizard.ComboBox17451Info, new RecordItemIndex(16));
            repo.SmartFit15InstallShieldWizard.ComboBox17451.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse scroll Vertical by 600 units.", new RecordItemIndex(17));
            Mouse.ScrollWheel(600);
            Delay.Milliseconds(300);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'List1000.Australia' at Center.", repo.List1000.AustraliaInfo, new RecordItemIndex(18));
            repo.List1000.Australia.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Delay", "Waiting for 2s.", new RecordItemIndex(19));
            Delay.Duration(2000, false);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'SmartFit15InstallShieldWizard.ButtonNext' at Center.", repo.SmartFit15InstallShieldWizard.ButtonNextInfo, new RecordItemIndex(20));
            repo.SmartFit15InstallShieldWizard.ButtonNext.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Delay", "Waiting for 3s.", new RecordItemIndex(21));
            Delay.Duration(3000, false);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'SmartFit15InstallShieldWizard.ButtonNext' at Center.", repo.SmartFit15InstallShieldWizard.ButtonNextInfo, new RecordItemIndex(22));
            repo.SmartFit15InstallShieldWizard.ButtonNext.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Delay", "Waiting for 3s.", new RecordItemIndex(23));
            Delay.Duration(3000, false);
            
            Installation_UserCodeCollection.ExtraNext_Button();
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'SmartFit15InstallShieldWizard.ButtonInstall' at Center.", repo.SmartFit15InstallShieldWizard.ButtonInstallInfo, new RecordItemIndex(25));
            repo.SmartFit15InstallShieldWizard.ButtonInstall.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Wait", "Waiting 5s for the attribute 'Visible' to equal the specified value 'True'. Associated repository item: 'SmartFit15InstallShieldWizard.ButtonFinish'", repo.SmartFit15InstallShieldWizard.ButtonFinishInfo, new RecordItemIndex(26));
            repo.SmartFit15InstallShieldWizard.ButtonFinishInfo.WaitForAttributeEqual(5000, "Visible", "True");
            
            Report.Log(ReportLevel.Info, "Validation", "Validating Exists on item 'SmartFit15InstallShieldWizard.ButtonFinish'.", repo.SmartFit15InstallShieldWizard.ButtonFinishInfo, new RecordItemIndex(27));
            Validate.Exists(repo.SmartFit15InstallShieldWizard.ButtonFinishInfo);
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'SmartFit15InstallShieldWizard.ButtonFinish' at Center.", repo.SmartFit15InstallShieldWizard.ButtonFinishInfo, new RecordItemIndex(28));
            repo.SmartFit15InstallShieldWizard.ButtonFinish.Click();
            Delay.Milliseconds(200);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}