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
    ///The Select_InstrumentSimulation recording.
    /// </summary>
    [TestModule("4e7f8dff-8644-4f59-93a4-e9ccd4e093c6", ModuleType.Recording, 1)]
    public partial class Select_InstrumentSimulation : ITestModule
    {
        /// <summary>
        /// Holds an instance of the Market_PreferencesRepository repository.
        /// </summary>
        public static Market_PreferencesRepository repo = Market_PreferencesRepository.Instance;

        static Select_InstrumentSimulation instance = new Select_InstrumentSimulation();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Select_InstrumentSimulation()
        {
            txtDevicename = "";
            TxtReturn = "";
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Select_InstrumentSimulation Instance
        {
            get { return instance; }
        }

#region Variables

        string _txtDevicename;

        /// <summary>
        /// Gets or sets the value of variable txtDevicename.
        /// </summary>
        [TestVariable("d5e0acf7-1257-43a3-b1e0-c8f271cbc8eb")]
        public string txtDevicename
        {
            get { return _txtDevicename; }
            set { _txtDevicename = value; }
        }

        string _TxtReturn;

        /// <summary>
        /// Gets or sets the value of variable TxtReturn.
        /// </summary>
        [TestVariable("78c1c50d-7349-4dd1-a2f6-426aabaa9fc4")]
        public string TxtReturn
        {
            get { return _TxtReturn; }
            set { _TxtReturn = value; }
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

            Report.Log(ReportLevel.Info, "Delay", "Waiting for 5s.", new RecordItemIndex(0));
            Delay.Duration(5000, false);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'Window.PARTContentHost' at Center.", repo.Window.PARTContentHostInfo, new RecordItemIndex(1));
            repo.Window.PARTContentHost.Click();
            Delay.Milliseconds(200);
            
            //Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FSWModule_Repo.Instrument_Selection_Textbox' at Center.", repo.FSWModule_Repo.Instrument_Selection_TextboxInfo, new RecordItemIndex(2));
            //repo.FSWModule_Repo.Instrument_Selection_Textbox.Click();
            //Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Delay", "Waiting for 2s.", new RecordItemIndex(3));
            Delay.Duration(2000, false);
            
            Change_Default_Language_UserCodeCollection.EnterDevicename(txtDevicename);
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Delay", "Waiting for 2s.", new RecordItemIndex(5));
            Delay.Duration(2000, false);
            
            Change_Default_Language_UserCodeCollection.Select_Device(txtDevicename);
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FSWModule_Repo.Remove_Button' at Center.", repo.FSWModule_Repo.Remove_ButtonInfo, new RecordItemIndex(7));
            repo.FSWModule_Repo.Remove_Button.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FSWModule_Repo.Simulate_Button' at Center.", repo.FSWModule_Repo.Simulate_ButtonInfo, new RecordItemIndex(8));
            repo.FSWModule_Repo.Simulate_Button.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Delay", "Waiting for 5s.", new RecordItemIndex(9));
            Delay.Duration(5000, false);
            
            Report.Log(ReportLevel.Info, "Wait", "Waiting 5s for the attribute 'Visible' to equal the specified value 'True'. Associated repository item: 'FSWModule_Repo.PARTContinueButton'", repo.FSWModule_Repo.PARTContinueButtonInfo, new RecordItemIndex(10));
            repo.FSWModule_Repo.PARTContinueButtonInfo.WaitForAttributeEqual(5000, "Visible", "True");
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FSWModule_Repo.PARTContinueButton' at Center.", repo.FSWModule_Repo.PARTContinueButtonInfo, new RecordItemIndex(11));
            repo.FSWModule_Repo.PARTContinueButton.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Wait", "Waiting 30s for the attribute 'Visible' to equal the specified value 'True'. Associated repository item: 'FSWModule_Repo.NavigationAutomationIdsMainAutomationId'", repo.FSWModule_Repo.NavigationAutomationIdsMainAutomationIdInfo, new RecordItemIndex(12));
            repo.FSWModule_Repo.NavigationAutomationIdsMainAutomationIdInfo.WaitForAttributeEqual(30000, "Visible", "True");
            
            Report.Log(ReportLevel.Info, "Delay", "Waiting for 5s.", new RecordItemIndex(13));
            Delay.Duration(5000, false);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
