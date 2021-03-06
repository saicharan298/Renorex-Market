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
    ///The PrefernceWindow recording.
    /// </summary>
    [TestModule("31cbf27a-82e3-4d09-89cf-7eaf81153005", ModuleType.Recording, 1)]
    public partial class PrefernceWindow : ITestModule
    {
        /// <summary>
        /// Holds an instance of the Market_PreferencesRepository repository.
        /// </summary>
        public static Market_PreferencesRepository repo = Market_PreferencesRepository.Instance;

        static PrefernceWindow instance = new PrefernceWindow();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public PrefernceWindow()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static PrefernceWindow Instance
        {
            get { return instance; }
        }

#region Variables

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

            Report.Log(ReportLevel.Info, "Delay", "Waiting for 4s.", new RecordItemIndex(0));
            Delay.Duration(4000, false);
            
            Mouse_Click_ComboBox(repo.Window.ComboBoxInfo);
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'Popup.Thumb' at 4;131.", repo.Popup.ThumbInfo, new RecordItemIndex(2));
            repo.Popup.Thumb.Click("4;131");
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'Popup.PARTUpLeftRepeatButton' at 2;2.", repo.Popup.PARTUpLeftRepeatButtonInfo, new RecordItemIndex(3));
            repo.Popup.PARTUpLeftRepeatButton.Click("2;2");
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse scroll Vertical by 240 units.", new RecordItemIndex(4));
            Mouse.ScrollWheel(240);
            Delay.Milliseconds(500);
            
            //Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'Popup.PARTUpLeftRepeatButton' at 4;5.", repo.Popup.PARTUpLeftRepeatButtonInfo, new RecordItemIndex(5));
            //repo.Popup.PARTUpLeftRepeatButton.Click("4;5");
            //Delay.Milliseconds(200);
            
            Mouse_Click_SomeListItem(repo.Popup.SomeListItemInfo);
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Delay", "Waiting for 2s.", new RecordItemIndex(7));
            Delay.Duration(2000, false);
            
            Mouse_Click_Save_Button(repo.Window.Save_ButtonInfo);
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Delay", "Waiting for 2s.", new RecordItemIndex(9));
            Delay.Duration(2000, false);
            
            Mouse_Click_Close(repo.Window.CloseInfo);
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Delay", "Waiting for 4s.", new RecordItemIndex(11));
            Delay.Duration(4000, false);
            
            Mouse_Click_PARTPositiveButton(repo.Window.PARTPositiveButtonInfo);
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
