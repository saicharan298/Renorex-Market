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
    ///The Test_Device_User_CodeColletion recording.
    /// </summary>
    [TestModule("d5c0288e-1345-493c-9826-14ef01b462c7", ModuleType.Recording, 1)]
    public partial class Test_Device_User_CodeColletion : ITestModule
    {
        /// <summary>
        /// Holds an instance of the Market_PreferencesRepository repository.
        /// </summary>
        public static Market_PreferencesRepository repo = Market_PreferencesRepository.Instance;

        static Test_Device_User_CodeColletion instance = new Test_Device_User_CodeColletion();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Test_Device_User_CodeColletion()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Test_Device_User_CodeColletion Instance
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

            Test_Device_UserCodeCollection.Fetching_Test_Device_Options();
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
