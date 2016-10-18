
using Microsoft.VisualBasic;
using SoundControl.Properties;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
//***************************************************************************\
// Module Name: SoundWindow.vb
// Project: SetDefaultAudioEndpoint http://sdae.codeplex.com/
// Copyright 2011 by jeff
// 
// This source is subject to the GNU General Public License version 2 (GPLv2).
// See http://www.gnu.org/licenses/gpl-2.0.html.
// All other rights reserved.
// 
// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
// EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED 
// WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
//***************************************************************************/

using System.Windows.Automation;
namespace SoundControl
{
    class SoundWindow
    {

        private readonly AutomationElement myWindow;

        private readonly bool myWindowWasAlreadyOpen;
        protected SoundWindow(AutomationElement window, bool windowWasAlreadyOpen)
        {
            myWindow = window;
            myWindowWasAlreadyOpen = windowWasAlreadyOpen;
        }

        public bool WasAlreadyOpen
        {
            get { return myWindowWasAlreadyOpen; }
        }

        public static SoundWindow Open()
        {

            //assume the Sound control panel window is already open
            bool soundWindowIsOpen = true;

            //use existing Sound window if it is open
            AutomationElement window = FindFirstSoundWindow();

            //Sound control panel window is not open, so open it
            if ((window == null))
            {
                //Sound control panel window was not already open
                soundWindowIsOpen = false;

                window = SoundWindow.OpenSoundControlPanelWindow();
            }

            return new SoundWindow(window, soundWindowIsOpen);
        }

        private static AutomationElement OpenSoundControlPanelWindow()
        {

            //this doesn't work: window process seems to enter idle state
            // before the actual AutomationElement window can be found
            //Dim windowProcess = Process.Start("control", "mmsys.cpl,,0")
            //Dim idleState As Boolean = windowProcess.WaitForInputIdle(5000)

            SoundWindow.StartControlPanelProcess();

            AutomationElement result = null;

            if ((!SoundWindow.TryFindSoundWindow(ref result)))
            {
                throw new ElementNotAvailableException("Could not open Sound control panel window.");
            }

            return result;
        }

        private static bool TryFindSoundWindow(ref AutomationElement window)
        {
            //calculate the end time to wait for the Sound control panel window to be found
            // by adding a certain number of milliseconds to the current time
            System.DateTime waitForEndTime = DateTime.Now.AddMilliseconds(Settings.Default.TryFindSoundWindowTimeoutMilliseconds);

            do
            {
                window = FindFirstSoundWindow();
            } while (!((window != null) || (System.DateTime.Now > waitForEndTime)));

            return window != null;
        }

        private static void StartControlPanelProcess()
        {
            System.Diagnostics.Process.Start("control", "mmsys.cpl,,0");
        }

        public void SetDefaultAudioPlaybackDevice(string name)
        {
            SetDefaultAudioDevice(name, PlaybackTab());
        }

        public void SetDefaultAudioRecordingDevice(string name)
        {
            SetDefaultAudioDevice(name, RecordingTab());
        }

        private void SetDefaultAudioDevice(string name, AutomationElement tabControl)
        {
            if ((myWindow == null))
            {
                throw new InvalidOperationException("The Sound control panel window is not open.");
            }
            else if ((string.IsNullOrEmpty(name)))
            {
                throw new ArgumentNullException("name");
            }

            DefaultAudioDeviceSetter command = new DefaultAudioDeviceSetter(tabControl);

            command.Execute(name);
        }

        private AutomationElement PlaybackTab()
        {
            return Tab("Playback");
        }

        private AutomationElement RecordingTab()
        {
            return Tab("Recording");
        }

        private AutomationElement Tab(string name)
        {
            AutomationElement result = TabControl().FindFirstInChildrenByName(name);

            if ((result == null))
            {
                throw new ElementNotAvailableException(string.Format(System.Globalization.CultureInfo.CurrentCulture, "Could not find {0} tab on Sound control panel window.", name));
            }

            return result;
        }

        private AutomationElement TabControl()
        {
            const string TabControlAutomationId = "12320";

            AutomationElement result = myWindow.FindFirstInChildrenByAutomationId(TabControlAutomationId);

            if ((result == null))
            {
                throw new ElementNotAvailableException("Could not find tabs on Sound control panel window.");
            }

            return result;
        }

        public void Close()
        {
            if ((myWindow == null))
            {
                throw new InvalidOperationException("The Sound control panel window is not open.");
            }

            WindowOkButton().ClickElement();
        }

        private AutomationElement WindowOkButton()
        {
            const string WindowOkButtonAutomationId = "1";

            AutomationElement result = myWindow.FindFirstInChildrenByAutomationId(WindowOkButtonAutomationId);

            if ((result == null))
            {
                throw new ElementNotAvailableException("Could not find OK button on Sound control panel window.");
            }

            return result;
        }

        private static AutomationElement FindFirstSoundWindow()
        {
            return AutomationElement.RootElement.FindFirstInChildrenByName("Sound");
        }

    }
}
