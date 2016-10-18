using Microsoft.VisualBasic;
using Microsoft.Win32;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Threading;
//***************************************************************************\
// Module Name: DefaultAudioDeviceSetter.vb
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
using System.Windows.Forms;

namespace SoundControl
{
    class DefaultAudioDeviceSetter
    {
        //On the Windows 7 "Sound" control panel, setting both the default Playback and
        // Recording devices with System.Windows.Automation proceeds in the same way.
        // So, abstract this procedure by breaking it into a separate class.
        // We need only have the Playback or Recording tab from the Sound control panel window.

        //Windows10 registry values at HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\{Render,Capture} 
        // change when setting the default playback\record device
        //By keeping track on these values we can identify the device
        //Heuristic: the Render\Capture key must have "Level:0", "Level:1", "Level:2" values 
        //Heuristic: key "{a45c254e-df1c-4efd-8020-67d146a850e0},2" is the device's Type
        //Heuristic: key "{b3f8fa53-0004-438e-9003-51a46e139bfc},6" is the device's FriendlyName
        //The dictionary consists of <Key(guid), Type, FriendlyName> = <Level:0, Level:1, Level:2>
        private Dictionary<Tuple<string, string, string>, Tuple<long, long, long>> m_renderDevices; 
        private Dictionary<Tuple<string, string, string>, Tuple<long, long, long>> m_captureDevices;

        private readonly AutomationElement myDeviceTab;
        public DefaultAudioDeviceSetter(AutomationElement deviceTab)
        {
            myDeviceTab = deviceTab;
            BuildRegValuesDictionaries();
        }

        private void BuildRegValuesDictionaries()
        {
            m_renderDevices = new Dictionary<Tuple<string, string, string>, Tuple<long, long, long>>();
            m_captureDevices = new Dictionary<Tuple<string, string, string>, Tuple<long, long, long>>();

            ReadFrom("Render", ref m_renderDevices);
            ReadFrom("Capture", ref m_captureDevices);
        }

        private void ReadFrom(string subkey, ref Dictionary<Tuple<string, string, string>, Tuple<long, long, long>> deviceList)
        {
            deviceList.Clear();
            const string TypeKey = "{a45c254e-df1c-4efd-8020-67d146a850e0},2";
            const string FriendlyNameKey = "{b3f8fa53-0004-438e-9003-51a46e139bfc},6";

            using (RegistryKey key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\" + subkey))
            {
                if (key == null)
                {
                    return;
                }
                foreach (var deviceGuid in key.GetSubKeyNames())
                {
                    RegistryKey deviceKey = key.OpenSubKey(deviceGuid);
                    if (deviceKey == null)
                    {
                        continue;
                    }

                    try
                    {
                        long level0 = (long)deviceKey.GetValue("Level:0");
                        long level1 = (long)deviceKey.GetValue("Level:1");
                        long level2 = (long)deviceKey.GetValue("Level:2");

                        //Open Properties subkey
                        RegistryKey propertiesKey = deviceKey.OpenSubKey("Properties");
                        if (propertiesKey == null)
                        {
                            continue;
                        }

                        string name = Convert.ToString(propertiesKey.GetValue(TypeKey));
                        string friendlyName = Convert.ToString(propertiesKey.GetValue(FriendlyNameKey));
                        deviceList.Add(new Tuple<string, string, string>(deviceGuid, name, friendlyName), new Tuple<long, long, long>(level0, level1, level2));
                    }
                    catch (Exception)
                    {
                        continue;
                    }
                }
            }
        }

        public void Execute(string name)
        {
            //select our tab; brings its device list into view
            SelectDeviceTab();

            if ((name.Equals("SWITCH_TO_NEXT")))
            {
                AutomationElement device = GetNextDevice();
                if ((device == null))
                {
                    Console.WriteLine(string.Format(System.Globalization.CultureInfo.CurrentCulture, "Failed to find any available device"));
                    return;
                }
                device.SelectElement();
                ClickSetDefaultButton();
                string friendlyName = GetFriendlyName(device.Current.Name, myDeviceTab.Current.Name); 
                Console.WriteLine(string.Format(System.Globalization.CultureInfo.CurrentCulture, "Selected device was set to '{0} - {1}'.", device.Current.Name, friendlyName));
                WriteBaloonTipSuccess(myDeviceTab.Current.Name, string.Format("{0} - {1}", device.Current.Name, friendlyName));
                return;
            }

            //select audio device by name
            SelectDevice(name);

            //throw exception if we can't set audio device as default
            if ((!CanSetAudioDeviceAsDefault()))
            {
                throw new InvalidOperationException(string.Format("The audio device '{0}' could not be set as the default audio device. It is either not enabled, or is already the default audio device.", name));
            }

            //click on Set Default button
            ClickSetDefaultButton();

        }

        private string GetFriendlyName(string name1, string name2)
        {
            if (name2 == "Playback")
            {
                var prevList = new Dictionary<Tuple<string, string, string>, Tuple<long, long, long>>(m_renderDevices);
                ReadFrom("Render", ref m_renderDevices);
                return GetFriendlyName(name1, prevList, m_renderDevices);
            }
            if (name2 == "Recording")
            {
                var prevList = new Dictionary<Tuple<string, string, string>, Tuple<long, long, long>>(m_captureDevices);
                ReadFrom("Capture", ref m_captureDevices);
                return GetFriendlyName(name1, prevList, m_captureDevices);
            }
            return "";
        }

        private string GetFriendlyName(string name1, Dictionary<Tuple<string, string, string>, Tuple<long, long, long>> prevList, Dictionary<Tuple<string, string, string>, Tuple<long, long, long>> currentList)
        {
            foreach (var deviceInfo in prevList)
            {
                if (currentList.ContainsKey(deviceInfo.Key))
                {
                    var newValues = currentList[deviceInfo.Key];
                    if (!deviceInfo.Value.Equals(newValues))
                    {
                        return deviceInfo.Key.Item3;
                    }
                }
            }
            return "";
        }

        private AutomationElement GetNextDevice()
        {
            bool foundSelected = false;
            bool first = true;
            AutomationElement firstorNothing = null;
            foreach (AutomationElement device in AudioDevices())
            {
                if (first)
                {
                    firstorNothing = device;
                    first = false;
                }
                device.SelectElement();
                //If the device can't be selected - move to next device and mark it '
                if (!(CanSetAudioDeviceAsDefault()))
                {
                    foundSelected = true;
                    continue;
                }

                //if we found the device in the prev iteration, return it'
                if (foundSelected)
                {
                    return device;
                }
            }

            return firstorNothing;
        }

        private bool CanSetAudioDeviceAsDefault()
        {
            //we can set currently-selected audio device as default
            // if the Set Default button is enabled
            return SetDefaultButtonOnDeviceTab().Current.IsEnabled;
        }

        private void ClickSetDefaultButton()
        {
            SetDefaultButtonOnDeviceTab().ClickElement();
        }

        private void SelectDevice(string name)
        {
            AutomationElement device = AudioDevice(name);

            if ((device == null))
            {
                DefaultAudioDeviceSetter.ThrowOnAudioDeviceNameNotFound(name, AudioDeviceListText());
            }

            device.SelectElement();
        }

        private void SelectDeviceTab()
        {
            DeviceTab.SelectElement();
        }

        private AutomationElement DeviceTab
        {
            get { return myDeviceTab; }
        }

        private AutomationElement SetDefaultButtonOnDeviceTab()
        {
            const string SetDefaultButtonAutomationId = "1002";

            AutomationElement result = DeviceTab.FindFirstInChildrenByAutomationId(SetDefaultButtonAutomationId);

            if ((result == null))
            {
                throw new ElementNotAvailableException("Could not find Set Default button on Sound control panel window.");
            }

            return result;
        }

        private AutomationElement AudioDevice(string name)
        {
            //find device name in found list
            AutomationElement result = AudioDeviceList().FindFirstInChildrenByName(name);

            return result;
        }

        private string AudioDeviceListText()
        {
            System.Text.StringBuilder builder = new System.Text.StringBuilder();

            builder.AppendLine();
            builder.AppendLine();

            foreach (string item in AudioDeviceNames())
            {
                builder.Append("  ");
                builder.AppendLine(item);
            }

            return builder.ToString();
        }

        private IEnumerable<string> AudioDeviceNames()
        {
            List<string> result = new List<string>();

            foreach (AutomationElement item in AudioDevices())
            {
                result.Add(item.Current.Name);
            }

            return result;
        }

        private AutomationElementCollection AudioDevices()
        {
            return AudioDeviceList().FindAll(TreeScope.Children, Condition.TrueCondition);
        }

        private AutomationElement AudioDeviceList()
        {
            //the audio device list is the control that shows each audio endpoint

            const string AudioDeviceListAutomationId = "1000";

            AutomationElement result = DeviceTab.FindFirstInChildrenByAutomationId(AudioDeviceListAutomationId);

            if ((result == null))
            {
                throw new ElementNotAvailableException("Could not find list of audio devices on Sound control panel window.");
            }

            return result;
        }


        private static void ThrowOnAudioDeviceNameNotFound(string name, string deviceListText)
        {
            string errorText = string.Format("The audio device '{0}' could not be found. Check the spelling of the audio device and try again. Available audio devices: {1}", name, deviceListText);

            throw new InvalidOperationException(errorText);
        }

        private static void WriteBaloonTipSuccess(string deviceTypeName, string device)
        {
            using (NotifyIcon ni = new NotifyIcon())
            {
                ni.Visible = true;
                ni.Icon = SystemIcons.Information;

                ni.BalloonTipText = string.Format(System.Globalization.CultureInfo.CurrentCulture, "Current audio {0} device was set to '{1}'.", deviceTypeName.ToLower(), device);
                ni.BalloonTipTitle = string.Format(System.Globalization.CultureInfo.CurrentCulture, "Default {0} Device Changed", deviceTypeName);
                ni.BalloonTipIcon = ToolTipIcon.None;
                ni.ShowBalloonTip(3000);
                ni.Visible = false;
            }
        }

    }
}
