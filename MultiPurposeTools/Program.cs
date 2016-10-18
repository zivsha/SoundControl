using Microsoft.VisualBasic;
using SoundControl;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
//***************************************************************************\
// Module Name: ConsoleApplication.vb
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

namespace MultiPurposeTools
{
    [Guid("49B8BE96-0A71-4C23-9FDF-86C24AC53690")]
    static class Program
    {
        [STAThread]
        public static void Main(string[] args)
        {
            using (Mutex mutex = new Mutex(false, "Global\\49B8BE96-0A71-4C23-9FDF-86C24AC53690"))
            {
                if (!mutex.WaitOne(0, false))
                {
                    Console.WriteLine("Instance already running");
                    return;
                }

                CommandLineParser options = new CommandLineParser();
                options.Parse(args);

                switch (options.Result)
                {
                    case CommandLineParser.CommandLineParserResult.OK:
                        Run(options.PlaybackDeviceName, options.RecordingDeviceName);
                        break;
                    case CommandLineParser.CommandLineParserResult.Usage:
                        Console.WriteLine(CommandLineParser.Usage);
                        break;
                    case CommandLineParser.CommandLineParserResult.Fail:
                        WriteConsoleError(options.ParseError);
                        break;
                    default:
                        throw new NotImplementedException();
                }
            }
        }

        private static void Run(string playbackDeviceName, string recordingDeviceName)
        {
            try
            {
                using (WindowsSound sound = new WindowsSound())
                {
                    try
                    {
                        SetDefaultAudioPlaybackDevice(sound, playbackDeviceName);
                        SetDefaultAudioRecordingDevice(sound, recordingDeviceName);
                    }
                    catch (InvalidOperationException ex)
                    {
                        WriteConsoleError(ex.Message);
                    }
                }
            }
            catch (WindowsSoundException ex)
            {
                WriteConsoleError(ex.ToString());
            }
        }

        private static void SetDefaultAudioPlaybackDevice(WindowsSound sound, string deviceName)
        {
            SetDefaultAudioDevice("playback", deviceName, sound.SetDefaultAudioPlaybackDevice);
        }

        private static void SetDefaultAudioRecordingDevice(WindowsSound sound, string deviceName)
        {
            SetDefaultAudioDevice("recording", deviceName, sound.SetDefaultAudioRecordingDevice);
        }

        private static void SetDefaultAudioDevice(string deviceTypeName, string deviceName, Action<string> action)
        {
            if ((string.IsNullOrEmpty(deviceName)))
            {
                return;
            }

            action.Invoke(deviceName);
            WriteConsoleSuccess(deviceTypeName, deviceName);
        }

        private static void WriteConsoleSuccess(string message, string device)
        {
            Console.WriteLine(string.Format(System.Globalization.CultureInfo.CurrentCulture, "The default audio {0} device was set to '{1}'.", message, device));
        }

        private static void WriteConsoleError(string message)
        {
            Console.Error.WriteLine();
            Console.Error.WriteLine(message);
        }

    }
}