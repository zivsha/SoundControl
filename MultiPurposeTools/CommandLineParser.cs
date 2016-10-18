using Microsoft.VisualBasic;
using MultiPurposeTools.Properties;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
//***************************************************************************\
// Module Name: CommandLineParser.vb
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
    class CommandLineParser
    {

        public enum CommandLineParserResult
        {
            Unparsed = 0,
            OK,
            Usage,
            Fail
        }

        private CommandLineParserResult myResult = CommandLineParserResult.Unparsed;
        private string myPlaybackDeviceName;
        private string myRecordingDeviceName;

        private string myParseError;
        public void Parse(string[] args)
        {
            if ((args.Length == 0))
            {
                myResult = CommandLineParserResult.Usage;
                return;
            }

            bool parseOk = true;

            foreach (string item in args)
            {
                if ((item.StartsWith("-", StringComparison.InvariantCulture) && (item.Length >= 2) && parseOk))
                {
                    switch (item.Substring(1, 1).ToUpper(System.Globalization.CultureInfo.InvariantCulture))
                    {
                        case "P":
                            myPlaybackDeviceName = item.Substring(2);
                            break;
                        case "R":
                            myRecordingDeviceName = item.Substring(2);
                            break;
                        case "?":
                        case "H":
                            myResult = CommandLineParserResult.Usage;
                            break;
                        default:
                            parseOk = false;
                            myParseError = item;
                            myResult = CommandLineParserResult.Fail;
                            break;
                    }
                }
                else {
                    parseOk = false;
                    myParseError = item;
                    myResult = CommandLineParserResult.Fail;
                }
            }

            if (((myResult != CommandLineParserResult.Fail) && (myResult != CommandLineParserResult.Usage) && parseOk))
            {
                myResult = CommandLineParserResult.OK;
            }
        }

        public CommandLineParserResult Result
        {
            get { return myResult; }
        }

        public string PlaybackDeviceName
        {
            get { return myPlaybackDeviceName; }
        }

        public string RecordingDeviceName
        {
            get { return myRecordingDeviceName; }
        }

        public static string Usage
        {
            get { return Resources.ConsoleApplicationUsageText; }
        }

        public string ParseError
        {
            get { return string.Format(System.Globalization.CultureInfo.CurrentCulture, "The command line option '{0}' is invalid.", myParseError) + Environment.NewLine + Environment.NewLine + CommandLineParser.Usage; }
        }

    }
}