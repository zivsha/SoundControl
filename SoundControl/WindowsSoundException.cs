using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
//***************************************************************************\
// Module Name: WindowsSoundException.vb
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

using System.Runtime.Serialization;
namespace SoundControl
{
    public class WindowsSoundException : Exception
    {

        public WindowsSoundException()
        {
        }

        public WindowsSoundException(string message) : base(message)
        {
        }

        protected WindowsSoundException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }

        public WindowsSoundException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}