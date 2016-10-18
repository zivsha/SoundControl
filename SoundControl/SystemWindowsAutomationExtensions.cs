using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
//***************************************************************************\
// Module Name: SystemWindowsAutomationExtensions.vb
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
    static class SystemWindowsAutomationExtensions
    {
        public static AutomationElement FindFirstInChildrenByName(this AutomationElement element, string name)
        {
            return element.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.NameProperty, name, PropertyConditionFlags.IgnoreCase));
        }

        public static AutomationElement FindFirstInChildrenByAutomationId(this AutomationElement element, string automationId)
        {
            return element.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.AutomationIdProperty, automationId));
        }

        public static void SelectElement(this AutomationElement element)
        {
            SelectionItemPattern target = (SelectionItemPattern)element.GetCurrentPattern(SelectionItemPattern.Pattern);

            target.Select();
        }

        public static void ClickElement(this AutomationElement element)
        {
            InvokePattern target = (InvokePattern)element.GetCurrentPattern(InvokePattern.Pattern);

            target.Invoke();
        }
    }
}