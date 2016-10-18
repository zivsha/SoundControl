using System;

//***************************************************************************\
// Module Name: WindowsSound.vb
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
    public class WindowsSound : IDisposable
    {

        //WindowsSound is a facade, a "simple interface to a complex subsystem" (GoF)

        private bool myDisposedValue;

        private SoundWindow myWindow;
        public void SetDefaultAudioPlaybackDevice(string deviceName)
        {
            try
            {
                Window.SetDefaultAudioPlaybackDevice(deviceName);
            }
            catch (ElementNotAvailableException ex)
            {
                WindowsSound.ThrowWindowsSoundException(ex);
            }
        }

        public void SetDefaultAudioRecordingDevice(string deviceName)
        {
            try
            {
                Window.SetDefaultAudioRecordingDevice(deviceName);
            }
            catch (ElementNotAvailableException ex)
            {
                WindowsSound.ThrowWindowsSoundException(ex);
            }
        }

        private SoundWindow Window
        {
            get
            {
                try
                {
                    if ((myWindow == null))
                    {
                        myWindow = SoundWindow.Open();
                    }
                }
                catch (ElementNotAvailableException ex)
                {
                    WindowsSound.ThrowWindowsSoundException(ex);
                }

                return myWindow;
            }
        }

        private static void ThrowWindowsSoundException(ElementNotAvailableException ex)
        {
            throw new WindowsSoundException("There was an error processing the request.", ex);
        }

        // IDisposable
        protected virtual void Dispose(bool disposing)
        {
            if (!myDisposedValue)
            {
                if (disposing)
                {
                    // free other state (managed objects)
                    if (((myWindow != null) && (!myWindow.WasAlreadyOpen)))
                    {
                        myWindow.Close();
                    }
                }

                // TODO: free your own state (unmanaged objects)
                // set large fields to null
                myWindow = null;
            }
            myDisposedValue = true;
        }

        #region " IDisposable Support "
        // This code added by Visual Basic to correctly implement the disposable pattern.
        public void Dispose()
        {
            // Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        #endregion

    }
}
