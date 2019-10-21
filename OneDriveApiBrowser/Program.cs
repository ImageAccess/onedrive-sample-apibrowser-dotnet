// ------------------------------------------------------------------------------
//  Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

using System;
using System.Windows.Forms;

namespace OneDriveApiBrowser
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        private static void Main ()
        {
            Application.EnableVisualStyles ();
            Application.SetCompatibleTextRenderingDefault (false);
            Application.Run (new FormBrowser ());
        }
    }
}