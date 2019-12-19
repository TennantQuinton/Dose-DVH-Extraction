using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace System.Windows.Forms
{
    static class FormExtensions
    {
        public static OpenFileDialog CreateOpenFileDialog()
        {

            // Create and format the file browsing window.
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Title = "Select the Input Information text file for Dose Extraction.",
                Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*",
                FilterIndex = 1,
                RestoreDirectory = false,
            };

            return openFileDialog;
        }
    }
}
