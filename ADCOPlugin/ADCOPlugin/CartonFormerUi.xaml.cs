using System.Windows.Controls;
using System.Diagnostics;
using Microsoft.VisualBasic;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.IO;
using System.Windows;
using Microsoft.Win32;
using static AngelSix.SolidDna.SolidWorksEnvironment;
using AngelSix.SolidDna;
using System;
using SldWorks;
using SwConst;

namespace ADCOPlugin
{
    
    /// <summary>
    /// Interaction logic for MyAddinControl.xaml
    /// </summary>
    public partial class MyAddinControl : UserControl
    {
        // NOTE: CHANGE STRINGS SPECIFYING THE FILE LOCATIONS TO THOSE YOU WANT TO USE FOR YOUR MACHINE: I.E. C:\\USERS\\TRENT WILL NOT WORK ON YOUR COMPUTER

        #region Private Members

        private const string typeGlue = "GLUE";
        private string typeLock;
        private const string lockPath = "LOCKPATH";
        private const string mCustomPropertyGlueA = "GlueA";
        private const string mCustomPropertyGlueB = "GlueB";
        private const string mCustomPropertyGlueC = "GlueC";

        #endregion

        #region Public Members

        // Declare a SolidWorks instance field
        SldWorks.SldWorks swApp;

        // Declare a SolidWorks part field
        SldWorks.ModelDoc2 swPart;

        // Error/warning handlers - Only to be used if OpenDoc method is updated to newer version
        int fileerror;
        int filewarning;

        // Default path for the template test box - Will be dynamic in final iteration of package
        public string gluePath = "C:\\Users\\trent\\Documents\\glueTemplate.SLDPRT";

        #endregion

        #region Constructor

        /// <summary>
        /// Default constructor
        /// </summary>
        public MyAddinControl()
        {
           InitializeComponent();
        }


        #endregion

        #region Load-Up/Initial Screen Functions

        /// <summary>
        /// Fired when the plugin first loads up
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void UserControl_Loaded(object sender, System.Windows.RoutedEventArgs e)
        {
            // Get the active SolidWorks application handle right after the plugin is loaded up
            swApp = (SldWorks.SldWorks)Marshal.GetActiveObject("SldWorks.Application");

            // By default show the Initial Screen (Former type selection)
            initScreen();

        }

        /// <summary>
        /// Initial screen to display when the plugin is first loaded in
        /// </summary>
        private void initScreen()
        {
            ThreadHelpers.RunOnUIThread(() =>
            {
                // Hide all other content except that which should be displayed on initial screen
                GlueContent.Visibility = System.Windows.Visibility.Hidden;
                InitContent.Visibility = System.Windows.Visibility.Visible;

            });
        }

        #endregion

        #region Glue-Formed Carton Functions

        /// <summary>
        /// Fired when the glue button on the initial page is clicked
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GlueButton_Click(object sender, RoutedEventArgs e)
        {
            // Change display to glue screen
            GlueScreen();

            // Open the default glue-formed carton template
            GlueOpen();

        }

        /// <summary>
        /// UI screen change - Set the displayed screen to that specific to glue-formed cartons
        /// </summary>
        private void GlueScreen()
        {
            // Show the glue screen content and hide all other content pages
            GlueContent.Visibility = System.Windows.Visibility.Visible;
            InitContent.Visibility = System.Windows.Visibility.Hidden;

        }

        /// <summary>
        /// Copy the template file and load the copied file to the SW screen
        /// </summary>
        public void GlueOpen()
        {
            // Destination path for copied file - will be dynamic/user-inputted in final iteration of the package
            string destPath = "C:\\Users\\trent\\Documents\\glueResult.SLDPRT";

            // Check that the file at the destination path does not already exist
            if (!File.Exists(destPath))
            {
                // If the destination file does not already exist, then copy the template from the source path to the destination path
                File.Copy(gluePath, destPath);

            }

            try
            {
                // Try to open the copied document
                swPart = swApp.OpenDoc(destPath, (int)swDocumentTypes_e.swDocPART);

            }
            catch (Exception)
            {
                // If an exception is thrown, notify the user that the file was not able to be opened
                MessageBox.Show(string.Format("File Open Failed"));

                // Clear any unused objects - should clean up any COM-related errors that arise after debug close
                Marshal.CleanupUnusedObjectsInCurrentContext();

                // Return to the parent function (GlueButtonClick)
                return;
            }
        }

        /// <summary>
        /// Fired when the apply button on the glue page is clicked
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GlueApplyButton_Click(object sender, RoutedEventArgs e)
        {

        }

        /// <summary>
        /// Fired when the reset button on the glue page is clicked
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GlueResetButton_Click(object sender, RoutedEventArgs e)
        {

        }

        /// <summary>
        /// Fired when the back button on the glue page is clicked
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GlueBackButton_Click(object sender, RoutedEventArgs e)
        {
            // When the back button is clicked, send the UI back to the previous page (initial screen)
            initScreen();
        }

        #endregion

        #region Lock-Formed Carton Functions

        /// <summary>
        /// Fired when the lock button on the initial page is clicked
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void LockButton_Click(object sender, RoutedEventArgs e)
        {

        }

        #endregion

        #region General Functions / WIP

        /// <summary>
        /// WIP used to route file-selection/opening process depending on user-input parameters
        /// </summary>
        private void FileHandling()
        {
           
        }

        #endregion
    }
}
