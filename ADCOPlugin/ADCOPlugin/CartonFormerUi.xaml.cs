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

        #endregion

        #region Public Members

        // Declare a SolidWorks instance field
        SldWorks.SldWorks swApp;

        // Declare a SolidWorks part field
        SldWorks.ModelDoc2 swPart;

        // INCH CONVERSION NUMBER: NO. OF INCHES IN 1 M - Most API functions assume the unit is in meters
        double INCH_CONVERSION = 39.370079;

        // Error/warning handlers - Only to be used if OpenDoc method is updated to newer version
        //int fileerror;
        //int filewarning;

        #endregion

        #region Demo Variables

        // Default path for the template test box - Will be dynamic in final iteration of package
        static string glueSrcPathDEFAULT = @"C:\Users\trent\Documents\glueTemplate.SLDPRT";

        // Destination path for copied file - will be dynamic/user-inputted in final iteration of the package
        static string glueDestPathDEFAULT = @"C:\Users\trent\Documents\";

        string glueDestPath;
        string glueSrcPath;

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
                LockContent.Visibility = System.Windows.Visibility.Hidden;
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

            // Before switching screens, checks to see if the user entered a project ID
            // If so, then the project ID will be the name of the actual part file and is appended to the default home path
            // If not, then some preset name is set
            if (projectID.Text != null)
            {
                glueDestPath = $@"{glueDestPathDEFAULT}{projectID.Text}.SLDPRT";
            }
            else
            {
                glueDestPath = $@"{glueDestPathDEFAULT}\CopyOfglueTemplate.SLDPRT";
            }

            // Change display to glue screen
            GlueScreen();

            // Open the default glue-formed carton template
            // GlueSet();

        }

        /// <summary>
        /// UI screen change - Set the displayed screen to that specific to glue-formed cartons
        /// </summary>
        private void GlueScreen()
        {
            // Show the glue screen content and hide all other content pages
            GlueContent.Visibility = System.Windows.Visibility.Visible;
            InitContent.Visibility = System.Windows.Visibility.Hidden;
            LockContent.Visibility = System.Windows.Visibility.Hidden;
            glueSourcePathBox.Text = glueSrcPathDEFAULT;
            glueDestPathBox.Text = glueDestPath;

        }

        /// <summary>
        /// Copy the template file and load the copied file to the SW screen
        /// </summary>
        public void GlueOpen(string destPath, string sourcePath)
        {

            // Check that the file at the destination path does not already exist
            if (!File.Exists(destPath))
            {
                // If the destination file does not already exist, then copy the template from the source path to the destination path
                File.Copy(sourcePath, destPath);
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
            return;
        }

        /// <summary>
        /// Fired when the apply button on the glue page is clicked
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GlueApplyButton_Click(object sender, RoutedEventArgs e)
        {
            GlueRead();
            GlueSet();
        }

        private void GlueRead()
        {
            
        }

        /// <summary>
        /// Adjusts the dimensions of a copied generic model - this is where most of the work is done
        /// </summary>
        public void GlueSet()
        {

            ThreadHelpers.RunOnUIThread(() =>
            {
                // Declare and initialize dimensions and paths based on glue screen fields
                double aDim = double.Parse(GlueAParam.Text) / INCH_CONVERSION;
                double bDim = double.Parse(GlueBParam.Text) / INCH_CONVERSION;
                double cDim = double.Parse(GlueCParam.Text) / INCH_CONVERSION;
                string destPath = glueDestPathBox.Text;
                string sourcePath = glueSourcePathBox.Text;
                
                // Copy and open the generic part
                GlueOpen(destPath, sourcePath);

                // Get the active document
                PartDoc part = swApp.ActiveDoc as PartDoc;

                // Get the part feature called exstrusionBase
                Feature Feat = part.FeatureByName("extrusionBase");

                // Select the feature extrusionBase - Replace current selection (normal click, not ctrl-click)
                Feat.Select2(false, -1);

                // Get dimension a of extrusionBase
                Dimension swDim = (Dimension)Feat.Parameter("a");

                // Set the new value of a as the value from the read-in field
                int errors = swDim.SetSystemValue3(aDim, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                // Get dimension b of extrusionBase
                swDim = (Dimension)Feat.Parameter("b");

                // Set the new value of b as the value from the read-in field
                errors = swDim.SetSystemValue3(bDim, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                // Get the part feature called extrusion
                Feat = part.FeatureByName("extrusion");

                // Replace the current selection
                Feat.Select2(false, -1);

                // Get dimension c of extrusion
                swDim = (Dimension)Feat.Parameter("c");

                // Set the new value of c as the value from the read-in field
                errors = swDim.SetSystemValue3(cDim, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                // Invoke all changes and rebuild the document
                part.EditRebuild();
            });
            return;
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

        private void GluePaperboard_Checked(object sender, RoutedEventArgs e)
        {

        }

        private void GlueCorrugated_Checked(object sender, RoutedEventArgs e)
        {

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
            LockContent.Visibility = System.Windows.Visibility.Visible;
            InitContent.Visibility = System.Windows.Visibility.Hidden;
            GlueContent.Visibility = System.Windows.Visibility.Hidden;
            LockSourcePathBox.Text = @"NOT\YET\IMPLEMENTED";
            LockDestPathBox.Text = @"NOT\YET\IMPLEMENTED";
        }

        private void LockBackButton_Click(object sender, RoutedEventArgs e)
        {
            initScreen();
        }

        private void LockApplyButton_Click(object sender, RoutedEventArgs e)
        {

        }

        private void LockResetButton_Click(object sender, RoutedEventArgs e)
        {

        }

        private void LockCorrugated_Checked(object sender, RoutedEventArgs e)
        {

        }

        private void LockPaperboard_Checked(object sender, RoutedEventArgs e)
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
