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
using System.Globalization;
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

        

        

        // INCH CONVERSION NUMBER: NO. OF INCHES IN 1 M - Most API functions assume the unit is in meters
        double INCH_CONVERSION = 39.370079;

        // Error/warning handlers - Only to be used if OpenDoc method is updated to newer version
        // int fileerror;
        // int filewarning;

        #endregion

        #region Demo Variables
        
        //Get current windows user for proper default destination
        static string userName = System.Environment.UserName;
        

        //Default routing for part/assembly files
        //Home folder of all template parts/assemblies
        static string glueSrcLibDEFAULT = $@"C:\Users\{userName}\Documents\ADCO Carton Former Templates";

        //Type of carton (glue/lock)
        static string[] formerType = { "GLUE", "LOCK" };

        //Overall component folder names
        static string[] formerElement = { "MANDREL", "FORMER PLATE" };

        //Part files in the glue mandrel domain
        static string[] glueMandrelParts = { "R&D5959 SHOW CARTON 5-30-18 LG.SLDPRT", "R&D15D350A.SLDPRT" , "R&D5086-21-108-H.SLDPRT", "R&D5321-31 SMC MGPM20N-100_MGPRod.SLDPRT", "R&D5321-31 SMC MGPM20N-100_MGPTube.SLDPRT", "R&D5634-15-108.SLDPRT", "R&D5959-11-104-H.SLDPRT","R&D5959-11-105-H.SLDPRT","R&D5959-11-106-3-H.SLDPRT","R&D5959-11-106-H.SLDPRT","R&D5959-11-107.SLDPRT","R&D5959-11-109-H.SLDPRT","R&D5959-11-110-H.SLDPRT","R&D06451-1062.sldprt","R&DAS2211FG-N01-07S.SLDPRT"};

        //Assembly files in the glue mandrel domain
        static string[] glueMandrelAssemblies = { "R&D5321-31 SMC MGPM20N-100.SLDASM", "R&DMGPM20N-100.SLDASM", "R&D5959-11-001.SLDASM" };

        //Default home folder of resulting project folder
        static string destLibDEFAULT = $@"C:\Users\{userName}\Documents";

        //Destination path to be set
        string destPath = destLibDEFAULT;
        string srcPath = glueSrcLibDEFAULT;
        static string lockSrcPathDEFAULT = @"C:\Users\trent\Documents\";

        //Overarching SW variables
        // Declare a SolidWorks instance field
        public SldWorks.SldWorks swApp;
        // Declare a SolidWorks part doc field
        public PartDoc swPart;
        // Declare a SolidWorks model doc field
        public ModelDoc2 swModel;
        // Declare a Solidworks feature field
        public Feature swFeat;
        // Declare a SolidWorks dimension field
        public Dimension swDim;
        // Declare a Solidworks assembly doc field
        public AssemblyDoc swAssem;
        // Declarea a model document extension field
        public ModelDocExtension modelExt;


        // Destination path for copied file - will be dynamic/user-inputted in final iteration of the package
        //static string glueDestPathDEFAULT = @"C:\Users\trent\Documents\";
        static string lockDestPathDEFAULT = @"C:\Users\trent\Documents\";



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

            // Before switching screens, checks to see if the user entered a project ID or customer name
            // If so, then the project ID will be the name of the actual part file and is appended to the default home path
            // If not, then some preset name is set
            if (projectID.Text != "" || customer.Text != "")
            {
                destPath = $@"{destLibDEFAULT}\{customer.Text}{projectID.Text}";
            }
            else
            {
                destPath = $@"{destLibDEFAULT}\TESTCOPY";
            }

            srcPath = $@"{glueSrcLibDEFAULT}\{formerType[0]}";

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
            glueSourcePathBox.Text = srcPath;
            glueDestPathBox.Text = destPath;

        }

        /// <summary>
        /// Copy the template file and load the copied file to the SW screen
        /// </summary>
        private void GlueOpen(bool copystate, int type, int idx, int component)
        {
            if (copystate == true && !Directory.Exists(destPath))
            {
                // Source, Destination, Copy Subdirectories?
                DirectoryCopy(srcPath, destPath, true);
            }

            if(copystate == false && type == 0)
            {
                //MessageBox.Show($@"{destPath}\{formerElement[component]}\{glueMandrelParts[idx]}");
                swModel = swApp.OpenDoc($@"{destPath}\{formerElement[component]}\{glueMandrelParts[idx]}", (int)swDocumentTypes_e.swDocPART);
                swPart = (PartDoc)swApp.ActiveDoc;
            }

            if(copystate == false && type == 1)
            {
                swModel = swApp.OpenDoc($@"{destPath}\{formerElement[component]}\{glueMandrelAssemblies[idx]}", (int)swDocumentTypes_e.swDocASSEMBLY);
                swAssem = (AssemblyDoc)swApp.ActiveDoc;
                modelExt = (ModelDocExtension)swModel.Extension;
            }

            //ModelDocExtension modelExt = (ModelDocExtension)model.Extension;
            //try
            //{
            //    // Try to open the copied document
            //swPart = swApp.OpenDoc(destPath, (int)swDocumentTypes_e.swDocPART);

            //}
            //catch (Exception)
            //{
            //    // If an exception is thrown, notify the user that the file was not able to be opened
            //    MessageBox.Show(string.Format("File Open Failed"));

            //    // Clear any unused objects - should clean up any COM-related errors that arise after debug close
            //    Marshal.CleanupUnusedObjectsInCurrentContext();

            //    // Return to the parent function (GlueButtonClick)
            //    return;
            //}
            return;
        }

        /// <summary>
        /// Fired when the apply button on the glue page is clicked
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GlueApplyButton_Click(object sender, RoutedEventArgs e)
        {
            int error = GlueRead();
            if (error == 0)
            {
                GlueSet();
                return (0);
            }
        }

        private int GlueRead()
        {
            if (GlueAParam.Text == "" || GlueBParam.Text == "" || GlueCParam.Text == "" || GlueDParam.Text == "" || GlueEParam.Text == "")
            {
                MessageBoxImage icon = MessageBoxImage.Warning;
                MessageBoxButton button = MessageBoxButton.OK;
                MessageBox.Show("You must enter values for all box dimensions.", "", button, icon);
                return (1);
            }

            if(double.Parse(GlueAParam.Text) < 7 || double.Parse(GlueAParam.Text) > 30)
            {
                MessageBoxImage icon = MessageBoxImage.Warning;
                MessageBoxButton button = MessageBoxButton.OK;
                MessageBox.Show("Value of A must be between 7\" and 30\"", "", button, icon);
                return (1);
            }

            if (double.Parse(GlueBParam.Text) < 7 || double.Parse(GlueBParam.Text) > 26)
            {
                MessageBoxImage icon = MessageBoxImage.Warning;
                MessageBoxButton button = MessageBoxButton.OK;
                MessageBox.Show("Value of B must be between 7\" and 26\"", "", button, icon);
                return (1);
            }

            if (double.Parse(GlueCParam.Text) < 7 || double.Parse(GlueCParam.Text) > 19)
            {
                MessageBoxImage icon = MessageBoxImage.Warning;
                MessageBoxButton button = MessageBoxButton.OK;
                MessageBox.Show("Value of C must be between 7\" and 19\"", "", button, icon);
                return (1);
            }

            if (double.Parse(GlueDParam.Text) < 5.5 || double.Parse(GlueDParam.Text) > 10)
            {
                MessageBoxImage icon = MessageBoxImage.Warning;
                MessageBoxButton button = MessageBoxButton.OK;
                MessageBox.Show("Value of D must be between 7\" and 10\"", "", button, icon);
                return (1);
            }

            if (double.Parse(GlueEParam.Text) < 1 || double.Parse(GlueEParam.Text) > 6)
            {
                MessageBoxImage icon = MessageBoxImage.Warning;
                MessageBoxButton button = MessageBoxButton.OK;
                MessageBox.Show("Value of E must be between 1\" and 6\"", "", button, icon);
                return (1);
            }

            return (0);
        }

        /// <summary>
        /// Adjusts the dimensions of a copied generic model - this is where most of the work is done
        /// </summary>
        private void GlueSet()
        {

            ThreadHelpers.RunOnUIThread(() =>
            {
                // Declare and initialize dimensions and paths based on glue screen fields
                double aDim = double.Parse(GlueAParam.Text) / INCH_CONVERSION;
                double bDim = double.Parse(GlueBParam.Text) / INCH_CONVERSION;
                double cDim = double.Parse(GlueCParam.Text) / INCH_CONVERSION;
                double dDim = double.Parse(GlueDParam.Text) / INCH_CONVERSION;
                double eDim = double.Parse(GlueEParam.Text) / INCH_CONVERSION;
                string destPath = glueDestPathBox.Text;
                string sourcePath = glueSourcePathBox.Text;
                int errors;

                // Copy and open the template library
                GlueOpen(true,-1,-1,-1);

                // Initialize index of part file in static array
                int idx = 0;
                int TYPE_PART = 0;
                int TYPE_ASSEM = 1;
                int COMPONENT_MAN = 0;
                int COMPONENT_FP = 1;

                // Cycle through mandrel parts
                while (true)
                {
                    


                    switch (idx)
                    {
                        case 0:
                            //GlueOpen(false, TYPE_PART, idx, COMPONENT_MAN);
                            //swFeat = swPart.FeatureByName("Sketch1");
                            //swFeat.Select2(false, -1);
                            //swDim = (Dimension)swFeat.Parameter("A1");
                            //errors = swDim.SetSystemValue3(aDim - 3 / INCH_CONVERSION * eDim - cDim / INCH_CONVERSION,(int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration,null);
                            //swDim = (Dimension)swFeat.Parameter("B1");
                            //errors = swDim.SetSystemValue3(bDim-0.125/INCH_CONVERSION, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
                            //swDim = (Dimension)swFeat.Parameter("C1");
                            //errors = swDim.SetSystemValue3(cDim, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
                            //swDim = (Dimension)swFeat.Parameter("C2");
                            //errors = swDim.SetSystemValue3(cDim, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
                            //swDim = (Dimension)swFeat.Parameter("D1");
                            //errors = swDim.SetSystemValue3(dDim, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
                            //swDim = (Dimension)swFeat.Parameter("E1");
                            //errors = swDim.SetSystemValue3(eDim, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
                            //swDim = (Dimension)swFeat.Parameter("E2");
                            //errors = swDim.SetSystemValue3(eDim, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
                            //swDim = (Dimension)swFeat.Parameter("E3");
                            //errors = swDim.SetSystemValue3(eDim, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
                            //swDim = (Dimension)swFeat.Parameter("E4");
                            //errors = swDim.SetSystemValue3(eDim, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
                            //swPart.EditRebuild();
                            //swModel.Save();
                            //swApp.CloseDoc($@"{destPath}\{formerElement[COMPONENT_MAN]}\{glueMandrelParts[idx]}");
                            idx = 6;
                            break;

                        case 6:
                            //MessageBox.Show("CASE 5");
                            GlueOpen(false, TYPE_PART, idx,COMPONENT_MAN);
                            swFeat = swPart.FeatureByName("Extrude1");
                            swFeat.Select2(false, -1);
                            swDim = (Dimension)swFeat.Parameter("MountLength");
                            errors = swDim.SetSystemValue3(cDim - 0.375 * 2.0 / INCH_CONVERSION, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
                            swPart.EditRebuild();
                            swModel.Save();
                            swApp.CloseDoc($@"{destPath}\{formerElement[COMPONENT_MAN]}\{glueMandrelParts[idx]}");
                            idx++;
                            break;

                        case 7:
                            //MessageBox.Show("CASE 6");
                            GlueOpen(false, TYPE_PART, idx, COMPONENT_MAN);
                            if (cDim <= 7.5/INCH_CONVERSION)
                            {
                                swFeat = swPart.FeatureByName("Cut-Extrude1");
                                swFeat.Select2(false, -1);
                                bool suppressionState = swModel.EditSuppress2();
                            }
                            swFeat = swPart.FeatureByName("Extrude1");
                            swFeat.Select2(false, -1);
                            swDim = (Dimension)swFeat.Parameter("SpreadLength");
                            errors = swDim.SetSystemValue3(cDim - 0.375 * 2.0 / INCH_CONVERSION, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
                            swPart.EditRebuild();
                            swModel.Save();
                            swApp.CloseDoc($@"{destPath}\{formerElement[COMPONENT_MAN]}\{glueMandrelParts[idx]}");
                            idx = 9;
                            break;

                        case 9:
                            //MessageBox.Show("CASE 8");
                            GlueOpen(false, TYPE_PART, idx, COMPONENT_MAN);
                            swFeat = swPart.FeatureByName("Sketch1");
                            swFeat.Select2(false, -1);
                            swDim = (Dimension)swFeat.Parameter("MandrelSideWidth");
                            errors = swDim.SetSystemValue3(dDim, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                            swFeat = swPart.FeatureByName("Sketch3");
                            swFeat.Select2(false, -1);
                            swDim = (Dimension)swFeat.Parameter("MandrelSideHoles");
                            double MandrelSideHolesdim = 0.5009 * (dDim - 9.1875 / INCH_CONVERSION) + 2.589 / INCH_CONVERSION;
                            errors = swDim.SetSystemValue3(MandrelSideHolesdim, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
                            swPart.EditRebuild();
                            swModel.Save();
                            swApp.CloseDoc($@"{destPath}\{formerElement[COMPONENT_MAN]}\{glueMandrelParts[idx]}");
                            idx = 10;
                            break;

                        case 10:
                            //MessageBox.Show("CASE 9");
                            double PushPlateWidth = dDim - 1.406 * 2.0/INCH_CONVERSION; 
                            GlueOpen(false, TYPE_PART, idx, COMPONENT_MAN);
                            if (PushPlateWidth <= 4.0/INCH_CONVERSION)
                            {
                                Debug.Print("Entered suppression if statement");
                                swFeat = swPart.FeatureByName("Cut-Extrude3");
                                swFeat.Select2(false, -1);
                                swFeat = swPart.FeatureByName("Fillet2");
                                swFeat.Select2(true, -2);
                                swFeat = swPart.FeatureByName("Fillet3");
                                swFeat.Select2(true, -3);

                                Debug.Print("Selected multiple parts");
                                bool suppressionState = swModel.EditSuppress2();
                            }
                            Debug.Print("Made it past if statement");
                            swFeat = swPart.FeatureByName("Base-Flange1");
                            swFeat.Select2(false, -4);
                            swDim = (Dimension)swFeat.Parameter("PushPlateWidth");
                            errors = swDim.SetSystemValue3(PushPlateWidth, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                            swFeat = swPart.FeatureByName("Sketch1");
                            swFeat.Select2(false, -4);
                            swDim = (Dimension)swFeat.Parameter("PushPlateLength");
                            double PushPlateLength = cDim - 0.828 * 2.0 / INCH_CONVERSION;
                            errors = swDim.SetSystemValue3(PushPlateLength, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
                            swPart.EditRebuild();
                            swModel.Save();
                            swApp.CloseDoc($@"{destPath}\{formerElement[COMPONENT_MAN]}\{glueMandrelParts[idx]}");
                            idx = 11;
                            break;

                        default:
                            MessageBox.Show("Something went wrong!");
                            break;
                    }

                    if (idx == 11)
                    {
                        break;
                    }
                        
                }

                idx = glueMandrelAssemblies.Length - 1;
                GlueOpen(false, TYPE_ASSEM, idx, COMPONENT_MAN);
                //PartDoc part = (PartDoc) swApp.ActiveDoc;
                ////ModelDoc2 swModel = swApp.ActiveDoc as ModelDoc2;

                //// Get the part feature called exstrusionBase
                //Feature Feat = part.FeatureByName("extrusionBase");

                //// Select the feature extrusionBase - Replace current selection (normal click, not ctrl-click)
                //Feat.Select2(false, -1);

                //// Get dimension a of extrusionBase
                ////Dimension swDim = (Dimension) Feat.Parameter("a");

                //// Set the new value of a as the value from the read-in field
                ////int errors = swDim.SetSystemValue3(aDim, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                //// Get dimension b of extrusionBase
                //swDim = (Dimension)Feat.Parameter("b");

                //// Set the new value of b as the value from the read-in field
                //errors = swDim.SetSystemValue3(bDim, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                //// Get the part feature called extrusion
                //Feat = part.FeatureByName("extrusion");

                //// Replace the current selection
                //Feat.Select2(false, -1);

                //// Get dimension c of extrusion
                //swDim = (Dimension)Feat.Parameter("c");

                //// Set the new value of c as the value from the read-in field
                //errors = swDim.SetSystemValue3(cDim, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                //// Invoke all changes and rebuild the document, save afterwards
                //part.EditRebuild();
                //swModel.Save();

            });
            return;
        }

        private void glueDestBrowse_Click(object sender, RoutedEventArgs e)
        {

        }

        private void glueSourceBrowse_Click(object sender, RoutedEventArgs e)
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
            LockSourcePathBox.Text = $@"{lockSrcPathDEFAULT}\lockTemplate";
            
            if (projectID.Text == "")
            {
                LockDestPathBox.Text = $@"{lockDestPathDEFAULT}Copy of lockTemplate";
            }
            else
            {
                LockDestPathBox.Text = $@"{lockDestPathDEFAULT}{projectID.Text} lockTemplate";
            }

            LockScreen();
        }

        private void LockScreen()
        {
            GlueContent.Visibility = System.Windows.Visibility.Hidden;
            InitContent.Visibility = System.Windows.Visibility.Hidden;
            LockContent.Visibility = System.Windows.Visibility.Visible;
        }

        private void LockBackButton_Click(object sender, RoutedEventArgs e)
        {
            initScreen();
        }

        private void LockApplyButton_Click(object sender, RoutedEventArgs e)
        {
            lockRead();
            lockSet();

        }

        private void lockSet()
        {
            ThreadHelpers.RunOnUIThread(() =>
            {


            if (LockAParam.Text == "" || LockBParam.Text == "" || LockCParam.Text == "")
            {
                MessageBoxImage icon = MessageBoxImage.Warning;
                MessageBoxButton button = MessageBoxButton.OK;
                MessageBox.Show("You must enter values for box dimensions.", "", button, icon);
                return;
            }

            // Declare and initialize dimensions and paths based on glue screen fields
            double aDim = double.Parse(LockAParam.Text) / INCH_CONVERSION;
            double bDim = double.Parse(LockBParam.Text) / INCH_CONVERSION;
            double cDim = double.Parse(LockCParam.Text) / INCH_CONVERSION;
            string destPath = LockDestPathBox.Text;
            string sourcePath = LockSourcePathBox.Text;
            string Component1 = "Block1.SLDPRT";
            string Component2 = "Block2.SLDPRT";
            string Component3 = "Cylinder.SLDPRT";

            MessageBox.Show("STARTING ON 1ST PART");

            #region 1st Part

            // Copy and open the generic block1 part

            LockOpen($"{destPath}{Component1}", $"{sourcePath}{Component1}");

            // Get the active document as a Part and a Model
            PartDoc part = swApp.ActiveDoc as PartDoc;
            ModelDoc2 swModel = swApp.ActiveDoc as ModelDoc2;

            // Get the part feature called exstrusionBase
            Feature Feat = part.FeatureByName("rectSketch");

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
            Feat = part.FeatureByName("Extrude");

            // Replace the current selection
            Feat.Select2(false, -1);

            // Get dimension c of extrusion
            swDim = (Dimension)Feat.Parameter("c");

            // Set the new value of c as the value from the read-in field
            errors = swDim.SetSystemValue3(cDim, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

            // Invoke all changes and rebuild the document, save afterwards
            part.EditRebuild();
            swModel.Save();
            swApp.CloseDoc($"{ destPath}{Component1}");


                #endregion

                MessageBox.Show("Beginning Second Part");

                #region 2nd Part

                double part2a = aDim * 2.0;
                double part2b = bDim / 2.0;
                double part2c = cDim * 3.0;

                // Copy and open the generic block1 part
                LockOpen($"{destPath}{Component2}", $"{sourcePath}{Component2}");

                // Get the active document as a Part and a Model
                part = swApp.ActiveDoc as PartDoc;
                swModel = swApp.ActiveDoc as ModelDoc2;

                // Get the part feature called exstrusionBase
                Feat = part.FeatureByName("rectSketch");

                // Select the feature extrusionBase - Replace current selection (normal click, not ctrl-click)
                Feat.Select2(false, -1);

                // Get dimension a of extrusionBase
                swDim = (Dimension)Feat.Parameter("a");

                // Set the new value of a as the value from the read-in field
                errors = swDim.SetSystemValue3(part2a, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                // Get dimension b of extrusionBase
                swDim = (Dimension)Feat.Parameter("b");

                // Set the new value of b as the value from the read-in field
                errors = swDim.SetSystemValue3(part2b, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                // Get the part feature called extrusion
                Feat = part.FeatureByName("Extrude");

                // Replace the current selection
                Feat.Select2(false, -1);

                // Get dimension c of extrusion
                swDim = (Dimension)Feat.Parameter("c");

                // Set the new value of c as the value from the read-in field
                errors = swDim.SetSystemValue3(part2c, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                // Invoke all changes and rebuild the document, save afterwards
                part.EditRebuild();
                swModel.Save();
                swApp.CloseDoc($"{destPath}{Component2}");

                #endregion

                MessageBox.Show("Beginning 3rd Part");

                #region 3rd Part

                double part3a = aDim * 2;
                double part3b = bDim * 0.5;
                double part3c = cDim / 0.2;

                // Copy and open the generic block1 part
                LockOpen($"{destPath}{Component3}", $"{sourcePath}{Component3}");

                // Get the active document as a Part and a Model
                part = swApp.ActiveDoc as PartDoc;
                swModel = swApp.ActiveDoc as ModelDoc2;

                // Get the part feature called exstrusionBase
                Feat = part.FeatureByName("cylSketch");

                // Select the feature extrusionBase - Replace current selection (normal click, not ctrl-click)
                Feat.Select2(false, -1);

                // Get dimension a of extrusionBase
                swDim = (Dimension)Feat.Parameter("a");

                // Set the new value of a as the value from the read-in field
                errors = swDim.SetSystemValue3(part3a, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                Feat = part.FeatureByName("cylExtrude");

                // Get dimension b of extrusionBase
                swDim = (Dimension)Feat.Parameter("b");

                // Set the new value of b as the value from the read-in field
                errors = swDim.SetSystemValue3(part3b, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);


                // Get the part feature called extrusion
                Feat = part.FeatureByName("holeCut");

                Debug.Print("Beginning to mess with the extruded cut");

                // Replace the current selection
                ExtrudeFeatureData swCutFeatureData = default(ExtrudeFeatureData);
                swCutFeatureData = (ExtrudeFeatureData)Feat.GetDefinition();

                Debug.Print("Made Some Data Objects");
                if (aDim >= 0.25)
                {
                    Debug.Print("Attempting to edit entity");
                    Entity ent = (Entity)part.GetEntityByName("endingFace", (int)swSelectType_e.swSelFACES);
                    if (ent != null)
                    {
                        Debug.Print("Attempting to Select a Face");
                        ent.Select4(false, null);
                    }
                    swCutFeatureData.SetEndCondition(true, (int)swEndConditions_e.swEndCondUpToSelection);
                }
                //else
                //{
                //    Debug.Print("Attempting to Suppress Feature");
                //    Feat.SetSuppression((int)swFeatureSuppressionAction_e.swSuppressFeature);
                //}

                // Invoke all changes and rebuild the document, save afterwards
                part.EditRebuild();
                swModel.Save();
                swApp.CloseDoc($"{destPath}{Component3}");

                //LockOpen($"{destPath}Assembly.SLDASM", $"{sourcePath}Assembly.SLDASM");\
                Debug.Print("Opening the Assembly File");
                if (!File.Exists($"{destPath}Assembly.SLDASM"))
                {
                    File.Copy($"{sourcePath}Assembly.SLDASM",$"{destPath}Assembly.SLDASM");
                }
                swPart = swApp.OpenDoc($"{destPath}Assembly.SLDASM", (int)swDocumentTypes_e.swDocASSEMBLY);
                Debug.Print("Setting Model and Assembly Variables");
                ModelDoc2 model = (ModelDoc2)swApp.ActiveDoc;
                ModelDocExtension modelExt = (ModelDocExtension)model.Extension;
                AssemblyDoc swAssem = (AssemblyDoc)model;
                Component component = default(Component);
                bool isSelected = modelExt.SelectByID2("lockTemplateBlock1-1@Copy of lockTemplateAssembly", "COMPONENT", 0, 0, 0, false, -1, (Callout)null, (int)swSelectOption_e.swSelectOptionDefault);
                bool returnVal = swAssem.ReplaceComponents($"{destPath}{Component1}", "", true, true);
                isSelected = modelExt.SelectByID2("lockTemplateBlock2-1@Copy of lockTemplateAssembly", "COMPONENT", 0, 0, 0, false, -1, (Callout)null, (int)swSelectOption_e.swSelectOptionDefault);
                returnVal = swAssem.ReplaceComponents($"{destPath}{Component2}", "", true, true);
                isSelected = modelExt.SelectByID2("lockTemplateCylinder-1@Copy of lockTemplateAssembly", "COMPONENT", 0, 0, 0, false, -1, (Callout)null, (int)swSelectOption_e.swSelectOptionDefault);
                returnVal = swAssem.ReplaceComponents($"{destPath}{Component3}", "", true, true);

                #endregion

            });
            return;
        }

        private void LockOpen(string destPath, string sourcePath)
        {

            MessageBox.Show($"File Attempting to Open: {destPath}");
            MessageBox.Show($"Copying From: {sourcePath}");

            // Check that the file at the destination path does not already exist
            if (!File.Exists(destPath))
            {
                // If the destination file does not already exist, then copy the template from the source path to the destination path
                File.Copy(sourcePath, destPath);
                MessageBox.Show("Successully Opened");
            }

            try
            {
                // Try to open the copied document
                swPart = swApp.OpenDoc(destPath, (int)swDocumentTypes_e.swDocPART);

            }
            catch (Exception)
            {
                // If an exception is thrown, notify the user that the file was not able to be opened
                MessageBox.Show("File Open Failed");

                // Clear any unused objects - should clean up any COM-related errors that arise after debug close
                Marshal.CleanupUnusedObjectsInCurrentContext();

                // Return to the parent function (GlueButtonClick)
                return;
            }
            return;
        }

        private void lockRead()
        {
            return;
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
        /// Copy Directory and it's subdirectories to a specified path
        /// </summary>
        private static void DirectoryCopy(string sourceDirName, string destDirName, bool copySubDirs)
        {
            // Get the subdirectories for the specified directory.
            DirectoryInfo dir = new DirectoryInfo(sourceDirName);

            if (!dir.Exists)
            {
                throw new DirectoryNotFoundException(
                    "Source directory does not exist or could not be found: "
                    + sourceDirName);
            }

            DirectoryInfo[] dirs = dir.GetDirectories();
            // If the destination directory doesn't exist, create it.
            if (!Directory.Exists(destDirName))
            {
                Directory.CreateDirectory(destDirName);
            }

            // Get the files in the directory and copy them to the new location.
            FileInfo[] files = dir.GetFiles();
            foreach (FileInfo file in files)
            {
                string temppath = Path.Combine(destDirName, file.Name);
                file.CopyTo(temppath, false);
            }

            // If copying subdirectories, copy them and their contents to new location.
            if (copySubDirs)
            {
                foreach (DirectoryInfo subdir in dirs)
                {
                    string temppath = Path.Combine(destDirName, subdir.Name);
                    DirectoryCopy(subdir.FullName, temppath, copySubDirs);
                }
            }
        }

        ///
        #endregion
    }
}
