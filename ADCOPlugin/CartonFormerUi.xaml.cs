using System.Windows.Controls;
using System.Diagnostics;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.IO;
using System.Text;
using System.Windows;
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

        #region Private Members

        #endregion

        #region Public Members

        // INCH CONVERSION NUMBER: NO. OF INCHES IN 1 M - Most API functions assume the unit is in meters
        double INCH_CONVERSION = 39.370079;

        // Error/warning handlers - Only to be used if OpenDoc method is updated to newer version
        // int fileerror;
        // int filewarning;

        //Get current windows user for proper default destination
        static string userName = System.Environment.UserName;

        //Default routing for part/assembly files
        //Home folder of all template parts/assemblies
        static string TemplateLib = $@"Z:\ADCO Carton Former Templates";//$@"C:\Users\{userName}\Documents\ADCO Carton Former Templates";

        //Home folder of the archive
        static string ArchiveLib = $@"Z:\ADCO Carton Former Archive";//$@"C:\Users\{userName}\Documents\ADCO Carton Former Archive";

        string TemplatePath = TemplateLib;
        string ArchivePath = ArchiveLib;
        static string lockSrcPathDEFAULT = $@"C:\Users\{userName}\Documents\";

        //Type of carton (glue/lock)
        static string[] formerType = { "GLUE", "LOCK" };

        //Overall component folder names
        static string[] formerElement = { "MANDREL", "CAVITY" };

        #region Template Part/Assembly File Names

        //Part files in the glue mandrel domain
        static string[] glueMandrelParts = {"CARTON TEMPLATE.SLDPRT",//0
                                            "WASHER ECCENT.SLDPRT" ,//1 x2
                                            "CYLINDER PUSHER MOUNT.SLDPRT",//2
                                            "R&D5321-31 SMC MGPM20N-100_MGPRod.SLDPRT",//3
                                            "R&D5321-31 SMC MGPM20N-100_MGPTube.SLDPRT",//4
                                            "EJECT CYLINDER CLAMP PLATE.SLDPRT",//5
                                            "MANDREL CENTER MOUNT BAR.SLDPRT",//6
                                            "MANDREL SPREADER.SLDPRT",//7 x2
                                            "MANDREL SIDE PLATE.SLDPRT",//8
                                            "MANDREL SIDE PLATE 2.SLDPRT",//9
                                            "PUSHER PLATE.SLDPRT",//10
                                            "MANDREL STEM.SLDPRT",//11
                                            "MANDREL MNT.SLDPRT",//12
                                            "ORANGE HANDLE.SLDPRT",//13
                                            "ELBOW.SLDPRT" };//14 x2

        //Part files in the glue forming plate domain
        static string[] glueCavityParts = {"BOLT MOUNT.SLDPRT",//0
                                                "FOLDING PLATE BOTTOM.SLDPRT",//1
                                                "FOLDING PLATE LEFT.SLDPRT",//2
                                                "FOLDING PLATE RIGHT.SLDPRT",//3
                                                "FOLDING PLATE TOP.SLDPRT",//4
                                                "GUIDE RAIL MIRROR.SLDPRT",//5
                                                "GUIDE RAIL.SLDPRT",//6
                                                "MAIN PLATE.SLDPRT",//7
                                                "MINOR PLATE.SLDPRT",//8
                                                "MOUNT LEFT.SLDPRT",//9
                                                "MOUNT RIGHT.SLDPRT",//10
                                                "NUTPLATE LEFT AND RIGHT.SLDPRT",//11
                                                "NUTPLATE TOP.SLDPRT",//12
                                                "NUTPLATE1.SLDPRT",//13
                                                "SIDE GUIDE LEFT.SLDPRT",//14
                                                "SIDE GUIDE RIGHT.SLDPRT",//15
                                                "SPACER.SLDPRT",//16
                                                "SPACER1.SLDPRT",//17
                                                "SPACER2.SLDPRT",//18
                                                "SPREADER BOTTOM.SLDPRT",//19
                                                "SPREADER TOP.SLDPRT",//20
                                                "STOP RAIL.SLDPRT" };//21

        //Assembly files in the glue mandrel domain
        static string[] glueMandrelAssemblies = { "CYLINDER ASSEMBLY.SLDASM",//0
                                                  "MANDREL ASSEMBLY.SLDASM" };//1

        //Assembly files in the forming plate mandrel domain
        static string[] glueCavityAssemblies = { "FORMER PLATE ASSEMBLY.SLDASM" }; //0

        #endregion

        char[] FILENAME = { 'F', 'S', 'U', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0' };

        #region Global Enumeration Declarations

        static int FILE_PART = 0;
        static int FILE_ASSEM = 1;
        static int TYPE_GLUE = 0;
        static int TYPE_LOCK = 1;
        static int COMPONENT_MAN = 0;
        static int COMPONENT_CAV = 1;
        string PartNum;

        #endregion

        #region SolidWorks Object Creations

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
        // Declare a model document extension field
        public ModelDocExtension modelExt;
        // Declare a Drawing Document
        public DrawingDoc swDrawing;

        #endregion

        // Destination path for copied file - will be dynamic/user-inputted in final iteration of the package
        static string glueDestPathDEFAULT = $@"C:\Users\{userName}\Documents\";
        static string lockDestPathDEFAULT = $@"C:\Users\{userName}\Documents\";



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
            TemplatePath = $@"{TemplateLib}\{formerType[TYPE_GLUE]}";
            ArchivePath = $@"{ArchiveLib}\{formerType[TYPE_GLUE]}";
            FILENAME[3] = TYPE_GLUE.ToString()[0];

            // Change display to glue screen
            GlueScreen();
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
        }

        private void GlueOpen(int type)
        {

            if (type == FILE_PART)
            {
                swModel = swApp.OpenDoc($@"{ArchivePath}\{new string(FILENAME)}.SLDPRT", (int)swDocumentTypes_e.swDocPART);
                swPart = (PartDoc)swApp.ActiveDoc;
            }

            if (type == FILE_ASSEM)
            {
                Debug.Print($@"{ArchivePath}\{new string(FILENAME)}.SLDASM");
                swModel = swApp.OpenDoc($@"{ArchivePath}\{new string(FILENAME)}.SLDASM", (int)swDocumentTypes_e.swDocASSEMBLY);
                swAssem = (AssemblyDoc)swApp.ActiveDoc;
                modelExt = (ModelDocExtension)swModel.Extension;
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
            int error = GlueRead();

            while (true)
            {
                if (error == 0 && !File.Exists($@"{ArchivePath}\~app.lock"))
                {
                    using (FileStream fs = File.Create($@"{ArchivePath}\~app.lock")) { }
                    GlueSet();
                    Debug.Print($@"{ArchivePath}\~app.lock");
                    File.Delete($@"{ArchivePath}\~app.lock");
                    Debug.Print("Deleted File");
                    break;
                }
                else
                {
                    MessageBox.Show("Glue archive is currently being altered. Press OK to check again.");
                }
            }
            return;
        }

        /// <summary>
        /// Checks that input strings meet a variety of requirements
        /// </summary>
        /// <returns></returns>
        private int GlueRead()
        {
            MessageBoxImage icon = MessageBoxImage.Warning;
            MessageBoxButton button = MessageBoxButton.OK;

            #region Glue Blank Entry Check

            if (GlueAParam.Text == "" || GlueBParam.Text == "" || GlueCParam.Text == "" || GlueDParam.Text == "" || GlueEParam.Text == "" || Thickness.Text == "")
            {
                MessageBox.Show("You must enter values for all box dimensions.", "", button, icon);
                return (1);
            }

            #endregion

            #region Glue Entry Decimal Formatting Check

            if (GlueAParam.Text.Contains("/") || GlueBParam.Text.Contains("/") || GlueCParam.Text.Contains("/") || GlueDParam.Text.Contains("/") || GlueEParam.Text.Contains("/") || Thickness.Text.Contains("/"))
            {
                MessageBox.Show("Please enter decimal values, not fractional.", "", button, icon);
                return (1);
            }

            #endregion

            #region Glue Dimension Range Check

            if (double.Parse(GlueAParam.Text) < 7 || double.Parse(GlueAParam.Text) > 30)
            {
                MessageBox.Show("Value of A must be between 7\" and 30\"", "", button, icon);
                return (1);
            }

            if (double.Parse(GlueBParam.Text) < 7 || double.Parse(GlueBParam.Text) > 26)
            {
                MessageBox.Show("Value of B must be between 7\" and 26\"", "", button, icon);
                return (1);
            }

            if (double.Parse(GlueCParam.Text) < 9.25 || double.Parse(GlueCParam.Text) > 19)
            {
                MessageBox.Show("Value of C must be between 9.25\" and 19\"", "", button, icon);
                return (1);
            }

            if (double.Parse(GlueDParam.Text) < 5.5 || double.Parse(GlueDParam.Text) > 10)
            {
                MessageBox.Show("Value of D must be between 6\" and 10\"", "", button, icon);
                return (1);
            }

            if (double.Parse(GlueEParam.Text) < 1 || double.Parse(GlueEParam.Text) > 6)
            {
                MessageBox.Show("Value of E must be between 1\" and 6\"", "", button, icon);
                return (1);
            }

            #endregion

            return (0);
        }

        /// <summary>
        /// Adjusts the dimensions of a copied generic model - this is where most of the work is done
        /// </summary>
        private void GlueSet()
        {

            ThreadHelpers.RunOnUIThread(() =>
            {
                string lastNum;
                int errors = 0;
                int warnings = 0;
                bool output;
                string[] checkReturn;

                string[] MANDRELPARTS = new string[(int)glueMandrelParts.GetLength(0)];
                for (int i = 0; i < MANDRELPARTS.Length; i++)
                {
                    MANDRELPARTS[i] = glueMandrelParts[i];
                }

                string[] CAVITYPARTS = new string[(int)glueCavityParts.GetLength(0)];
                for (int i = 0; i < CAVITYPARTS.Length; i++)
                {
                    CAVITYPARTS[i] = glueCavityParts[i];
                }

                string redundantCav = "0000000000000000000000";
                StringBuilder stringBuilder = new StringBuilder(redundantCav, redundantCav.Length);

                // Initialize index of part file in static array

                #region Glue Input Handling

                string aDimStr = GlueAParam.Text;
                string bDimStr = GlueBParam.Text;
                string cDimStr = GlueCParam.Text;
                string dDimStr = GlueDParam.Text;
                string eDimStr = GlueEParam.Text;
                string ThiccDimStr = Thickness.Text;
                string[] dimStrs = { aDimStr, bDimStr, cDimStr, dDimStr, eDimStr, ThiccDimStr };


                double aDim = double.Parse(aDimStr) / INCH_CONVERSION;
                double bDim = double.Parse(bDimStr) / INCH_CONVERSION;
                double cDim = double.Parse(cDimStr) / INCH_CONVERSION;
                double dDim = double.Parse(dDimStr) / INCH_CONVERSION;
                double eDim = double.Parse(eDimStr) / INCH_CONVERSION;
                double ThiccDim = double.Parse(ThiccDimStr) / INCH_CONVERSION;

            #endregion

                #region Glue Part Redimensioning
                    #region Cavity Parts
                    FILENAME[4] = COMPONENT_CAV.ToString()[0];
                    for (int idx = 0; idx < glueCavityParts.Length; idx++)
                    {
                        Debug.Print("GlueSet Cavity:" + idx.ToString());
                        FILENAME[5] = idx.ToString("D2")[0];
                        FILENAME[6] = idx.ToString("D2")[1];
                        checkReturn = GlueRedundant(ArchivePath, glueCavityParts[idx], idx, dimStrs, FILE_PART, COMPONENT_CAV);
                        Debug.Print("Exited GlueRedundant");   
                        stringBuilder[idx] = checkReturn[1][0];
                        Debug.Print("Assigned chars to stringBuilder");
                        redundantCav = stringBuilder.ToString();
                        Debug.Print("Assigned stringBuilder to redundantCav");
                        lastNum = checkReturn[0];
                        Debug.Print("Assigned last checkReturn value to lastNum");
                        for (int i = 0; i < lastNum.Length; i++)
                        {
                            FILENAME[14 - i] = (char)lastNum[i];
                        }
                        CAVITYPARTS[idx] = $"{new string(FILENAME)}.SLDPRT";
                        switch (idx)
                        {
                        case 1:
                            #region Part 1 FOLDING PLATE BOTTOM
                            if (redundantCav[idx] == '0')
                            {
                                File.Copy($@"{TemplatePath}\{glueCavityParts[idx]}",
                                        $@"{ArchivePath}\{new string(FILENAME)}.SLDPRT",
                                        true);
                                GlueOpen(FILE_PART);

                                swFeat = swPart.FeatureByName("Sketch1");
                                swFeat.Select2(false, -1);
                                swDim = (Dimension)swFeat.Parameter("FPB");
                                errors = swDim.SetSystemValue3(cDim + 4 * ThiccDim - 0.15625 / INCH_CONVERSION, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                                swPart.EditRebuild();
                                swModel.Save();
                                swApp.CloseDoc($@"{ArchivePath}\{new string(FILENAME)}.SLDPRT");
                            }

                            break;
                        #endregion

                        case 2:
                            #region Part 2 FOLDING PLATE LEFT
                            if (redundantCav[idx] == '0')
                            {
                                File.Copy($@"{TemplatePath}\{glueCavityParts[idx]}",
                                        $@"{ArchivePath}\{new string(FILENAME)}.SLDPRT",
                                        true);

                                GlueOpen(FILE_PART);

                                swFeat = swPart.FeatureByName("Extrude1");
                                swFeat.Select2(false, -1);
                                swDim = (Dimension)swFeat.Parameter("FPL");
                                errors = swDim.SetSystemValue3(dDim + 2 * ThiccDim - 1 / INCH_CONVERSION, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                                swPart.EditRebuild();
                                swModel.Save();
                                swApp.CloseDoc($@"{ArchivePath}\{new string(FILENAME)}.SLDPRT");
                            }

                            break;
                        #endregion

                        case 3:
                            #region Part 3 FOLDING PLATE RIGHT
                            if (redundantCav[idx] == '0')
                            {
                                File.Copy($@"{TemplatePath}\{glueCavityParts[idx]}",
                                        $@"{ArchivePath}\{new string(FILENAME)}.SLDPRT",
                                        true);

                                GlueOpen(FILE_PART);

                                swFeat = swPart.FeatureByName("Extrude1");
                                swFeat.Select2(false, -1);
                                swDim = (Dimension)swFeat.Parameter("FPR");
                                errors = swDim.SetSystemValue3(dDim + 2 * ThiccDim - 1 / INCH_CONVERSION, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                                swPart.EditRebuild();
                                swModel.Save();
                                swApp.CloseDoc($@"{ArchivePath}\{new string(FILENAME)}.SLDPRT");
                            }

                            break;
                        #endregion

                        case 4:
                            #region Part 4 FOLDING PLATE TOP
                            if (redundantCav[idx] == '0')
                            {
                                File.Copy($@"{TemplatePath}\{glueCavityParts[idx]}",
                                        $@"{ArchivePath}\{new string(FILENAME)}.SLDPRT",
                                        true);

                                GlueOpen(FILE_PART);

                                swFeat = swPart.FeatureByName("Sketch1");
                                swFeat.Select2(false, -1);
                                swDim = (Dimension)swFeat.Parameter("FPT");
                                errors = swDim.SetSystemValue3(cDim + 4 * ThiccDim - 0.15625 / INCH_CONVERSION, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                                swPart.EditRebuild();
                                swModel.Save();
                                swApp.CloseDoc($@"{ArchivePath}\{new string(FILENAME)}.SLDPRT");
                            }

                            break;
                        #endregion

                        case 6:
                            #region Part 6 GUIDE RAIL
                            if (redundantCav[idx] == '0')
                            {
                                File.Copy($@"{TemplatePath}\{glueCavityParts[idx]}",
                                            $@"{ArchivePath}\{new string(FILENAME)}.SLDPRT",
                                            true);

                                GlueOpen(FILE_PART);

                                swFeat = swPart.FeatureByName("Base-Flange1");
                                swFeat.Select2(false, -1);
                                swDim = (Dimension)swFeat.Parameter("GRail");
                                errors = swDim.SetSystemValue3(dDim + 2 * ThiccDim + 9.8125 / INCH_CONVERSION, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                                swPart.EditRebuild();
                                swModel.Save();
                                swApp.CloseDoc($@"{ArchivePath}\{new string(FILENAME)}.SLDPRT");
                            }

                            break;
                        #endregion

                        case 7:
                            #region Part 7 MAIN PLATE
                            if (redundantCav[idx] == '0')
                            {
                                File.Copy($@"{TemplatePath}\{glueCavityParts[idx]}",
                                        $@"{ArchivePath}\{new string(FILENAME)}.SLDPRT",
                                        true);

                                GlueOpen(FILE_PART);

                                swFeat = swPart.FeatureByName("Sketch10");
                                swFeat.Select2(false, -1);
                                swDim = (Dimension)swFeat.Parameter("MainC");
                                errors = swDim.SetSystemValue3(cDim, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                                swDim = (Dimension)swFeat.Parameter("MainD");
                                errors = swDim.SetSystemValue3(dDim, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                                swPart.EditRebuild();
                                swModel.Save();
                                swApp.CloseDoc($@"{ArchivePath}\{new string(FILENAME)}.SLDPRT");
                            }

                            break;
                        #endregion

                        case 9:
                            #region Part 9 MOUNT LEFT
                            if (redundantCav[idx] == '0')
                            {
                                File.Copy($@"{TemplatePath}\{glueCavityParts[idx]}",
                                        $@"{ArchivePath}\{new string(FILENAME)}.SLDPRT",
                                        true);

                                GlueOpen(FILE_PART);
                                swFeat = swPart.FeatureByName("Extrude1");
                                swFeat.Select2(false, -1);
                                swDim = (Dimension)swFeat.Parameter("MountL");
                                errors = swDim.SetSystemValue3(dDim + 2 * ThiccDim + 11.5625 / INCH_CONVERSION, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
                                swPart.EditRebuild();
                                swModel.Save();
                                swApp.CloseDoc($@"{ArchivePath}\{new string(FILENAME)}.SLDPRT");
                            }

                            break;
                        #endregion

                        case 10:
                            #region Part 10 MOUNT RIGHT
                            if (redundantCav[idx] == '0')
                            {
                                File.Copy($@"{TemplatePath}\{glueCavityParts[idx]}",
                                        $@"{ArchivePath}\{new string(FILENAME)}.SLDPRT",
                                        true);

                                GlueOpen(FILE_PART);
                                swFeat = swPart.FeatureByName("Extrude1");
                                swFeat.Select2(false, -1);
                                swDim = (Dimension)swFeat.Parameter("MountR");
                                errors = swDim.SetSystemValue3(dDim + 2 * ThiccDim + 11.5625 / INCH_CONVERSION, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
                                swPart.EditRebuild();
                                swModel.Save();
                                swApp.CloseDoc($@"{ArchivePath}\{new string(FILENAME)}.SLDPRT");
                            }

                            break;
                        #endregion

                        case 12:
                            #region Part 12 NUTPLATE TOP
                            if (redundantCav[idx] == '0')
                            {
                                File.Copy($@"{TemplatePath}\{glueCavityParts[idx]}",
                                        $@"{ArchivePath}\{new string(FILENAME)}.SLDPRT",
                                        true);

                                GlueOpen(FILE_PART);

                                swFeat = swPart.FeatureByName("Extrude1");
                                swFeat.Select2(false, -1);
                                swDim = (Dimension)swFeat.Parameter("NPT");
                                errors = swDim.SetSystemValue3(cDim + 4 * ThiccDim - 2.5 / INCH_CONVERSION, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                                swPart.EditRebuild();
                                swModel.Save();
                                swApp.CloseDoc($@"{ArchivePath}\{new string(FILENAME)}.SLDPRT");
                            }

                            break;
                        #endregion

                        case 19:
                            #region Part 19 SPREADER BOTTOM
                            if (redundantCav[idx] == '0')
                            {
                                File.Copy($@"{TemplatePath}\{glueCavityParts[idx]}",
                                        $@"{ArchivePath}\{new string(FILENAME)}.SLDPRT",
                                        true);

                                GlueOpen(FILE_PART);

                                swFeat = swPart.FeatureByName("Sketch1");
                                swFeat.Select2(false, -1);
                                swDim = (Dimension)swFeat.Parameter("SpreadWidthB");
                                errors = swDim.SetSystemValue3(cDim + 4 * ThiccDim + 10.5 / INCH_CONVERSION, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
                                swPart.EditRebuild();
                                swModel.Save();
                                swApp.CloseDoc($@"{ArchivePath}\{new string(FILENAME)}.SLDPRT");
                            }

                            break;
                        #endregion

                        case 20:
                            #region Part 20 SPREADER TOP
                            if (redundantCav[idx] == '0')
                            {
                                File.Copy($@"{TemplatePath}\{glueCavityParts[idx]}",
                                        $@"{ArchivePath}\{new string(FILENAME)}.SLDPRT",
                                        true);

                                GlueOpen(FILE_PART);

                                swFeat = swPart.FeatureByName("Sketch1");
                                swFeat.Select2(false, -1);
                                swDim = (Dimension)swFeat.Parameter("SpreadWidthT");
                                errors = swDim.SetSystemValue3(cDim + 4 * ThiccDim + 10.5 / INCH_CONVERSION, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
                                swPart.EditRebuild();
                                swModel.Save();
                                swApp.CloseDoc($@"{ArchivePath}\{new string(FILENAME)}.SLDPRT");
                            }

                            break;
                            #endregion

                        case 21:
                            #region Part 21 STOP RAIL
                            if (redundantCav[idx] == '0')
                            {
                                File.Copy($@"{TemplatePath}\{glueCavityParts[idx]}",
                                        $@"{ArchivePath}\{new string(FILENAME)}.SLDPRT",
                                        true);

                                GlueOpen(FILE_PART);

                                swFeat = swPart.FeatureByName("Base-Flange1");
                                swFeat.Select2(false, -1);
                                swDim = (Dimension)swFeat.Parameter("D2");
                                errors = swDim.SetSystemValue3(cDim + 4 * ThiccDim - 4.5 / INCH_CONVERSION, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
                                swPart.EditRebuild();
                                swModel.Save();
                                swApp.CloseDoc($@"{ArchivePath}\{new string(FILENAME)}.SLDPRT");
                            }

                            break;
                            #endregion

                    default:
                                if (!File.Exists($@"{ArchivePath}\{new string(FILENAME)}.SLDPRT"))
                                {
                                    File.Copy($@"{TemplatePath}\{glueCavityParts[idx]}",
                                        $@"{ArchivePath}\{new string(FILENAME)}.SLDPRT",
                                        true);
                                }
                                break;
                        }
                    }
                #endregion

                #region Mandrel Parts
                // Cycle through mandrel parts


                string redundantMan = "000000000000000";
                stringBuilder = new StringBuilder(redundantMan, redundantMan.Length);

                FILENAME[4] = COMPONENT_MAN.ToString()[0];
                    for (int idx = 0; idx < glueMandrelParts.Length; idx++)
                    {
                        Debug.Print("GlueSet:" + idx.ToString());
                        FILENAME[5] = idx.ToString("D2")[0];
                        FILENAME[6] = idx.ToString("D2")[1];
                        checkReturn = GlueRedundant(ArchivePath, glueMandrelParts[idx], idx, dimStrs, FILE_PART, COMPONENT_MAN);
                        stringBuilder[idx] = checkReturn[1][0];
                        redundantMan = stringBuilder.ToString();
                        lastNum = checkReturn[0];
                        for (int i = 0; i < lastNum.Length; i++)
                        {
                            FILENAME[14 - i] = (char)lastNum[i];
                        }
                        MANDRELPARTS[idx] = $"{new string(FILENAME)}.SLDPRT";
                        switch (idx)
                        {
                            case 0:
                                #region Part 0
                                if (redundantMan[idx] == '0')
                                {
                                    Debug.Print("Line 370");
                                    File.Copy($@"{TemplatePath}\{glueMandrelParts[idx]}",
                                        $@"{ArchivePath}\{new string(FILENAME)}.SLDPRT",
                                        true);

                                    Debug.Print("Line 375");
                                    GlueOpen(FILE_PART);
                                    Debug.Print("Line 377");

                                    swFeat = swPart.FeatureByName("Extrude1");
                                    swFeat.Select2(false, -1);
                                    swDim = (Dimension)swFeat.Parameter("Thicc");
                                    errors = swDim.SetSystemValue3(ThiccDim, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                                    swFeat = swPart.FeatureByName("Sketch1");
                                    swFeat.Select2(false, -1);
                                    swDim = (Dimension)swFeat.Parameter("A1");
                                    errors = swDim.SetSystemValue3(dDim, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                                    swDim = (Dimension)swFeat.Parameter("A2");
                                    errors = swDim.SetSystemValue3(dDim + 0.03125 / INCH_CONVERSION, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                                    //swDim = (Dimension)swFeat.Parameter("B1");
                                    //errors = swDim.SetSystemValue3(bDim - 0.125 / INCH_CONVERSION, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                                    swDim = (Dimension)swFeat.Parameter("C1");
                                    errors = swDim.SetSystemValue3(cDim, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                                    swDim = (Dimension)swFeat.Parameter("C2");
                                    errors = swDim.SetSystemValue3(cDim, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                                    swDim = (Dimension)swFeat.Parameter("E1");
                                    errors = swDim.SetSystemValue3(eDim, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                                    swDim = (Dimension)swFeat.Parameter("E2");
                                    errors = swDim.SetSystemValue3(eDim, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                                    swDim = (Dimension)swFeat.Parameter("E3");
                                    errors = swDim.SetSystemValue3(eDim, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                                    swDim = (Dimension)swFeat.Parameter("E4");
                                    errors = swDim.SetSystemValue3(eDim, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
                                

                                    swPart.EditRebuild();
                                    swModel.Save();
                                    swApp.CloseDoc($@"{ArchivePath}\{new string(FILENAME)}.SLDPRT");
                                }
                                break;
                            #endregion

                            case 6:
                                #region Part 6

                                if (redundantMan[6] == '0')
                                {
                                    File.Copy($@"{TemplatePath}\{glueMandrelParts[idx]}",
                                        $@"{ArchivePath}\{new string(FILENAME)}.SLDPRT",
                                        true);

                                    GlueOpen(FILE_PART);

                                    swFeat = swPart.FeatureByName("Extrude1");
                                    swFeat.Select2(false, -1);
                                    swDim = (Dimension)swFeat.Parameter("MountLength");
                                    errors = swDim.SetSystemValue3(cDim - 0.375 * 2.0 / INCH_CONVERSION, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                                    swPart.EditRebuild();
                                    swModel.Save();
                                    swApp.CloseDoc($@"{ArchivePath}\{new string(FILENAME)}.SLDPRT");
                                }
                                break;
                            #endregion

                            case 7:
                                #region Part 7

                                if (redundantMan[7] == '0')
                                {
                                    File.Copy($@"{TemplatePath}\{glueMandrelParts[idx]}",
                                        $@"{ArchivePath}\{new string(FILENAME)}.SLDPRT",
                                        true);

                                    GlueOpen(FILE_PART);

                                    if (cDim <= 7.5 / INCH_CONVERSION)
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
                                    swApp.CloseDoc($@"{ArchivePath}\{new string(FILENAME)}.SLDPRT");
                                }
                                break;
                            #endregion

                            case 8:
                                #region Part 8
                                if (redundantMan[8] == '0')
                                {
                                    File.Copy($@"{TemplatePath}\{glueMandrelParts[idx]}",
                                        $@"{ArchivePath}\{new string(FILENAME)}.SLDPRT",
                                        true);

                                    GlueOpen(FILE_PART);

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
                                    swApp.CloseDoc($@"{ArchivePath}\{new string(FILENAME)}.SLDPRT");
                                }
                                break;
                            #endregion

                            case 9:
                                #region Part 9

                                if (redundantMan[9] == '0')
                                {
                                    File.Copy($@"{TemplatePath}\{glueMandrelParts[idx]}",
                                        $@"{ArchivePath}\{new string(FILENAME)}.SLDPRT",
                                        true);

                                    GlueOpen(FILE_PART);

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
                                    swApp.CloseDoc($@"{ArchivePath}\{new string(FILENAME)}.SLDPRT");
                                }
                                break;
                            #endregion

                            case 10:
                                #region Part 10

                                if (redundantMan[10] == '0')
                                {

                                    File.Copy($@"{TemplatePath}\{glueMandrelParts[idx]}",
                                        $@"{ArchivePath}\{new string(FILENAME)}.SLDPRT",
                                        true);

                                    GlueOpen(FILE_PART);

                                    double PushPlateWidth = dDim - 1.406 * 2.0 / INCH_CONVERSION;

                                    if (PushPlateWidth <= 4.0 / INCH_CONVERSION)
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
                                    swApp.CloseDoc($@"{ArchivePath}\{new string(FILENAME)}.SLDPRT");
                                }
                                break;
                            #endregion

                            default:
                                if (!File.Exists($@"{ArchivePath}\{new string(FILENAME)}.SLDPRT"))
                                {
                                    File.Copy($@"{TemplatePath}\{glueMandrelParts[idx]}",
                                        $@"{ArchivePath}\{new string(FILENAME)}.SLDPRT",
                                        true);
                                }
                                break;
                        }
                    }
                #endregion
                #endregion

                #region Glue Assembly Manipulation

                #region Cavity Assemblies
                FILENAME[4] = COMPONENT_CAV.ToString()[0];
                string[] CAVITYASSEMBLIES = new string[(int)glueCavityAssemblies.GetLength(0)];
                redundantCav = "0";
                stringBuilder = new StringBuilder(redundantCav, redundantCav.Length);
                for (int i = 0; i < CAVITYASSEMBLIES.Length; i++)
                {
                    CAVITYASSEMBLIES[i] = glueCavityAssemblies[i];
                }

                // Manipulate cavity assembly
                for (int idx = 0; idx < glueCavityAssemblies.Length; idx++)
                {
                    Debug.Print($@"Assembly Mainpulation: {idx.ToString()}");
                    FILENAME[5] = idx.ToString("D2")[0];
                    FILENAME[6] = idx.ToString("D2")[1];
                    checkReturn = GlueRedundant(ArchivePath, glueCavityAssemblies[idx], idx, dimStrs, FILE_ASSEM, COMPONENT_CAV);
                    stringBuilder[idx] = checkReturn[1][0];
                    redundantCav = stringBuilder.ToString();
                    lastNum = checkReturn[0];
                    for (int i = 0; i < lastNum.Length; i++)
                    {
                        FILENAME[14 - i] = (char)lastNum[i];
                    }
                    CAVITYASSEMBLIES[idx] = $"{new string(FILENAME)}.SLDASM";
                    Debug.Print(CAVITYASSEMBLIES[idx]);
                    switch (idx)
                    {
                        case 0:
                            if (redundantCav[idx] == '0')
                            {
                                Debug.Print($@"Cavity Switch: {TemplatePath}\{glueCavityAssemblies[idx]}");
                                Debug.Print($@"Cavity Switch: {ArchivePath}\{CAVITYASSEMBLIES[idx]}");
                                File.Copy($@"{TemplatePath}\{glueCavityAssemblies[idx]}",
                                    $@"{ArchivePath}\{CAVITYASSEMBLIES[idx]}",
                                    true);

                                for (int idx2 = 0; idx2 < glueCavityParts.Length; idx2++)
                                {
                                    if (idx2 != 0 && idx2 != 5 && idx2 != 8 && idx2 != 13 && idx2 != 14 && idx2 != 15 && idx2 != 16 && idx2 != 17 && idx2 != 18)
                                    {
                                        output = swApp.ReplaceReferencedDocument($@"{ArchivePath}\{CAVITYASSEMBLIES[idx]}", $@"{TemplatePath}\{glueCavityParts[idx2]}", $@"{ArchivePath}\{CAVITYPARTS[idx2]}");
                                    }
                                }
                            }
                            else
                            {
                                // Carton substitution even if different carton than was first used to generate tooling
                                output = swApp.ReplaceReferencedDocument($@"{ArchivePath}\{CAVITYASSEMBLIES[idx]}", $@"{TemplatePath}\{glueCavityParts[0]}", $@"{ArchivePath}\{CAVITYPARTS[0]}");
                            }
                            break;

                        default:
                            if (!File.Exists($@"{ArchivePath}\{CAVITYASSEMBLIES[idx]}"))
                            {
                                File.Copy($@"{TemplatePath}\{glueCavityAssemblies[idx]}",
                                   $@"{ArchivePath}\{CAVITYASSEMBLIES[idx]}",
                                   true);
                            }
                            output = swApp.ReplaceReferencedDocument($@"{ArchivePath}\{CAVITYASSEMBLIES[idx]}", $@"{TemplatePath}\{glueCavityParts[3]}", $@"{ArchivePath}\{CAVITYPARTS[3]}");
                            output = swApp.ReplaceReferencedDocument($@"{ArchivePath}\{CAVITYASSEMBLIES[idx]}", $@"{TemplatePath}\{glueCavityParts[4]}", $@"{ArchivePath}\{CAVITYPARTS[4]}");
                            break;
                    }
                }

                GlueOpen(FILE_ASSEM);
                output = swModel.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref errors, ref warnings);
                swModel.Rebuild((int)swRebuildOptions_e.swRebuildAll);
                output = swModel.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref errors, ref warnings);
                swApp.CloseDoc($@"{ArchivePath}\{CAVITYASSEMBLIES[glueCavityAssemblies.Length - 1]}");
                GlueOpen(FILE_ASSEM);
                #endregion

                #region Mandrel Assemblies
                FILENAME[4] = COMPONENT_MAN.ToString()[0];
                string[] MANDRELASSEMBLIES = new string[(int)glueMandrelAssemblies.GetLength(0)];
                redundantMan = "00";
                stringBuilder = new StringBuilder(redundantMan, redundantMan.Length);
                for (int i = 0; i < MANDRELASSEMBLIES.Length; i++)
                {
                    MANDRELASSEMBLIES[i] = glueMandrelAssemblies[i];
                }

                // Manipulate mandrel assembly
                for (int idx = 0; idx < glueMandrelAssemblies.Length; idx++)
                {
                    Debug.Print($@"Assembly Mainpulation: {idx.ToString()}");
                    FILENAME[5] = idx.ToString("D2")[0];
                    FILENAME[6] = idx.ToString("D2")[1];
                    checkReturn = GlueRedundant(ArchivePath, glueMandrelAssemblies[idx], idx, dimStrs, FILE_ASSEM, COMPONENT_MAN);
                    stringBuilder[idx] = checkReturn[1][0];
                    redundantMan = stringBuilder.ToString();
                    lastNum = checkReturn[0];
                    for (int i = 0; i < lastNum.Length; i++)
                    {
                        FILENAME[14 - i] = (char)lastNum[i];
                    }
                    MANDRELASSEMBLIES[idx] = $"{new string(FILENAME)}.SLDASM";
                    Debug.Print(MANDRELASSEMBLIES[idx]);
                    switch (idx)
                    {
                        case 1:
                            if (redundantMan[idx] == '0')
                            {

                                File.Copy($@"{TemplatePath}\{glueMandrelAssemblies[idx]}",
                                    $@"{ArchivePath}\{MANDRELASSEMBLIES[idx]}",
                                    true);

                                for (int idx2 = 1; idx2 < glueMandrelParts.Length; idx2++)
                                {
                                    if (idx2 != 3)
                                    {
                                        Debug.Print("If Statement");
                                        Debug.Print($@"{ArchivePath}\{MANDRELASSEMBLIES[idx]}");
                                        Debug.Print($@"{TemplatePath}\{glueMandrelParts[idx2]}");
                                        Debug.Print($@"{ArchivePath}\{MANDRELPARTS[idx2]}");
                                        output = swApp.ReplaceReferencedDocument($@"{ArchivePath}\{MANDRELASSEMBLIES[idx]}", $@"{TemplatePath}\{glueMandrelParts[idx2]}", $@"{ArchivePath}\{MANDRELPARTS[idx2]}");
                                        Debug.Print(idx2.ToString());
                                    }
                                    else if (idx2 == 3)  // Assembly Substitution
                                    {
                                        Debug.Print("Else Statement");
                                        Debug.Print($@"{ArchivePath}\{MANDRELASSEMBLIES[idx]}");
                                        Debug.Print($@"{TemplatePath}\{glueMandrelAssemblies[0]}");
                                        Debug.Print($@"{ArchivePath}\{MANDRELASSEMBLIES[0]}");
                                        output = swApp.ReplaceReferencedDocument($@"{ArchivePath}\{MANDRELASSEMBLIES[idx]}", $@"{TemplatePath}\{glueMandrelAssemblies[0]}", $@"{ArchivePath}\{MANDRELASSEMBLIES[0]}");
                                        idx2++;
                                    }
                                }
                            }
                            else
                            {
                                // Carton substitution even if different carton than was first used to generate tooling
                                output = swApp.ReplaceReferencedDocument($@"{ArchivePath}\{MANDRELASSEMBLIES[idx]}", $@"{TemplatePath}\{glueMandrelParts[0]}", $@"{ArchivePath}\{MANDRELPARTS[0]}");
                            }
                            break;

                        default:
                            if (!File.Exists($@"{ArchivePath}\{MANDRELASSEMBLIES[idx]}"))
                            {
                                File.Copy($@"{TemplatePath}\{glueMandrelAssemblies[idx]}",
                                   $@"{ArchivePath}\{MANDRELASSEMBLIES[idx]}",
                                   true);
                            }
                            output = swApp.ReplaceReferencedDocument($@"{ArchivePath}\{MANDRELASSEMBLIES[idx]}", $@"{TemplatePath}\{glueMandrelParts[3]}", $@"{ArchivePath}\{MANDRELPARTS[3]}");
                            output = swApp.ReplaceReferencedDocument($@"{ArchivePath}\{MANDRELASSEMBLIES[idx]}", $@"{TemplatePath}\{glueMandrelParts[4]}", $@"{ArchivePath}\{MANDRELPARTS[4]}");
                            break;
                    }
                }

                GlueOpen(FILE_ASSEM);
                output = swModel.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref errors, ref warnings);
                swModel.Rebuild((int)swRebuildOptions_e.swRebuildAll);
                output = swModel.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref errors, ref warnings);
                //swApp.CloseDoc($@"{ArchivePath}\{MANDRELASSEMBLIES[glueMandrelAssemblies.Length-1]}");
                #endregion

                #endregion
            });
            return;
        }

        string[] GlueRedundant(string ArchivePath, string PartName, int PartNum, string[] DimStrs, int fileType, int component)
        {
            bool flag = false;
            string line;
            string checkStr;
            string[] values;
            string lastNum = "";
            string redundant = "";

            #region Part File Handling

            if (fileType == FILE_PART)
            {
                Debug.Print("Line 637");
                #region Part Important Dimension Set

                if (component == COMPONENT_MAN)
                {
                    // List of strings defining which dimensions each part is dependent on
                    // The string number corresponds to the part defined in glueMandrelParts string array
                    string[] checkLoc = {  "03456",
                                        "",
                                        "",
                                        "",
                                        "",
                                        "",
                                        "03",
                                        "03",
                                        "04",
                                        "04",
                                        "034",
                                        "",
                                        "",
                                        "",
                                        ""};

                    // Specify a particular dependency string in checkLoc for the current part
                    checkStr = checkLoc[PartNum];
                }
                else
                {
                    string[] checkLoc = {  "",//0
                                        "036",//1
                                        "046",//2
                                        "046",//3
                                        "036",//4
                                        "",//5
                                        "046",//6
                                        "0346",//7
                                        "",//8
                                        "046",//9
                                        "046",//10
                                        "046",//11
                                        "036",//12
                                        "",//13
                                        "",//14
                                        "",//15
                                        "",//16
                                        "",//17
                                        "",//18
                                        "036",//19
                                        "036",//20
                                        "036"};//21

                    // Specify a particular dependency string in checkLoc for the current part
                    checkStr = checkLoc[PartNum];
                }

                

                #endregion


                if (checkStr != "")
                {
                    Debug.Print("Line 661");
                    #region Glue Archive Part CSV Parsing Setup
                    // Rename the file that will be searched for in the archive (search for csv instead of solidworks file)
                    PartName = PartName.Replace(".SLDPRT", " LOG.csv");

                    // Create an object that will be used to read from the csv file
                    StreamReader reader = new StreamReader(File.OpenRead($@"{ArchivePath}\{PartName}"));

                    // Read in csv lists
                    List<string> idx = new List<string>();
                    List<string> A = new List<string>();
                    List<string> B = new List<string>();
                    List<string> C = new List<string>();
                    List<string> D = new List<string>();
                    List<string> E = new List<string>();
                    List<string> t = new List<string>();

                    #endregion


                    #region Glue Archive Part CSV Parsing

                    // If the file isn't blank, do the following:
                    if (!reader.EndOfStream)
                    {
                        // Do the following until the end of the csv file is reached
                        while (!reader.EndOfStream)
                        {
                            // Read in a line from the csv file
                            line = reader.ReadLine();

                            // If the string in each position is not empty, read in each value to the corresponding variable list
                            if (!String.IsNullOrWhiteSpace(line))
                            {
                                values = line.Split(',');
                                idx.Add(values[0]);
                                A.Add(values[1]);
                                B.Add(values[2]);
                                C.Add(values[3]);
                                D.Add(values[4]);
                                E.Add(values[5]);
                                t.Add(values[6]);
                            }
                        }
                        // Close the file
                        reader.Close();

                        // Convert the list objects to standard string arrays
                        string[] listidx = idx.ToArray();
                        string[] listA = A.ToArray();
                        string[] listB = B.ToArray();
                        string[] listC = C.ToArray();
                        string[] listD = D.ToArray();
                        string[] listE = E.ToArray();
                        string[] listt = t.ToArray();

                        // Recombine each string array to form a standardized, 2-D string matrix
                        string[][] strList = new string[][] { listidx, listA, listB, listC, listD, listE, listt };

                        //  2-D matrix is designed as follows:
                        //  Entry history direction ==>>
                        //  ________________________
                        //  | idx0 | idx1  | idx2  | <-- Row 0
                        //  |______|_______|_______|
                        //  | A0   | A1    | A2    | <-- Row 1
                        //  |______|_______|_______|
                        //  | B0   | B1    | B2    | <-- Row 2
                        //  |______|_______|_______|
                        //  | C0   | C1    | C2    | <-- Row 3
                        //  |______|_______|_______|
                        //  | D0   | D1    | D2    | <-- Row 4
                        //  |______|_______|_______|
                        //  | E0   | E1    | E2    | <-- Row 5
                        //  |______|_______|_______|
                        //  | t0   | t1    | t2    | <-- Row 6
                        //  |__.___|__.____|__.____|
                        //    /|\    /|\     /|\
                        //     |      |       |
                        //   Col 0  Col 1    Col 2

                        #endregion

                        #region Glue Archive Redundancy Checking Loops
                        // Each column represents an entry set, so cycle through columns after each row is checked (establish column first)
                        for (int col = 0; col < listA.Length; col++)
                        {
                            flag = false;
                            // For the current column, cycle through and check each row as indexed by the checkStr
                            for (int check = 1; check < checkStr.Length; check++)
                            {
                                // If the value in the row being checked does not match the input dimension set, then mark it with the flag and proceed to the next column
                                if (strList[(int)Char.GetNumericValue(checkStr[check])][col] != DimStrs[(int)Char.GetNumericValue(checkStr[check]) - 1])
                                {
                                    flag = true;
                                    break;
                                }
                            }
                            // If all important values are the same, then flag will remain false and the following code will be executed
                            if (!flag)
                            {
                                lastNum = strList[0][col];
                                redundant = "1";
                                //MessageBox.Show("This part has been made before");
                                break;
                            }
                        }

                        #endregion

                        #region Uniqueness and 1st entry Set Cases
                        // If flag is true after loops are executed, then no previous entry sets match the input, and the new index is the length of the list (or the last index value + 1)
                        if (flag)
                        {
                            lastNum = listA.Length.ToString();
                            redundant = "0";
                            string append = lastNum + "," + DimStrs[0] + "," + DimStrs[1] + "," + DimStrs[2] + "," + DimStrs[3] + "," + DimStrs[4] + "," + DimStrs[5];
                            TextWriter writer = new StreamWriter($@"{ArchivePath}\{PartName}", true);
                            writer.WriteLine(append);
                            writer.Close();
                        }
                        Debug.Print("Line 784");
                        // Return the index of either the matching parts (if identical) or the new index added (if unique)
                        return new string[] { lastNum, redundant };
                    }
                    // If the file is blank, do the following:
                    else
                    {
                        reader.Close();
                        Debug.Print("Line 791");
                        lastNum = "0";
                        redundant = "0";
                        string append = lastNum + "," + DimStrs[0] + "," + DimStrs[1] + "," + DimStrs[2] + "," + DimStrs[3] + "," + DimStrs[4] + "," + DimStrs[5];
                        Debug.Print("Line 795");
                        Debug.Print($@"{ArchivePath}\{PartName}");
                        TextWriter writer = new StreamWriter($@"{ArchivePath}\{PartName}", true);
                        Debug.Print("Line 798");
                        writer.WriteLine(append);
                        writer.Close();
                        Debug.Print("Line 799");
                        return new string[] { lastNum, redundant };
                    }
                    #endregion
                }
                else
                {
                    lastNum = "0";
                    redundant = "2";
                    return new string[] { lastNum, redundant };
                }

            }

            #endregion

            #region Assembly File Handling
            if (fileType == FILE_ASSEM)
            {
                #region Part Important Dimension Set

                if (component == COMPONENT_MAN)
                {
                    // List of strings defining which dimensions each part is dependent on
                    // The string number corresponds to the part defined in glueMandrelParts string array
                    string[] checkLoc = {  "",//0
                                       "034" };//1
                    // Specify a particular dependency string in checkLoc for the current part
                    checkStr = checkLoc[PartNum];
                }
                else
                {
                    string[] checkLoc = {  "",//0
                                       "034" };//1
                    // Specify a particular dependency string in checkLoc for the current part
                    checkStr = checkLoc[PartNum];
                }


                

                if (checkStr != "")
                {

                    PartName = PartName.Replace(".SLDASM", " LOG.csv");

                    // Create an object that will be used to read from the csv file
                    StreamReader reader = new StreamReader(File.OpenRead($@"{ArchivePath}\{PartName}"));

                    // Read in csv lists
                    List<string> idx = new List<string>();
                    List<string> A = new List<string>();
                    List<string> B = new List<string>();
                    List<string> C = new List<string>();
                    List<string> D = new List<string>();
                    List<string> E = new List<string>();
                    List<string> t = new List<string>();

                    #endregion


                    #region Glue Archive Part CSV Parsing

                    // If the file isn't blank, do the following:
                    if (!reader.EndOfStream)
                    {
                        // Do the following until the end of the csv file is reached
                        while (!reader.EndOfStream)
                        {
                            // Read in a line from the csv file
                            line = reader.ReadLine();

                            // If the string in each position is not empty, read in each value to the corresponding variable list
                            if (!String.IsNullOrWhiteSpace(line))
                            {
                                values = line.Split(',');
                                idx.Add(values[0]);
                                A.Add(values[1]);
                                B.Add(values[2]);
                                C.Add(values[3]);
                                D.Add(values[4]);
                                E.Add(values[5]);
                                t.Add(values[6]);
                            }
                        }
                        // Close the file
                        reader.Close();

                        // Convert the list objects to standard string arrays
                        string[] listidx = idx.ToArray();
                        string[] listA = A.ToArray();
                        string[] listB = B.ToArray();
                        string[] listC = C.ToArray();
                        string[] listD = D.ToArray();
                        string[] listE = E.ToArray();
                        string[] listt = t.ToArray();

                        // Recombine each string array to form a standardized, 2-D string matrix
                        string[][] strList = new string[][] { listidx, listA, listB, listC, listD, listE, listt };

                        //  2-D matrix is designed as follows:
                        //  Entry history direction ==>>
                        //  ________________________
                        //  | idx0 | idx1  | idx2  | <-- Row 0
                        //  |______|_______|_______|
                        //  | A0   | A1    | A2    | <-- Row 1
                        //  |______|_______|_______|
                        //  | B0   | B1    | B2    | <-- Row 2
                        //  |______|_______|_______|
                        //  | C0   | C1    | C2    | <-- Row 3
                        //  |______|_______|_______|
                        //  | D0   | D1    | D2    | <-- Row 4
                        //  |______|_______|_______|
                        //  | E0   | E1    | E2    | <-- Row 5
                        //  |______|_______|_______|
                        //  | t0   | t1    | t2    | <-- Row 6
                        //  |__.___|__.____|__.____|
                        //    /|\    /|\     /|\
                        //     |      |       |
                        //   Col 0  Col 1    Col 2

                        #endregion

                        #region Glue Archive Redundancy Checking Loops
                        // Each column represents an entry set, so cycle through columns after each row is checked (establish column first)
                        for (int col = 0; col < listA.Length; col++)
                        {
                            flag = false;
                            // For the current column, cycle through and check each row as indexed by the checkStr
                            for (int check = 1; check < checkStr.Length; check++)
                            {
                                // If the value in the row being checked does not match the input dimension set, then mark it with the flag and proceed to the next column
                                if (strList[(int)Char.GetNumericValue(checkStr[check])][col] != DimStrs[(int)Char.GetNumericValue(checkStr[check]) - 1])
                                {
                                    flag = true;
                                    break;
                                }
                            }
                            // If all important values are the same, then flag will remain false and the following code will be executed
                            if (!flag)
                            {
                                lastNum = strList[0][col];
                                redundant = "1";
                                //MessageBox.Show("This part has been made before");
                                break;
                            }
                        }

                        #endregion

                        #region Uniqueness and 1st entry Set Cases
                        // If flag is true after loops are executed, then no previous entry sets match the input, and the new index is the length of the list (or the last index value + 1)
                        if (flag)
                        {
                            lastNum = listA.Length.ToString();
                            redundant = "0";
                            string append = lastNum + "," + DimStrs[0] + "," + DimStrs[1] + "," + DimStrs[2] + "," + DimStrs[3] + "," + DimStrs[4] + "," + DimStrs[5];
                            TextWriter writer = new StreamWriter($@"{ArchivePath}\{PartName}", true);
                            writer.WriteLine(append);
                            writer.Close();
                        }
                        Debug.Print("Line 1018");
                        // Return the index of either the matching parts (if identical) or the new index added (if unique)
                        return new string[] { lastNum, redundant };
                    }
                    // If the file is blank, do the following:
                    else
                    {
                        reader.Close();
                        Debug.Print("Line 1026");
                        lastNum = "0";
                        redundant = "0";
                        string append = lastNum + "," + DimStrs[0] + "," + DimStrs[1] + "," + DimStrs[2] + "," + DimStrs[3] + "," + DimStrs[4] + "," + DimStrs[5];
                        Debug.Print("Line 795");
                        Debug.Print($@"{ArchivePath}\{PartName}");
                        TextWriter writer = new StreamWriter($@"{ArchivePath}\{PartName}", true);
                        Debug.Print("Line 798");
                        writer.WriteLine(append);
                        writer.Close();
                        Debug.Print("Line 799");
                        return new string[] { lastNum, redundant };
                    }
                    #endregion
                }
                else
                {
                    lastNum = "0";
                    redundant = "2";
                    return new string[] { lastNum, redundant };
                }
            }

            #endregion

            // Exception catch
            return new string[] { "", "" };
        }

        /// <summary>
        /// Fired when the reset button on the glue page is clicked
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GlueResetButton_Click(object sender, RoutedEventArgs e)
        {
            GlueAParam.Text = "";
            GlueBParam.Text = "";
            GlueCParam.Text = "";
            GlueDParam.Text = "";
            GlueEParam.Text = "";
            Thickness.Text = "";
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
            LockSourcePathBox.Text = $@"{lockSrcPathDEFAULT}\lockTemplate";

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
                    File.Copy($"{sourcePath}Assembly.SLDASM", $"{destPath}Assembly.SLDASM");
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

        #endregion
    }
}