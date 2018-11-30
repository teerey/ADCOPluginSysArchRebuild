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
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace ADCOPlugin
{
    
    /// <summary>
    /// Interaction logic for MyAddinControl.xaml
    /// </summary>
    public partial class MyAddinControl : UserControl
    {
        #region Private Members

        private const string typeGlue = "GLUE";
        private string typeLock;
        public string gluePath = "C:\\Users\\trent\\Documents\\glueTemplate.SLDPRT";
        private const string lockPath = "LOCKPATH";
        private const string mCustomPropertyGlueA = "GlueA";
        private const string mCustomPropertyGlueB = "GlueB";
        private const string mCustomPropertyGlueC = "GlueC";
        public SldWorks swApp;
        ModelDoc2 swPart;
        int fileerror;
        int filewarning;

        #endregion

        //public void GlueOpen()
        //{
        //        try
        //        {
        //            swPart = (ModelDoc2)swApp.OpenDoc6(gluePath, (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", ref fileerror, ref filewarning);

        //        }
        //        catch (Exception)
        //        {

        //            MessageBox.Show(string.Format("File Open Failed {0} {0}", fileerror, filewarning));
        //            return;
        //        }
        //}



        #region Constructor

        /// <summary>
        /// Default constructor
        /// </summary>
        public MyAddinControl()
        {
           InitializeComponent();
        }


        #endregion

        /// <summary>
        /// Fired when the plugin first loads up
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void UserControl_Loaded(object sender, System.Windows.RoutedEventArgs e)
        {
            // By default show the Initial Screen (Former type selection)
            initScreen();

        }

        private void initScreen()
        {
            ThreadHelpers.RunOnUIThread(() =>
            {
                GlueContent.Visibility = System.Windows.Visibility.Hidden;
                InitContent.Visibility = System.Windows.Visibility.Visible;

                
            });
        }

        #region Type-Specific Functions

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GlueApplyButton_Click(object sender, RoutedEventArgs e)
        {

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GlueResetButton_Click(object sender, RoutedEventArgs e)
        {

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GlueBackButton_Click(object sender, RoutedEventArgs e)
        {
            initScreen();
        }

        #endregion

        private void LockButton_Click(object sender, RoutedEventArgs e)
        {

        }

        private void GlueButton_Click(object sender, RoutedEventArgs e)
        {
            GlueScreen();
            //GlueOpen();
            
        }

        private void GlueScreen()
        {
            // If glue is checked, change visibility of init and glue content screens so that glue info is visible
            GlueContent.Visibility = System.Windows.Visibility.Visible;
            InitContent.Visibility = System.Windows.Visibility.Hidden;
            
            
            
        }
    }

}
