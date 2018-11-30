using System.Windows.Controls;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Windows;
using Microsoft.Win32;
using static AngelSix.SolidDna.SolidWorksEnvironment;
using AngelSix.SolidDna;
using System;

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
        private string gluePath;
        private const string lockPath = "LOCKPATH";
        private const string mCustomPropertyGlueA = "GlueA";
        private const string mCustomPropertyGlueB = "GlueB";
        private const string mCustomPropertyGlueC = "GlueC";

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
        }

        private void GlueScreen()
        {
            // If glue is checked, change visibility of init and glue content screens so that glue info is visible
            GlueContent.Visibility = System.Windows.Visibility.Visible;
            InitContent.Visibility = System.Windows.Visibility.Hidden;
        }
    }

}
