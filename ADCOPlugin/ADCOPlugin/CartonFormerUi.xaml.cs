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
            
            GlueContent.Visibility = System.Windows.Visibility.Hidden;
            InitContent.Visibility = System.Windows.Visibility.Visible;

        }

        #region Type-Specific Functions
        /// <summary>
        /// Checks to see if user checked Lock carton type on init screen & carries out all lock carton functionalities if true
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TypeLockCheck_Checked(object sender, RoutedEventArgs e)
        {
            
        }

        /// <summary>
        /// Checks to see if user checked "Glue" carton type on init screen & carries out all glue carton functionalites if true
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TypeGlueCheck_Checked(object sender, RoutedEventArgs e)
        {

            // Check if Glue box is checked on itit screen
            if (TypeGlueCheck.IsChecked.Value)
            {
                // If glue is checked, change visibility of init and glue content screens so that glue info is visible
                GlueContent.Visibility = System.Windows.Visibility.Visible;
                InitContent.Visibility = System.Windows.Visibility.Hidden;

                if (BackButton.IsPressed)
                {
                    initScreen();
                    TypeGlueCheck.IsChecked.Value = 0;
                }

            }
            else
            {
                // If glue is not checked, make sure that init screen is visible and glue content remains hidden
                GlueContent.Visibility = System.Windows.Visibility.Hidden;
                InitContent.Visibility = System.Windows.Visibility.Visible;
            }
        }



        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ResetButton_Click(object sender, RoutedEventArgs e)
        {

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BackButton_Click(object sender, RoutedEventArgs e)
        {

        }

        #endregion
    }

}
