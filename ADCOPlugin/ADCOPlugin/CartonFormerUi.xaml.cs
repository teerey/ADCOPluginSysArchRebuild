using System.Windows.Controls;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Windows;
using Microsoft.Win32;
using static AngelSix.SolidDna.SolidWorksEnvironment;
using AngelSix.SolidDna;

namespace ADCOPlugin
{
    /// <summary>
    /// Interaction logic for MyAddinControl.xaml
    /// </summary>
    public partial class MyAddinControl : UserControl
    {
        #region Private Members

        private const string typeGlue = "GLUE";
        private const string typeLock = "LOCK";
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
            InitContent.Visibility = System.Windows.Visibility.Visible;
            GlueContent.Visibility = System.Windows.Visibility.Hidden;
        }

        #region Check Type Selection

        /// <summary>
        /// Checks for a form type selection to alter taskpane content visibility
        /// </summary>

        #endregion

        private void TypeLockCheck_Checked(object sender, RoutedEventArgs e)
        {
            
        }

        private void TypeGlueCheck_Checked(object sender, RoutedEventArgs e)
        {
            GlueContent.Visibility = TypeGlueCheck.IsChecked.Value ? System.Windows.Visibility.Visible : System.Windows.Visibility.Hidden;
            InitContent.Visibility = TypeGlueCheck.IsChecked.Value ? System.Windows.Visibility.Hidden : System.Windows.Visibility.Visible;
        }
    }

}
