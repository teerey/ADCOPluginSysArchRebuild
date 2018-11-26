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

            // Check to see if a former type is selected
            SolidWorksEnvironment.Application.ActiveModelInformationChanged += Application_ActiveModelInformationChanged;
            checkFormSelect();

        }

        private void Application_ActiveModelInformationChanged(Model obj)
        {
            checkFormSelect();
        }

        #region Check Type Selection

        /// <summary>
        /// Checks for a form type selection to alter taskpane content visibility
        /// </summary>

        private void checkFormSelect()
        {
            AngelSix.SolidDna.ThreadHelpers.RunOnUIThread(() =>
            {
                if (TypeGlueCheck.IsChecked.Value)
                {
                    InitContent.Visibility = System.Windows.Visibility.Hidden;
                    GlueContent.Visibility = System.Windows.Visibility.Visible;
                    return;
                }
            });
        }
        

        #endregion
    }

}
