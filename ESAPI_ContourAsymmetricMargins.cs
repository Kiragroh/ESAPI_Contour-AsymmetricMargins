using System;
using System.Linq;
using System.Text;
using System.Windows;
using System.IO;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using System.Windows.Controls;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Threading.Tasks;

//change version for possible new approval in clinical mode
[assembly: AssemblyVersion("2.1")]
[assembly: ESAPIScript(IsWriteable = true)]

namespace VMS.TPS
{

    public class Script
    {
        const string SCRIPT_NAME = "zHK-Builder";
                
        public Script()
        {

        }

        
        public void Execute(ScriptContext context /*, System.Windows.Window window, ScriptEnvironment environment*/)

        {

            if (context.Patient == null || context.StructureSet == null)

            {

                MessageBox.Show("Please load a patient, 3D image, and structure set before running this script.", SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);

                return;

            }
            
            string msg1 = "";

            string msg2 = "";
            
            StructureSet ss = context.StructureSet;

            
            // get list of structures for loaded plan

            var listStructures = context.StructureSet.Structures;          
            
            // define PTV (selected)

            //Structure ptv = listStructures.Where(x => !x.IsEmpty && x.Id.ToUpper().Contains("PTV1 REKTM")).FirstOrDefault();

            //Structure ptv = listStructures.Where(x => x.Id == context.PlanSetup.TargetVolumeID).FirstOrDefault();

            var ptv = SelectStructureWindow.SelectStructure(ss);

            if (ptv == null) return;

            else msg2 += string.Format("'{0}' found \n", ptv.Id);

               
            context.Patient.BeginModifications();   // enable writing with this script.

            //============================

            // FIND OAR

            //============================

            
            // find Rectum

            Structure rectum = listStructures.Where(x => !x.IsEmpty && (x.Id.ToUpper().Contains("REKTUM") || x.Id.ToUpper().Contains("RECTUM")) && x.DicomType.Equals("ORGAN") & !x.Id.ToUpper().Contains("HK")).FirstOrDefault();

            if (rectum == null) msg1 += string.Format("Not found: '{0}' \n", "Rektum (OAR)");

            else msg2 += string.Format("'{0}' found \n", rectum.Id);



            // find Blase

            Structure blase = listStructures.Where(x => !x.IsEmpty && x.Id.ToUpper().Contains("BLASE") && x.DicomType.Equals("ORGAN") & !x.Id.ToUpper().Contains("HK")).FirstOrDefault();

            if (blase == null) msg1 += string.Format("Not found: '{0}' \n", "Blase (OAR)");

            else msg2 += string.Format("'{0}' found \n", blase.Id);





            MessageBox.Show(msg2 + "\n" + msg1, SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Information);



            //============================

            // GENERATE HelpStructures

            //============================

            int RectumHkCount = 1;

            int BlaseHkCount = 1;

            foreach (Structure scan in listStructures)

            {
                if (scan.Id.ToUpper().Contains("ZHK REKTUM")) RectumHkCount++;
                if (scan.Id.ToUpper().Contains("ZHK BLASE")) BlaseHkCount++;
            }
            
            // HK Rektum

            if (rectum != null)

            {

                Structure hk_rectum = ss.AddStructure("CONTROL", "zHK Rektum_" + RectumHkCount);

                hk_rectum.SegmentVolume = rectum.Margin(8.0);

                hk_rectum.SegmentVolume = hk_rectum.Sub(ptv);

            }



            // HK Blase (german for bladder)

            if (blase != null)

            {

                double x1 = 5;
                double y1 = 10;
                double z1 = 15;
                double x2 = 5;
                double y2 = 20;
                double z2 = 8;
                AxisAlignedMargins margins = new AxisAlignedMargins(StructureMarginGeometry.Outer, x1, y1, z1, x2, y2, z2);
                Structure hk_blase = ss.AddStructure("CONTROL", "zHK Blase_" + BlaseHkCount);

                hk_blase.SegmentVolume = blase.AsymmetricMargin(margins);

                hk_blase.SegmentVolume = hk_blase.Sub(ptv);

            }
                        
            //============================

            // remove structures unneccesary for optimization

            //============================
            //ss.RemoveStructure(ptv_3mm);
            //ss.RemoveStructure(ptv_minus);
            //ss.RemoveStructure(Iso95);
           
        }

    }

    class SelectStructureWindow : Window
    {
        public static Structure SelectStructure(StructureSet ss)
        {
            m_w = new Window();            
            m_w.WindowStartupLocation = WindowStartupLocation.Manual;
            m_w.Left = 500;
            m_w.Top = 150;
            m_w.Width = 300;
            m_w.Height = 350;
            
            m_w.Title = "Choose Target:";

            var grid = new Grid();

            m_w.Content = grid;

            var list = new ListBox();

            foreach (var s in ss.Structures.Where(s => !s.IsEmpty))

            {

                var tempStruct = s.ToString();

                if (tempStruct.ToUpper().Contains("PTV") || tempStruct.ToUpper().Contains("ZHK") || tempStruct.ToUpper().Contains("SIB") || tempStruct.ToUpper().Contains("CTV") || tempStruct.ToUpper().Contains("GTV"))

                    if (tempStruct.Contains(":"))
                    {
                        int index = tempStruct.IndexOf(":");
                        tempStruct = tempStruct.Substring(0, index);
                    }
                list.Items.Add(s);            
            }

            list.VerticalAlignment = VerticalAlignment.Top;
            list.Margin = new Thickness(10, 10, 10, 55);
            grid.Children.Add(list);
            var button = new Button();
            button.Content = "OK";
            button.Height = 40;
            button.VerticalAlignment = VerticalAlignment.Bottom;
            button.Margin = new Thickness(10, 10, 10, 10);
            button.Click += button_Click;
            grid.Children.Add(button);

            if (m_w.ShowDialog() == true)
            {
                return (Structure)list.SelectedItem;
            }
            return null;
        }
               
        static Window m_w = null;

        static void button_Click(object sender, RoutedEventArgs e)
        {
            m_w.DialogResult = true;
            m_w.Close();
        }

    }

}
