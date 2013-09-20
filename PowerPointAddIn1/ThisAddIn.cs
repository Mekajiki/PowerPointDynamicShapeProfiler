using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Diagnostics;
using System.Reflection;
using Tracing = System.Diagnostics.Tracing;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;


namespace PowerPointAddIn1
{
    public partial class ThisAddIn
    {
        void Application_PresentationNewSlide(PowerPoint.Slide Sld)
        {
          
            //PowerPoint.Shape textBox = Sld.Shapes.AddTextbox(
            //Office.MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 500, 50);
            //textBox.TextFrame.TextRange.InsertAfter("This text was added by using code.");
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.SlideSelectionChanged += Application_PresentationNewSlide;

        }

        private void Application_PresentationNewSlide(PowerPoint.SlideRange SldRange)
        {
            PowerPoint.Slide Sld = this.Application.ActiveWindow.View.Slide;
            foreach (PowerPoint.Shape shape in Sld.Shapes)
            {
                Debug.WriteLine("id:" + shape.GetHashCode());

                Debug.WriteLine("isChild?:" + shape.Child);

                Debug.WriteLine("hasChart?:" + shape.HasChart);

                PrintTypeProfile(shape);
            }
        }

        private static void PrintTypeProfile(Object shape)
        {
            Type type = shape.GetType();
            Debug.WriteLine("GetType:" + type);
            String typeName = Microsoft.VisualBasic.Information.TypeName(shape);
            Debug.WriteLine("VBTypeName:" + typeName);
            
            Debug.Write("Interfaces:");
            foreach (Type i in type.GetInterfaces())
            {
                Debug.Write(i + ", ");
            }
            Debug.WriteLine("");

            Debug.Write("Methods:");
            foreach (System.Reflection.MethodInfo info in shape.GetType().GetMethods())
            {
                Debug.Write(info.Name + ", ");
            }
            Debug.WriteLine("\n\n\n\n");
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO で生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// このメソッドの内容をコード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
