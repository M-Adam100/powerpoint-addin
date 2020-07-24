using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Xml;
using System.Net;
using HtmlAgilityPack;
using HtmlDocument = HtmlAgilityPack.HtmlDocument;
using System.Text.RegularExpressions;

namespace PowerPointAddIn1
{
    public partial class CustomePane : UserControl
    {
       
        public CustomePane()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                var client = new WebClient();
                String text = client.DownloadString(maskedTextBox2.Text);

                HtmlDocument htmlDoc = new HtmlDocument();
       
                htmlDoc.LoadHtml(text);
                string result = htmlDoc.DocumentNode.InnerText;
                result = WebUtility.HtmlDecode(result);

                PowerPoint.Application ppApp = Globals.ThisAddIn.Application;
                PowerPoint.SlideRange ppSR = ppApp.ActiveWindow.Selection
                    .SlideRange;
                PowerPoint.Shape ppShap = ppSR.Shapes
                    .AddTextbox(Office.MsoTextOrientation
                    .msoTextOrientationHorizontal, 0, 0, 900, 25);
                ppShap.TextEffect.Text = result;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void maskedTextBox2_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }
    }
}
