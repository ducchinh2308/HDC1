using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace HDC1
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnTemplate_Click(object sender, RibbonControlEventArgs e)
        {
            string pathTemplate = @"C:\Users\Administrator\Documents\Custom Office Templates\HuynhDucChinh_Template.dotx"
            cHDC_Function.ApplyTemplate(pathTemplate);
        }
    }
}
