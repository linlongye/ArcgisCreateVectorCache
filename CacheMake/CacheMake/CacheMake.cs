using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.Framework;

namespace CacheMake
{
    public class CacheMake : ESRI.ArcGIS.Desktop.AddIns.Button
    {
        public CacheMake()
        {
        }

        protected override void OnClick()
        {
            //
            //  TODO: Sample code showing how to access button host
            //
            IApplication application = ArcMap.Application;

            new Form1(ArcMap.Document, application).ShowDialog();
            
        }

        protected override void OnUpdate()
        {
            Enabled = ArcMap.Application != null;
        }
    }
}