using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ESRI.ArcGIS.ArcMapUI;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.DataSourcesFile;
using ESRI.ArcGIS.Display;
using ESRI.ArcGIS.Framework;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.Geometry;
using Font = System.Drawing.Font;
using Path = System.IO.Path;

namespace CacheMake
{
    public partial class Form1 : Form
    {
        private readonly IMxDocument _mxDocument;
        private IWorkspace _pWorkspace;
        private IFeatureWorkspace _pFeatureWorkspace;
        private readonly IWorkspaceFactory _pWorkspaceFactory = new ShapefileWorkspaceFactory();
        private IApplication _pApplication;

        public Form1()
        {
            InitializeComponent();
        }

        public Form1(IMxDocument mxDocument, IApplication app)
        {
            InitializeComponent();
            this._mxDocument = mxDocument;
            this._pApplication = app;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fl = new FolderBrowserDialog();
            if (fl.ShowDialog().Equals(DialogResult.OK))
            {
                textBox1.Text = fl.SelectedPath;
                this._pWorkspace = _pWorkspaceFactory.OpenFromFile(fl.SelectedPath, 0);
                _pFeatureWorkspace = _pWorkspace as IFeatureWorkspace;

                //如果当前地图文档中的图层数大于，则清空当前地图图层
                if (_mxDocument.FocusMap.LayerCount > 0)
                {
                    _mxDocument.FocusMap.ClearLayers();
                }

                //获取所选文件夹内的所有shp文件
                var files = Directory.GetFiles(textBox1.Text).Where(s => s.EndsWith(".shp"));
                foreach (string file in files)
                {
                    Debug.WriteLine(file);
                    IFeatureLayer pFeatureLayer = new FeatureLayer();
                    if (_pFeatureWorkspace != null)
                    {
                        IFeatureClass pFeatureClass =
                            _pFeatureWorkspace.OpenFeatureClass(Path.GetFileNameWithoutExtension(file));
                        pFeatureLayer.FeatureClass = pFeatureClass;
                        pFeatureLayer.Name = pFeatureClass.AliasName;
                    }
                    ILayer pLayer = (ILayer) pFeatureLayer;
                    _mxDocument.FocusMap.AddLayer(pLayer);
                    //将当前视图范围设置为最大的Extent
                    _mxDocument.ActiveView.Extent = _mxDocument.ActiveView.FullExtent;
                    //刷新当前视图
                    _mxDocument.ActiveView.Refresh();
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            IMap _map = _mxDocument.FocusMap;
            // 重新排列图层顺序，前提是图层名称相对固定
            for (int i = 0; i < _map.LayerCount; i++)
            {
                ILayer pLayer = _map.Layer[i];
                /*IFeatureLayer pFeatureLayer = pLayer as IFeatureLayer;
                if (pFeatureLayer.FeatureClass.ShapeType == esriGeometryType.esriGeometryPoint)
                {
                    SetCicleSymbol(pLayer);
                }*/
                if (pLayer.Name.Contains("County_point"))
                {
                    _map.MoveLayer(pLayer, 0);
                }
                else if (pLayer.Name.Contains("ProvincialCapital_point"))
                {
                    _map.MoveLayer(pLayer, 1);
                }
                else if (pLayer.Name.Contains("Town_point"))
                {
                    _map.MoveLayer(pLayer, 2);
                }
                else if (pLayer.Name.Contains("Village_point"))
                {
                    _map.MoveLayer(pLayer, 3);
                }
                else if (pLayer.Name.Contains("CityCapital_point"))
                {
                    _map.MoveLayer(pLayer, 4);
                }
                else if (pLayer.Name.Contains("CountyBoundary_line"))
                {
                    _map.MoveLayer(pLayer, 5);
                }
                else if (pLayer.Name.Contains("ProvinceBoundary_line"))
                {
                    _map.MoveLayer(pLayer, 6);
                }
                else if (pLayer.Name.Contains("CityBoundary_line"))
                {
                    _map.MoveLayer(pLayer, 7);
                }
                else if (pLayer.Name.Contains("highway"))
                {
                    _map.MoveLayer(pLayer, 8);
                }
                else
                {
                    _map.MoveLayer(pLayer, 9);
                }
            }

            IEnumLayer mapLayers = _map.Layers[null, false];
            mapLayers.Reset();
            ILayer enuLayer = mapLayers.Next();
            while (enuLayer != null)
            {
                //LabelLayer(enuLayer, 10, "NAME");
                if (((IFeatureLayer) enuLayer).FeatureClass.ShapeType == esriGeometryType.esriGeometryPoint)
                {
                    SetCicleSymbol(enuLayer);
                    if (enuLayer.Name.Contains("ProvincialCapital_point") ||
                        enuLayer.Name.Contains("CityCapital_point"))
                    {
                        LabelLayer(enuLayer, 10, "NAME", 2500000, 1000000);
                    }
                    if (enuLayer.Name.Contains("County_point"))
                    {
                        LabelLayer(enuLayer, 10, "NAME", 1000000, 250000);
                    }
                    if (enuLayer.Name.Contains("Town_point"))
                    {
                        LabelLayer(enuLayer, 10, "NAME", 250000, 100000);
                    }
                    if (enuLayer.Name.Contains("Village_point"))
                    {
                        LabelLayer(enuLayer, 10, "NAME", 100000, 0);
                    }
                }

                if (((IFeatureLayer) enuLayer).FeatureClass.ShapeType == esriGeometryType.esriGeometryPolyline)
                {
                    if (enuLayer.Name.Contains("highway"))
                    {
                        //MessageBox.Show(((IFeatureLayer) enuLayer).FeatureClass.ShapeType.ToString());
                        IRgbColor pColor = new RgbColor();
                        pColor.Red = 255;
                        pColor.Green = 205;
                        pColor.Blue = 70;
                        SetLineStyle(enuLayer, 4, pColor);
                        LabelHighWay(enuLayer);
                    }
                    if (enuLayer.Name.Contains("CityBoundary_line"))
                    {
                        IRgbColor fillCr = new RgbColor();
                        fillCr.Red = 189;
                        fillCr.Green = 189;
                        fillCr.Blue = 197;
                        SetLineStyle(enuLayer, 1, fillCr);
                    }
                    if (enuLayer.Name.Contains("ProvinceBoundary_line"))
                    {
                        IRgbColor fillCr = new RgbColor();
                        fillCr.Red = 189;
                        fillCr.Green = 189;
                        fillCr.Blue = 197;
                        SetLineStyle(enuLayer, 2, fillCr);
                    }
                    if (enuLayer.Name.Contains("CountyBoundary_line"))
                    {
                        IRgbColor fillCr = new RgbColor();
                        fillCr.Red = 189;
                        fillCr.Green = 189;
                        fillCr.Blue = 197;
                        SetLineStyle(enuLayer);
                    }
                }
                if (((IFeatureLayer) enuLayer).FeatureClass.ShapeType == esriGeometryType.esriGeometryPolygon)
                {
//                    UniqueValueRenderer(enuLayer as IFeatureLayer, new[] { "NAME" });
                    IRgbColor pColor = new RgbColor();
                    pColor.Red = 255;
                    pColor.Green = 243;
                    pColor.Blue = 240;
                    SetStyleHasNoOutLine(enuLayer, esriSimpleFillStyle.esriSFSSolid, pColor);
                }


                enuLayer = mapLayers.Next();
            }
            //if (File.Exists(textBox1.Text + "/result.mxd")) File.Delete(textBox1.Text + "/result.mxd");
            _pApplication.SaveDocument(textBox1.Text + "/result" + GetCurrentTimeUnix() + ".mxd");
            MessageBox.Show(@"保存mxd成功");
            //_pApplication.StatusBar.set_Message(_pApplication.StatusBar.Panes,"MXD文档保存成功");
        }

        //获取当前时间戳
        private long GetCurrentTimeUnix()
        {
            TimeSpan cha = (DateTime.Now - TimeZone.CurrentTimeZone.ToLocalTime(new System.DateTime(1970, 1, 1)));
            long t = (long) cha.TotalSeconds;
            return t;
        }

        private void LabelHighWay(ILayer pLayer)
        {
            IGeoFeatureLayer pGeoFeatureLayer = pLayer as IGeoFeatureLayer;
            if (pGeoFeatureLayer != null)
            {
                IAnnotateLayerPropertiesCollection panAnnotateLayerPropertiesCollection =
                    pGeoFeatureLayer.AnnotationProperties;
                panAnnotateLayerPropertiesCollection.Clear();


                IRgbColor pColor = new RgbColor();
                pColor.Red = 0;
                pColor.Blue = 0;
                pColor.Green = 0;
                IFormattedTextSymbol pTextSymbol = new TextSymbol();
                //ITextSymbol pTextSymbol = new TextSymbol();
                pTextSymbol.Background = CreateBalloonCallout() as ITextBackground;
                pTextSymbol.Color = pColor;
                pTextSymbol.Size = 10;
                pTextSymbol.Direction = esriTextDirection.esriTDHorizontal;
                Font font = new System.Drawing.Font("宋体", 10);
                pTextSymbol.Font = ESRI.ArcGIS.ADF.COMSupport.OLE.GetIFontDispFromFont(font) as stdole.IFontDisp;
                IBasicOverposterLayerProperties properties = new BasicOverposterLayerProperties();
                IFeatureLayer pFeatureLayer = pLayer as IFeatureLayer;

                switch (pFeatureLayer.FeatureClass.ShapeType)
                {
                    case esriGeometryType.esriGeometryPolygon:
                        properties.FeatureType = esriBasicOverposterFeatureType.esriOverposterPolygon;
                        break;
                    case esriGeometryType.esriGeometryPoint:
                        properties.FeatureType = esriBasicOverposterFeatureType.esriOverposterPoint;
                        break;
                    case esriGeometryType.esriGeometryPolyline:
                        properties.FeatureType = esriBasicOverposterFeatureType.esriOverposterPolyline;
                        break;
                }
                ILabelEngineLayerProperties2 properties2 =
                    new LabelEngineLayerProperties() as ILabelEngineLayerProperties2;
                if (properties2 != null)
                {
                    properties2.Expression = "[LXDM]";
                    properties2.Symbol = pTextSymbol;
                    properties2.BasicOverposterLayerProperties = properties;

                    IAnnotateLayerProperties p = properties2 as IAnnotateLayerProperties;
//                    p.AnnotationMaximumScale = maxScale;
//                    p.AnnotationMinimumScale = minScale;
                    panAnnotateLayerPropertiesCollection.Add(p);
                }
            }
            if (pGeoFeatureLayer != null) pGeoFeatureLayer.DisplayAnnotation = true;
            _mxDocument.ActivatedView.Refresh();
        }

        //生成气泡标注
        public IBalloonCallout CreateBalloonCallout( /*double x, double y*/)
        {
            IRgbColor pRgbClr = new RgbColor();
            pRgbClr.Red = 115;
            pRgbClr.Blue = 115;
            pRgbClr.Green = 178;
            ISimpleFillSymbol pSmplFill = new SimpleFillSymbol();
            pSmplFill.Color = pRgbClr;
            pSmplFill.Style = esriSimpleFillStyle.esriSFSSolid;
            IBalloonCallout pBllnCallout = new BalloonCallout();

            pBllnCallout.Style = esriBalloonCalloutStyle.esriBCSRoundedRectangle;
            pBllnCallout.Symbol = pSmplFill;

            pBllnCallout.LeaderTolerance = 15;
            /*IPoint pPoint = new Point();
            pPoint.X =x;
            pPoint.Y =y;
            pBllnCallout.AnchorPoint =pPoint;*/
            return pBllnCallout;
        }

        /// <summary>
        /// 标注图层
        /// </summary>
        /// <param name="pLayer">需要标注的图层</param>
        /// <param name="fontSize">字体大小</param>
        /// <param name="labelField">标注字段</param>
        /// <param name="minScale">显示标注的最小比例尺</param>
        /// <param name="maxScale">显示标注的最大比例尺</param>
        /// 注意：一问最后两个参数是比例尺的分母，所有minScale的值应该比maxSacle的值大
        private void LabelLayer(ILayer pLayer, int fontSize, string labelField, double minScale, double maxScale)
        {
            IGeoFeatureLayer pGeoFeatureLayer = pLayer as IGeoFeatureLayer;
            if (pGeoFeatureLayer != null)
            {
                IAnnotateLayerPropertiesCollection panAnnotateLayerPropertiesCollection =
                    pGeoFeatureLayer.AnnotationProperties;
                panAnnotateLayerPropertiesCollection.Clear();
                IRgbColor pColor = new RgbColor();
                pColor.Red = 0;
                pColor.Blue = 0;
                pColor.Green = 0;
                ITextSymbol pTextSymbol = new TextSymbol();
                pTextSymbol.Color = pColor;
                pTextSymbol.Size = fontSize;
                Font font = new System.Drawing.Font("宋体", fontSize);
                pTextSymbol.Font = ESRI.ArcGIS.ADF.COMSupport.OLE.GetIFontDispFromFont(font) as stdole.IFontDisp;
                IBasicOverposterLayerProperties properties = new BasicOverposterLayerProperties();
                IFeatureLayer pFeatureLayer = pLayer as IFeatureLayer;

                switch (pFeatureLayer.FeatureClass.ShapeType)
                {
                    case esriGeometryType.esriGeometryPolygon:
                        properties.FeatureType = esriBasicOverposterFeatureType.esriOverposterPolygon;
                        break;
                    case esriGeometryType.esriGeometryPoint:
                        properties.FeatureType = esriBasicOverposterFeatureType.esriOverposterPoint;
                        break;
                    case esriGeometryType.esriGeometryPolyline:
                        properties.FeatureType = esriBasicOverposterFeatureType.esriOverposterPolyline;
                        break;
                }
                ILabelEngineLayerProperties2 properties2 =
                    new LabelEngineLayerProperties() as ILabelEngineLayerProperties2;
                if (properties2 != null)
                {
                    properties2.Expression = "[" + labelField + "]";
                    properties2.Symbol = pTextSymbol;
                    properties2.BasicOverposterLayerProperties = properties;

                    IAnnotateLayerProperties p = properties2 as IAnnotateLayerProperties;
                    p.AnnotationMaximumScale = maxScale;
                    p.AnnotationMinimumScale = minScale;
                    panAnnotateLayerPropertiesCollection.Add(p);
                }
            }

            if (pGeoFeatureLayer != null) pGeoFeatureLayer.DisplayAnnotation = true;
            _mxDocument.ActiveView.Refresh();
        }

        private void SetLineStyle(ILayer pLayer, double lineWidth, IRgbColor pColor)
        {
            //MessageBox.Show(pLayer.Name);
            IGeoFeatureLayer pGeoFeatureLayer = pLayer as IGeoFeatureLayer;
            ISimpleRenderer pSimpleRender = new SimpleRenderer();
            ISimpleLineSymbol lineSymbol = new SimpleLineSymbol();
            lineSymbol.Width = lineWidth;
            lineSymbol.Color = pColor;
            pSimpleRender.Symbol = lineSymbol as ISymbol;
            if (pGeoFeatureLayer != null) pGeoFeatureLayer.Renderer = pSimpleRender as IFeatureRenderer;
            _mxDocument.ActiveView.Refresh();
            _mxDocument.CurrentContentsView.Refresh(null);
        }

        private void SetLineStyle(ILayer pLayer)
        {
            IGeoFeatureLayer pGeoFeatureLayer = pLayer as IGeoFeatureLayer;
            ISimpleRenderer pSimpleRender = new SimpleRenderer();
            IRgbColor fillCr = new RgbColor();
            fillCr.Red = 189;
            fillCr.Green = 189;
            fillCr.Blue = 197;
            ISimpleLineSymbol pLineSymbol = new SimpleLineSymbol();
            pLineSymbol.Style = esriSimpleLineStyle.esriSLSDot;
            pLineSymbol.Width = 1;
            pLineSymbol.Color = fillCr;
            pSimpleRender.Symbol = pLineSymbol as ISymbol;
            if (pGeoFeatureLayer != null) pGeoFeatureLayer.Renderer = pSimpleRender as IFeatureRenderer;
            _mxDocument.ActiveView.Refresh();
            _mxDocument.CurrentContentsView.Refresh(null);
        }

        private ICartographicLineSymbol CreateCartographicLineSymbol(IRgbColor rgbColor)
        {
            if (rgbColor == null)
            {
                return null;
            }

            ICartographicLineSymbol cartographicLineSymbol = new CartographicLineSymbol();
            ILineProperties lineProperties = cartographicLineSymbol as ILineProperties;
            lineProperties.Offset = 0;

            // Here's how to do a template for the pattern of marks and gaps
            Double[] hpe = new Double[6];
            hpe[0] = 0;
            hpe[1] = 7;
            hpe[2] = 1;
            hpe[3] = 1;
            hpe[4] = 1;
            hpe[5] = 0;
            /* hpe[0] = 1;
             hpe[1] = 0;
             hpe[2] = 1;
             hpe[3] = 0;
             hpe[4] = 1;
             hpe[5] = 0;
             hpe[6] = 1;
             hpe[7] = 0;
             hpe[11] = 0;*/

            ITemplate template = new Template();
            template.Interval = 1;
            /* for (Int32 i = 0; i < hpe.Length; i = i + 2)
             {
                 template.AddPatternElement(hpe[i], hpe[i + 1]);
             }*/
            for (int i = 0; i < 4; i += 2)
            {
                template.AddPatternElement(i, 0);
            }


            lineProperties.Template = template;

            // Set the basic and cartographic line properties
            cartographicLineSymbol.Width = 1;
            cartographicLineSymbol.Cap = esriLineCapStyle.esriLCSButt;
            cartographicLineSymbol.Join = esriLineJoinStyle.esriLJSRound;
            cartographicLineSymbol.Color = rgbColor;

            return cartographicLineSymbol;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pLayer"></param>
        /// <param name="fillStyle"></param>
        /// <param name="lineColor"></param>
        public void SetStyleHasNoOutLine(ILayer pLayer, esriSimpleFillStyle fillStyle, IRgbColor lineColor)
        {
            IGeoFeatureLayer pGeoFeatureLayer = pLayer as IGeoFeatureLayer;
            ISimpleRenderer pSimpleRender = new SimpleRenderer();
            IRgbColor fillCr = new RgbColor();
            fillCr.Red = 189;
            fillCr.Green = 189;
            fillCr.Blue = 197;
            fillCr.NullColor = true;
            ISimpleLineSymbol pSimpleLineSymbol = new SimpleLineSymbol();
            pSimpleLineSymbol.Width = 1;
            pSimpleLineSymbol.Color = fillCr;
            ISimpleFillSymbol pSimpleFillSysmbol = new SimpleFillSymbol();
            pSimpleFillSysmbol.Style = fillStyle;
            pSimpleFillSysmbol.Outline = pSimpleLineSymbol;
            pSimpleFillSysmbol.Color = lineColor;
            pSimpleRender.Symbol = pSimpleFillSysmbol as ISymbol;
            pGeoFeatureLayer.Renderer = pSimpleRender as IFeatureRenderer;
            _mxDocument.ActiveView.Refresh();
            _mxDocument.CurrentContentsView.Refresh(null);
        }

        /// <summary>
        /// 单一值渲染（单字段）
        /// </summary>
        /// <param name="pLayer"></param>
        /// <param name="RenderField">渲染字段</param>
        /// <param name="FillStyle">填充样式</param>
        /// <param name="valueCount">字段的唯一值个数</param>
        public void CreateUniqueValueRander(ILayer pLayer, string RenderField, esriSimpleFillStyle FillStyle,
            int valueCount)
        {
            IGeoFeatureLayer geoFeatureLayer;
            geoFeatureLayer = pLayer as IGeoFeatureLayer;
            IUniqueValueRenderer uniqueValueRenderer = new UniqueValueRenderer();
            //可以设置多个字段
            uniqueValueRenderer.FieldCount = 1; //0-3个
            uniqueValueRenderer.set_Field(0, RenderField);

            //简单填充符号
            ISimpleFillSymbol simpleFillSymbol = new SimpleFillSymbol();
            simpleFillSymbol.Style = FillStyle;

            IFeatureCursor featureCursor = geoFeatureLayer.FeatureClass.Search(null, false);
            IFeature feature;

            if (featureCursor != null)
            {
                IEnumColors enumColors = CreateAlgorithmicColorRamp();
                int fieldIndex = geoFeatureLayer.FeatureClass.Fields.FindField(RenderField);
                for (int i = 0; i < valueCount; i++)
                {
                    feature = featureCursor.NextFeature();
                    string nameValue = feature.Value[fieldIndex].ToString();
                    simpleFillSymbol = new SimpleFillSymbol();
                    simpleFillSymbol.Color = enumColors.Next();
                    uniqueValueRenderer.AddValue(nameValue, RenderField, simpleFillSymbol as ISymbol);
                }
            }

            geoFeatureLayer.Renderer = uniqueValueRenderer as IFeatureRenderer;
        }


        /// <summary>
        ///  单一值渲染(多字段）
        /// </summary>
        /// <param name="layerName">图层名</param>
        /// <param name="RenderField">多字段名</param>
        /// <param name="FillStyle">样式</param>
        /// <param name="valueCount">每个字段中唯一值的个数</param>
        public void CreateUniqueValueRander(ILayer pLayer, string[] RenderField, esriSimpleFillStyle FillStyle,
            int[] valueCount)
        {
            IGeoFeatureLayer geoFeatureLayer;
            geoFeatureLayer = pLayer as IGeoFeatureLayer;
            IUniqueValueRenderer uniqueValueRenderer = new UniqueValueRenderer();
            //可以设置多个字段
            uniqueValueRenderer.FieldCount = RenderField.Length; //0-3个
            for (int i = 0; i < RenderField.Length; i++)
            {
                uniqueValueRenderer.set_Field(i, RenderField[i]);
            }

            //简单填充符号
            ISimpleFillSymbol simpleFillSymbol = new SimpleFillSymbol();
            simpleFillSymbol.Style = FillStyle;

            IFeatureCursor featureCursor = geoFeatureLayer.FeatureClass.Search(null, false);
            IFeature feature;

            if (featureCursor != null)
            {
                for (int i = 0; i < RenderField.Length; i++)
                {
                    IEnumColors enumColors = CreateAlgorithmicColorRamp();
                    int fieldIndex = geoFeatureLayer.FeatureClass.Fields.FindField(RenderField[i]);
                    for (int j = 0; j < valueCount[i]; j++)
                    {
                        feature = featureCursor.NextFeature();
                        string nameValue = feature.get_Value(fieldIndex).ToString();
                        simpleFillSymbol = new SimpleFillSymbol();
                        simpleFillSymbol.Color = enumColors.Next();
                        uniqueValueRenderer.AddValue(nameValue, RenderField[i], simpleFillSymbol as ISymbol);
                    }
                }
            }
            geoFeatureLayer.Renderer = uniqueValueRenderer as IFeatureRenderer;
        }

        private void SetCicleSymbol(ILayer pLayer)
        {
            //MessageBox.Show(pLayer.Name);
            ISimpleMarkerSymbol pMarkerSymbol = new SimpleMarkerSymbol();
            ISimpleRenderer pSimpleRender = new SimpleRenderer();
            pMarkerSymbol.Style = esriSimpleMarkerStyle.esriSMSCircle;
            pMarkerSymbol.Outline = false;
            IColor pColor = new RgbColor();
            pColor.NullColor = true;
            pMarkerSymbol.Color = pColor;

            pSimpleRender.Symbol = pMarkerSymbol as ISymbol;

            ((IGeoFeatureLayer) pLayer).Renderer = pSimpleRender as IFeatureRenderer;
            //刷新视图
            _mxDocument.ActiveView.Refresh();
            //刷新左侧TOC目录树，如果不刷新，改变了点样式之后点的符号依旧不会变化
            _mxDocument.CurrentContentsView.Refresh(null);
        }

        private IEnumColors CreateAlgorithmicColorRamp()
        {
            IRandomColorRamp pRandomColorRamp = new RandomColorRamp();
            pRandomColorRamp.StartHue = 0;
            pRandomColorRamp.EndHue = 120;
            pRandomColorRamp.MinValue = 0;
            pRandomColorRamp.MaxValue = 90;
            pRandomColorRamp.MinSaturation = 0;
            pRandomColorRamp.MaxSaturation = 45;
            pRandomColorRamp.Size = 20;
            pRandomColorRamp.UseSeed = true;
            pRandomColorRamp.Seed = 40;
            bool bture = true;
            pRandomColorRamp.CreateRamp(out bture);
            IEnumColors pEnuColors = pRandomColorRamp.Colors;
            return pEnuColors;
        }

        private void UniqueValueRenderer(IFeatureLayer pFeatLyr, string[] sFieldName)
        {
            IUniqueValueRenderer pUniqueValueRender;

            IColor pNextUniqueColor;

            IEnumColors pEnumRamp;

            ITable pTable;

            IRow pNextRow;

            ICursor pCursor;

            IQueryFilter pQueryFilter;

            IRandomColorRamp pRandColorRamp = new RandomColorRamp();

            pRandColorRamp.StartHue = 0;

            pRandColorRamp.MinValue = 0;

            pRandColorRamp.MinSaturation = 15;

            pRandColorRamp.EndHue = 360;

            pRandColorRamp.MaxValue = 100;

            pRandColorRamp.MaxSaturation = 30;

            IQueryFilter pQueryFilter1 = new QueryFilter();

            pRandColorRamp.Size = pFeatLyr.FeatureClass.FeatureCount(pQueryFilter1);

            bool bSuccess = false;

            pRandColorRamp.CreateRamp(out bSuccess);
            if (sFieldName.Length == 1)
            {
                IFeatureLayer pFLayer = pFeatLyr as IFeatureLayer;

                IGeoFeatureLayer geoLayer = pFeatLyr as IGeoFeatureLayer;

                IFeatureClass fcls = pFLayer.FeatureClass;

                IQueryFilter pqf = new QueryFilter();

                IFeatureCursor fCursor = fcls.Search(pqf, false);

                IRandomColorRamp rx = new RandomColorRamp();

                rx.MinSaturation = 15;

                rx.MaxSaturation = 30;

                rx.MinValue = 85;

                rx.MaxValue = 100;

                rx.StartHue = 0;

                rx.EndHue = 360;

                rx.Size = 100;

                bool ok;
                ;

                rx.CreateRamp(out ok);

                IEnumColors RColors = rx.Colors;

                RColors.Reset();

                IUniqueValueRenderer pRender = new UniqueValueRenderer();

                pRender.FieldCount = 1;

                pRender.set_Field(0, sFieldName[0]);

                IFeature pFeat = fCursor.NextFeature();

                int index = pFeat.Fields.FindField(sFieldName[0]);

                while (pFeat != null)
                {
                    ISimpleFillSymbol symd = new SimpleFillSymbol();

                    symd.Style = esriSimpleFillStyle.esriSFSSolid;

                    symd.Outline.Width = 1;

                    symd.Color = RColors.Next();

                    string valuestr = pFeat.get_Value(index).ToString();

                    pRender.AddValue(valuestr, valuestr, symd as ISymbol);

                    pFeat = fCursor.NextFeature();
                }

                geoLayer.Renderer = pRender as IFeatureRenderer;
                _mxDocument.ActiveView.Refresh();
            }
            if (sFieldName.Length == 2)
            {
                string sFieldName1 = sFieldName[0];

                string sFieldName2 = sFieldName[1];

                IGeoFeatureLayer pGeoFeatureL = (IGeoFeatureLayer) pFeatLyr;

                pUniqueValueRender = new UniqueValueRenderer();

                pTable = (ITable) pGeoFeatureL;

                int pFieldNumber = pTable.FindField(sFieldName1);

                int pFieldNumber2 = pTable.FindField(sFieldName2);

                pUniqueValueRender.FieldCount = 2;

                pUniqueValueRender.set_Field(0, sFieldName1);

                pUniqueValueRender.set_Field(1, sFieldName2);

                pEnumRamp = pRandColorRamp.Colors;

                pNextUniqueColor = null;

                pQueryFilter = new QueryFilter();

                pQueryFilter.AddField(sFieldName1);

                pQueryFilter.AddField(sFieldName2);

                pCursor = pTable.Search(pQueryFilter, true);

                pNextRow = pCursor.NextRow();

                string codeValue;

                while (pNextRow != null)
                {
                    codeValue = pNextRow.get_Value(pFieldNumber).ToString() + pUniqueValueRender.FieldDelimiter +
                                pNextRow.get_Value(pFieldNumber2).ToString();

                    pNextUniqueColor = pEnumRamp.Next();

                    if (pNextUniqueColor == null)
                    {
                        pEnumRamp.Reset();

                        pNextUniqueColor = pEnumRamp.Next();
                    }

                    IFillSymbol pFillSymbol;

                    ILineSymbol pLineSymbol;

                    IMarkerSymbol pMarkerSymbol;

                    switch (pGeoFeatureL.FeatureClass.ShapeType)
                    {
                        case esriGeometryType.esriGeometryPolygon:
                        {
                            pFillSymbol = new SimpleFillSymbol();

                            pFillSymbol.Color = pNextUniqueColor;

                            pUniqueValueRender.AddValue(codeValue, sFieldName1 + " " + sFieldName2,
                                (ISymbol) pFillSymbol);

                            break;
                        }

                        case esriGeometryType.esriGeometryPolyline:
                        {
                            pLineSymbol = new SimpleLineSymbol();

                            pLineSymbol.Color = pNextUniqueColor;

                            pUniqueValueRender.AddValue(codeValue, sFieldName1 + " " + sFieldName2,
                                (ISymbol) pLineSymbol);

                            break;
                        }

                        case esriGeometryType.esriGeometryPoint:
                        {
                            pMarkerSymbol = new SimpleMarkerSymbol();

                            pMarkerSymbol.Color = pNextUniqueColor;

                            pUniqueValueRender.AddValue(codeValue, sFieldName1 + " " + sFieldName2,
                                (ISymbol) pMarkerSymbol);

                            break;
                        }
                    }

                    pNextRow = pCursor.NextRow();
                }

                pGeoFeatureL.Renderer = (IFeatureRenderer) pUniqueValueRender;

                _mxDocument.ActiveView.Refresh();
            }

            else if (sFieldName.Length == 3)
            {
                string sFieldName1 = sFieldName[0];

                string sFieldName2 = sFieldName[1];

                string sFieldName3 = sFieldName[2];

                IGeoFeatureLayer pGeoFeatureL = (IGeoFeatureLayer) pFeatLyr;

                pUniqueValueRender = new UniqueValueRenderer();

                pTable = (ITable) pGeoFeatureL;

                int pFieldNumber = pTable.FindField(sFieldName1);

                int pFieldNumber2 = pTable.FindField(sFieldName2);

                int pFieldNumber3 = pTable.FindField(sFieldName3);

                pUniqueValueRender.FieldCount = 3;

                pUniqueValueRender.set_Field(0, sFieldName1);

                pUniqueValueRender.set_Field(1, sFieldName2);

                pUniqueValueRender.set_Field(2, sFieldName3);

                pEnumRamp = pRandColorRamp.Colors;

                pNextUniqueColor = null;

                pQueryFilter = new QueryFilter();

                pQueryFilter.AddField(sFieldName1);

                pQueryFilter.AddField(sFieldName2);

                pQueryFilter.AddField(sFieldName3);

                pCursor = pTable.Search(pQueryFilter, true);

                pNextRow = pCursor.NextRow();

                string codeValue;

                while (pNextRow != null)
                {
                    codeValue = pNextRow.get_Value(pFieldNumber).ToString() + pUniqueValueRender.FieldDelimiter +
                                pNextRow.get_Value(pFieldNumber2).ToString() + pUniqueValueRender.FieldDelimiter +
                                pNextRow.get_Value(pFieldNumber3).ToString();

                    pNextUniqueColor = pEnumRamp.Next();

                    if (pNextUniqueColor == null)
                    {
                        pEnumRamp.Reset();

                        pNextUniqueColor = pEnumRamp.Next();
                    }

                    IFillSymbol pFillSymbol;

                    ILineSymbol pLineSymbol;

                    IMarkerSymbol pMarkerSymbol;

                    switch (pGeoFeatureL.FeatureClass.ShapeType)
                    {
                        case esriGeometryType.esriGeometryPolygon:
                        {
                            pFillSymbol = new SimpleFillSymbol();

                            pFillSymbol.Color = pNextUniqueColor;

                            pUniqueValueRender.AddValue(codeValue, sFieldName1 + " " + sFieldName2 + "" + sFieldName3,
                                (ISymbol) pFillSymbol);

                            break;
                        }

                        case esriGeometryType.esriGeometryPolyline:
                        {
                            pLineSymbol = new SimpleLineSymbol();

                            pLineSymbol.Color = pNextUniqueColor;

                            pUniqueValueRender.AddValue(codeValue, sFieldName1 + " " + sFieldName2 + "" + sFieldName3,
                                (ISymbol) pLineSymbol);

                            break;
                        }

                        case esriGeometryType.esriGeometryPoint:
                        {
                            pMarkerSymbol = new SimpleMarkerSymbol();

                            pMarkerSymbol.Color = pNextUniqueColor;

                            pUniqueValueRender.AddValue(codeValue, sFieldName1 + " " + sFieldName2 + "" + sFieldName3,
                                (ISymbol) pMarkerSymbol);

                            break;
                        }
                    }

                    pNextRow = pCursor.NextRow();
                }

                pGeoFeatureL.Renderer = (IFeatureRenderer) pUniqueValueRender;

                _mxDocument.ActiveView.Refresh();
            }
        }
    }
}