
using System;
using System.IO;
using System.Xml;
using SkiaSharp;

namespace PSWordCloud
{
    internal class Image : IDisposable
    {
        internal SKRect Viewbox { get; }

        internal SKPoint Centre { get => _centrePoint ??= new SKPoint(Viewbox.MidX, Viewbox.MidY); }

        internal SKPoint Origin { get => Viewbox.Location; }

        internal float AspectRatio { get => _aspect ??= Viewbox.Width / Viewbox.Height; }

        internal float MaxDrawRadius { get => _maxDrawRadius ??= GetMaxRadius(); }

        internal SKColor BackgroundColor { get; }

        internal SKRegion ClippingRegion { get; }

        internal SKRectI ClippingBounds { get => ClippingRegion.Bounds; }

        internal SKRegion OccupiedSpace { get; }

        internal SKCanvas Canvas { get; }

        private readonly SKDynamicMemoryWStream _memoryStream;
        private SKPoint? _centrePoint;
        private float? _aspect;
        private float? _maxDrawRadius;
        private readonly bool _allowOverflow;

        internal Image(SKSizeI size, SKColor backgroundColor, bool allowOverflow)
        {
            _allowOverflow = allowOverflow;
            _memoryStream = new SKDynamicMemoryWStream();
            Viewbox = new SKRect(left: 0, top: 0, right: size.Width, bottom: size.Height);
            Canvas = SKSvgCanvas.Create(Viewbox, _memoryStream);
            OccupiedSpace = new SKRegion();
            ClippingRegion = new SKRegion();
            BackgroundColor = backgroundColor;

            SetClippingRegion(Viewbox, _allowOverflow);
            DrawBackground(backgroundColor);
        }

        internal Image(string backgroundPath, bool allowOverflow)
            : this(SKBitmap.Decode(backgroundPath), allowOverflow)
        {
        }

        private Image(SKBitmap background, bool allowOverflow)
            : this(new SKSizeI(background.Width, background.Height), SKColor.Empty, allowOverflow)
        {
            try
            {
                DrawBackground(background);
                BackgroundColor = WCUtils.GetAverageColor(background.Pixels);
            }
            finally
            {
                background.Dispose();
            }
        }

        private void SetClippingRegion(SKRect viewbox, bool allowOverflow)
        {
            if (allowOverflow)
            {
                ClippingRegion.SetRect(
                    SKRectI.Round(SKRect.Inflate(
                        viewbox,
                        viewbox.Width * (Constants.BleedAreaScale - 1),
                        viewbox.Height * (Constants.BleedAreaScale - 1))));
            }
            else
            {
                ClippingRegion.SetRect(SKRectI.Round(viewbox));
            }
        }

        private void DrawBackground(SKColor color)
            => Canvas.Clear(color);

        private void DrawBackground(SKBitmap bitmap)
            => Canvas.DrawBitmap(bitmap, x: 0, y: 0);

        private float GetMaxRadius()
        {
            float radius = SKPoint.Distance(Origin, Centre);

            if (_allowOverflow)
            {
                radius *= Constants.BleedAreaScale;
            }

            return radius;
        }

        private void EnsureSvgViewboxIsSet(XmlDocument xml)
        {
            // Check if the SVG already has a viewbox attribute on the root SVG element.
            // If not, make sure we set it correctly.
            var svgElement = xml.GetElementsByTagName("svg")[0] as XmlElement;
            if (svgElement?.GetAttribute("viewbox") == string.Empty)
            {
                svgElement.SetAttribute(
                    "viewbox",
                    $"{Viewbox.Location.X} {Viewbox.Location.Y} {Viewbox.Width} {Viewbox.Height}");
            }
        }

        private void FinalizeImage()
        {
            // Canvas has to be flushed & disposed for it to write the final </svg> tag to the memory stream.
            Canvas.Flush();
            Canvas.Dispose();
            _memoryStream.Flush();
        }

        /// <summary>
        /// Finalizes the image, disposing of the <see cref="Canvas"/>, and writes the resulting data from a memory
        /// stream to an <see cref="XmlDocument"/>.
        /// </summary>
        /// <remarks>
        /// As part of this method, the XML is examined to ensure it contains a `viewbox` attribute.
        /// If not, that will be added to the XML before it is returned.
        /// </remarks>
        /// <returns>The <see cref="XmlDocument"/> containing the written SVG data.</returns>
        internal XmlDocument GetFinalXml()
        {
            FinalizeImage();

            using SKData data = _memoryStream.DetachAsData();
            using var reader = new StreamReader(data.AsStream(), leaveOpen: false);

            var imageXml = new XmlDocument();
            imageXml.LoadXml(reader.ReadToEnd());
            EnsureSvgViewboxIsSet(imageXml);

            return imageXml;
        }

        /// <summary>
        /// Draws the <paramref name="path"/> on the <see cref="Canvas"/> using the specified <paramref name="brush"/>.
        /// The <see cref="OccupiedSpace"/> region will be updated to include the drawn path.
        /// </summary>
        /// <param name="path">The <see cref="SKPath"/> to draw.</param>
        /// <param name="brush">The <see cref="SKPaint"/> brush to draw with.</param>
        internal void DrawPath(SKPath path, SKPaint brush)
        {
            OccupiedSpace.CombineWithPath(path, SKRegionOperation.Union);
            Canvas.DrawPath(path, brush);
        }

        public void Dispose()
        {
            ClippingRegion.Dispose();
            OccupiedSpace.Dispose();
            Canvas.Dispose();
            _memoryStream.Dispose();
        }
    }
}
