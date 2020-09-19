
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

        internal SKRegion ClippingRegion { get; }

        internal SKRectI ClippingBounds { get => ClippingRegion.Bounds; }

        internal SKRegion OccupiedSpace { get; }

        internal SKCanvas Canvas { get; }

        private readonly SKDynamicMemoryWStream _memoryStream;
        private SKPoint? _centrePoint;
        private float? _aspect;

        internal Image(SKRect viewbox, bool allowOverflow)
        {
            _memoryStream = new SKDynamicMemoryWStream();
            Viewbox = viewbox;
            Canvas = SKSvgCanvas.Create(Viewbox, _memoryStream);
            OccupiedSpace = new SKRegion();
            ClippingRegion = new SKRegion();

            SetClippingRegion(Viewbox, allowOverflow);
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

        private void FinalizeImage()
        {
            // Canvas has to be flushed & disposed for it to write the final </svg> tag to the memory stream.
            Canvas.Flush();
            Canvas.Dispose();
            _memoryStream.Flush();
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
