using System;
using System.IO;
using System.Xml;
using SkiaSharp;

namespace PSWordCloud
{
    /// <summary>
    /// Defines the <see cref="Image"/> class.
    /// Keeps track of the various image components needed for drawing a word cloud and generating the SVG output.
    /// </summary>
    internal class Image : IDisposable
    {
        /// <summary>
        /// Gets the viewbox that defines the visible area of the image.
        /// </summary>
        internal SKRect Viewbox { get; }

        /// <summary>
        /// Gets the centre point of the <see cref="Image"/>.
        /// </summary>
        internal SKPoint Centre { get => _centrePoint ??= new SKPoint(Viewbox.MidX, Viewbox.MidY); }

        /// <summary>
        /// Gets the origin point of the <see cref="Image"/>.
        /// </summary>
        internal SKPoint Origin { get => Viewbox.Location; }

        /// <summary>
        /// Gets the aspect ratio of the <see cref="Image"/>.
        /// </summary>
        internal float AspectRatio { get => _aspect ??= Viewbox.Width / Viewbox.Height; }

        /// <summary>
        /// Gets the maximum draw radius of the <see cref="Image"/>.
        /// </summary>
        internal float MaxDrawRadius { get => _maxDrawRadius ??= GetMaxRadius(); }

        /// <summary>
        /// Gets the background color of the <see cref="Image"/>.
        /// </summary>
        /// <value>The average color of all the background pixels.</value>
        internal SKColor BackgroundColor { get; }

        /// <summary>
        /// Gets the <see cref="SKRegion"/> that defines the clipping bounds of the image.
        /// Any items that cross this boundary should not be drawn.
        /// </summary>
        /// <value>A region slightly larger than the <see cref="Viewbox"/>.</value>
        internal SKRegion ClippingRegion { get; }

        /// <summary>
        /// Gets the bounds of the <see cref="ClippingRegion"/>.
        /// </summary>
        internal SKRectI ClippingBounds { get => ClippingRegion.Bounds; }

        /// <summary>
        /// Gets the occupied space of the <see cref="Image"/>.
        /// Each item drawn will update this region.
        /// </summary>
        internal SKRegion OccupiedSpace { get; }

        /// <summary>
        /// Gets the <see cref="SKCanvas"/> that captures the drawing state for the <see cref="Image"/>.
        /// </summary>
        internal SKCanvas Canvas { get; }

        private readonly SKDynamicMemoryWStream _memoryStream;
        private SKPoint? _centrePoint;
        private float? _aspect;
        private float? _maxDrawRadius;
        private readonly bool _allowOverflow;

        /// <summary>
        /// Creates a new instance of the <see cref="Image"/> class, with the specified <paramref name="size"/> and
        /// <paramref name="backgroundColor"/>.
        /// </summary>
        /// <param name="size">The size of the image.</param>
        /// <param name="backgroundColor">The color to fill the background with.</param>
        /// <param name="allowOverflow">If true, defines the acceptable draw bounds as slightly larger than the image, leaving
        /// room for some drawing to occur that overflows the given image bounds.</param>
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

        /// <summary>
        /// Creates a new instance of the <see cref="Image"/> class, using the target <paramref name="backgroundPath"/>
        /// to create a new <see cref="SKBitmap"/> and define the image bounds to match the background bitmap.
        /// </summary>
        /// <param name="backgroundPath">The full path to a background image to use.</param>
        /// <param name="allowOverflow">If true, defines the acceptable draw bounds as slightly larger than the image, leaving
        /// room for some drawing to occur that overflows the given image bounds.</param>
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
                BackgroundColor = Utils.GetAverageColor(background.Pixels);
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

        /// <summary>
        /// Dispose the <see cref="IDisposable"/> managed objects used by this instance.
        /// </summary>
        public void Dispose()
        {
            ClippingRegion.Dispose();
            OccupiedSpace.Dispose();
            Canvas.Dispose();
            _memoryStream.Dispose();
        }
    }
}
