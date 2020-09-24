

using System;
using System.Collections.Generic;
using System.Numerics;
using SkiaSharp;

namespace PSWordCloud
{
    /// <summary>
    /// Defines an ellipse comprised of a collection of points.
    /// </summary>
    internal class Ellipse
    {
        /// <summary>
        /// Gets the list of points that comprises this <see cref="Ellipse"/>.
        /// </summary>
        internal IReadOnlyList<SKPoint> Points { get => _points; }

        /// <summary>
        /// Gets the point at the centre of the ellipse.
        /// </summary>
        internal SKPoint Centre { get; }

        /// <summary>
        /// Gets the ratio between the two axes of the ellipse.
        /// </summary>
        internal float AspectRatio { get; }

        /// <summary>
        /// Gets the radius value used to construct the ellipse.
        /// </summary>
        /// <value>If the <see cref="AspectRatio"/> is less than 1, this represents the semi-major axis. Otherwise, this
        /// represents the semi-minor axis.</value>
        internal float Radius { get; }

        /// <summary>
        /// Gets the minimum distance between each point on the ellipse.
        /// </summary>
        internal float PointDistance { get;  }

        private readonly List<SKPoint> _points;
        private readonly LockingRandom _safeRandom;

        /// <summary>
        /// Creates a new instance of the <see cref="Ellipse"/> class.
        /// </summary>
        /// <param name="radius">One semi-axis of the ellipse.</param>
        /// <param name="radialStep">The typical distance between two points on the ellipse.</param>
        /// <param name="image">The image which is used to determine the size of the other axis of the ellipse.</param>
        /// <param name="randomSeed">A random seed value.</param>
        internal Ellipse(float radius, float radialStep, Image image, int? randomSeed)
        {
            _safeRandom = randomSeed is null
                ? new LockingRandom()
                : new LockingRandom((int)randomSeed);

            _points = new List<SKPoint>();
            PointDistance = radialStep;

            Centre = image.Centre;
            AspectRatio = image.AspectRatio;
            Radius = radius;

            CalculatePoints();
        }

        /// <summary>
        /// Creates a new instance of the <see cref="Ellipse"/> class.
        /// </summary>
        /// <param name="radius">One semi-axis of the ellipse.</param>
        /// <param name="radialStep">The typical distance between two points on the ellipse.</param>
        /// <param name="image">The image which is used to determine the size of the other axis of the ellipse.</param>
        internal Ellipse(float semiMinorAxis, float radialStep, Image image)
            : this(semiMinorAxis, radialStep, image, randomSeed: null)
        {
        }

        private void CalculatePoints()
        {
            if (Radius == 0)
            {
                _points.Add(Centre);
                return;
            }

            GeneratePoints();
        }

        private void GeneratePoints()
        {
            float startingAngle = _safeRandom.PickRandomQuadrant();
            float angleIncrement = GetAngleIncrement(Radius, PointDistance);
            bool clockwise = _safeRandom.RandomFloat() > 0.5;

            float maxAngle;
            if (clockwise)
            {
                maxAngle = startingAngle + 360;
            }
            else
            {
                maxAngle = startingAngle - 360;
                angleIncrement *= -1;
            }

            float angle = startingAngle;

            do
            {
                _points.Add(GetEllipsePoint(Radius, angle));
                angle += angleIncrement;
            } while (clockwise ? angle <= maxAngle : angle >= maxAngle);
        }

        private static float GetAngleIncrement(float radius, float step)
            => step * Constants.BaseAngularIncrement / (15 * (float)Math.Sqrt(radius));

        private SKPoint GetEllipsePoint(float semiMinorAxis, float degrees)
        {
            Complex point = Complex.FromPolarCoordinates(semiMinorAxis, degrees.ToRadians());
            float xPosition = Centre.X + (float)point.Real * AspectRatio;
            float yPosition = Centre.Y + (float)point.Imaginary;

            return new SKPoint(xPosition, yPosition);
        }
    }
}
