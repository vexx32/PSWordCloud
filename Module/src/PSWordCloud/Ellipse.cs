

using System;
using System.Collections.Generic;
using System.Numerics;
using SkiaSharp;

namespace PSWordCloud
{
    internal class Ellipse
    {
        internal IReadOnlyList<SKPoint> Points { get => _points; }

        internal SKPoint Centre { get; }

        internal float AspectRatio { get; }

        internal float Radius { get; }

        internal float PointDistance { get;  }

        private readonly List<SKPoint> _points;

        private readonly LockingRandom _safeRandom;

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
