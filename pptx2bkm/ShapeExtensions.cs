using System;
using DocumentFormat.OpenXml.Presentation;

namespace pptx2bkm
{
    public static class ShapeExtensions
    {
        public static bool IsTitleShape(this Shape shape)
        {
            if (shape == null) throw new ArgumentNullException(nameof(shape));

            var placeHolderShape = shape.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties
                .GetFirstChild<PlaceholderShape>();
            if (placeHolderShape?.Type != null && placeHolderShape.Type.HasValue)
            {
                switch ((PlaceholderValues)placeHolderShape.Type)
                {
                    case PlaceholderValues.Title:
                    case PlaceholderValues.CenteredTitle:
                        return true;
                    default:
                        return false;
                }
            }

            return false;
        }
    }
}