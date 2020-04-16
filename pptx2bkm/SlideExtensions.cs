using System;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Presentation;

namespace pptx2bkm
{
    public static class SlideExtensions
    {
        public static string GetSlideTitle(this Slide slide)
        {
            if (slide == null) throw new ArgumentNullException(nameof(slide));

            string paragraphSeparator = null;
            var shapes = slide.Descendants<Shape>().Where(x => ShapeExtensions.IsTitleShape(x));

            var builder = new StringBuilder();
            foreach (var shape in shapes)
            {
                foreach (var paragraph in shape.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>())
                {
                    builder.Append(paragraphSeparator);
                    foreach (var text in paragraph.Descendants<DocumentFormat.OpenXml.Drawing.Text>())
                    {
                        builder.Append(text.Text);
                    }

                    paragraphSeparator = Environment.NewLine;
                }
            }

            return builder.ToString();
        }
    }
}