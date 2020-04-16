using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Office2010.PowerPoint;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace pptx2bkm
{
    public static class PresentationDocumentExtensions
    {
        public static IEnumerable<Slide> GetSlides(this PresentationDocument document)
        {
            if (document == null) throw new ArgumentNullException(nameof(document));

            var presentationPart = document.PresentationPart;
            var presentation = presentationPart?.Presentation;
            if (presentation?.SlideIdList == null) yield break;

            foreach (var slideId in presentation.SlideIdList.Elements<SlideId>())
            {
                var slidePart = presentationPart.GetPartById(slideId.RelationshipId) as SlidePart;
                if (slidePart?.Slide != null)
                {
                    yield return slidePart.Slide;
                }
            }
        }

        public static IEnumerable<Slide> GetSlides(this PresentationDocument document, Section section)
        {
            var presentationPart = document.PresentationPart;
            var presentation = presentationPart.Presentation;
            foreach (var entry in section.SectionSlideIdList.Elements<SectionSlideIdListEntry>())
            {
                var slideId = presentation.SlideIdList.Cast<SlideId>().FirstOrDefault(x => x.Id == entry.Id);
                if (slideId != null)
                {
                    var slidePart = presentationPart.GetPartById(slideId.RelationshipId) as SlidePart;
                    if (slidePart?.Slide != null)
                    {
                        yield return slidePart.Slide;
                    }
                }
            }
        }

        public static int RetrieveNumberOfSlide(this PresentationDocument document, Slide slide)
        {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (slide == null) throw new ArgumentNullException(nameof(slide));

            var presentationPart = document.PresentationPart;
            var presentation = presentationPart?.Presentation;
            if (presentation?.SlideIdList == null) return 0;

            var targetSlideId = GetSlideIdBySlide(presentationPart, slide);

            int index = 1;
            foreach (var slideId in presentation.SlideIdList.Elements<SlideId>())
            {
                if (slideId == targetSlideId) break;
                index++;
            }

            return index;
        }

        private static SlideId GetSlideIdBySlide(PresentationPart presentationPart, Slide slide)
        {
            if (presentationPart == null) throw new ArgumentNullException(nameof(presentationPart));
            if (slide == null) throw new ArgumentNullException(nameof(slide));

            Presentation presentation = presentationPart.Presentation;
            return presentation.SlideIdList
                .Cast<SlideId>()
                .FirstOrDefault(x => (presentationPart.GetPartById(x.RelationshipId) as SlidePart)?.Slide == slide);
        }
    }
}