using System;
using System.CommandLine;
using System.CommandLine.Invocation;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Office2010.PowerPoint;
using DocumentFormat.OpenXml.Packaging;

namespace pptx2bkm
{
    class Program
    {
        static int Main(string[] args)
        {
            var rootCommand = new RootCommand()
            {
                new Option(new[] {"--inputFile", "-i"})
                {
                    Argument = new Argument<string>(),
                    Required = true
                },
                new Option(new[] {"--outputFile", "-o"})
                {
                    Argument = new Argument<string>(),
                    Required = true
                }
            };

            rootCommand.Handler = CommandHandler.Create<string, string>(Run);
            return rootCommand.Invoke(args);
        }

        static void Run(string inputFile, string outputFile)
        {
            if (!File.Exists(inputFile)) throw new FileNotFoundException("FILE NOT FOUND", inputFile);
            if (File.Exists(outputFile)) throw new InvalidOperationException($"FILE ALRADY EXIST : {outputFile}");

            using var pptx = PresentationDocument.Open(inputFile, false);
            var presentationPart = pptx.PresentationPart;
            var presentation = presentationPart.Presentation;

            StringBuilder builder = new StringBuilder();

            var sections = presentation.Descendants<Section>();
            foreach (var section in sections)
            {
                var slides = pptx.GetSlides(section).ToList();
                if (!slides.Any()) continue;

                builder.AppendLine($"{section.Name} /{pptx.RetrieveNumberOfSlide(slides.First())}");

                foreach (var slide in slides.Where(slide => slide.Show == null || (slide.Show.HasValue && slide.Show.Value)))
                {
                    builder.AppendLine($"\t{slide.GetSlideTitle()} /{pptx.RetrieveNumberOfSlide(slide)}");
                }
            }

            File.WriteAllText(outputFile, builder.ToString(), Encoding.UTF8);
        }
    }
}
