using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides;
using System.IO;
using PdfSharp;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Drawing.Layout;
using System.Diagnostics;


namespace pptParse
{
  class Program
  {
    static void Main(string[] args)
    {


      var liscense = File.OpenRead("Aspose.Total.lic");
      var mylicense = new Aspose.Slides.License();
      mylicense.SetLicense(liscense);

      var fs = File.Open(args[0], FileMode.Open);

      var presentation = new Presentation(fs);

      ParseDocument(presentation);
      ProcessText(presentation);

      


    }

    static void ProcessText(Presentation presentation)
    {

      var slides = presentation.Slides;
      
      var presentationdata = new List<SlideMetaData>();

      foreach (var slide in slides)
      {

        var slidedata = new SlideMetaData();

        CollectText(slide.Shapes, slidedata);

        presentationdata.Add(slidedata);

      }

      var wordkey = new Dictionary<string, List<Tuple<int, int, int>>>();

      foreach (var slidedata in presentationdata)
      {

        var dataindex = presentationdata.IndexOf(slidedata);

        foreach (var text in slidedata.text)
        {

          var textindex = slidedata.text.IndexOf(text);

          var words = text.Split('\x20');

          var linesum = 0;

          for (var i = 0; i < words.Length; i++)
          //foreach (var word in words)
          {

            List<Tuple<int, int, int>> list;
            var ele = new Tuple<int, int, int>(dataindex, textindex, linesum);

            if (wordkey.TryGetValue(words[i].ToLowerInvariant(), out list))
            {
              list.Add(ele);
            }
            else
            {

              list = new List<Tuple<int, int, int>>();
              list.Add(ele);

              wordkey.Add(words[i].ToLowerInvariant(), list);
            }

            linesum += words[i].Length;
            linesum++;

          }


        }

      }

     
      while (true)
      {
        Console.WriteLine("enter a key to search on: ");
        var input = Console.ReadLine().ToLowerInvariant();
              

        List<Tuple<int, int, int>> result;

        if (!wordkey.TryGetValue(input, out result))
        {
          continue;
        }


        foreach (var text in result)
        {

          var slideindex = text.Item1;
          var textindex = text.Item2;
          var wordindex = text.Item3;

          var line = presentationdata[slideindex].text[textindex];


          var min = Math.Max(wordindex - 25, 0);
          var max = Math.Min(line.Length - 1, wordindex + 25);

          var context = line.Substring(min, max - min);

        }
               

      }


    }

    static void CollectText(IShapeCollection shapes, SlideMetaData slidedata)
    {

      foreach (var shape in shapes)
      {

        if (!shape.IsTextHolder)
        {
          // continue;
        }

        if (shape.GetType().Equals(typeof(GroupShape)))
        {
          
          var groupshape = shape as GroupShape;
          CollectText(groupshape.Shapes, slidedata);

          continue;
        }

        var textshape = shape as AutoShape;

        if (textshape == null)
        {
          continue;
        }

        var innertext = textshape.TextFrame.Text;


        if (slidedata.title == null && shape.Name.IndexOf("title", StringComparison.OrdinalIgnoreCase) >= 0)
        {        
          slidedata.title = innertext;          
        }
        else
        {
          slidedata.text.Add(innertext);
        }
      }
    }

    static void ParseDocument(Presentation presentation)
    {


      var slides = presentation.Slides;
      var size = new System.Drawing.Size(960, 720);

      for (var i = 0; i < slides.Count; i++)
      {

        var slide = slides[i];
        var notes = slide.NotesSlide;
        var note = "";

        if (notes != null)
        {
          note = notes.NotesTextFrame.Text;
        }


        using (var ms = new MemoryStream())
        {
          var bmp = slide.GetThumbnail(size);
          bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
          PdfDocument document = new PdfDocument();
          //document.Info.Title = "Created with PDFsharp";

          PdfPage page = document.AddPage();
          XGraphics gfx = XGraphics.FromPdfPage(page);

          XFont font = new XFont("Verdana", 16);


          XImage image = XImage.FromGdiPlusImage(bmp);


          double width = page.Width - 80;
          double height = (width * image.PixelHeight) / image.PixelWidth;

          gfx.DrawImage(image, 40, 0, width, height);


          var rect = new XRect(40, height + 40, page.Width - 40, page.Height - (height + 40));
          XTextFormatter tf = new XTextFormatter(gfx);
          tf.DrawString(note, font, XBrushes.Black, rect, XStringFormats.TopLeft);

          document.Save("images/" + i + ".pdf");
        }

      }

    }


  }

  public class SlideMetaData
  {
    public List<string> text { get; set; }
    public string title { get; set; }

    public SlideMetaData()
    {
      text = new List<string>();
    }


  }
}
