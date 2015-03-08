﻿using System;
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

      var fs = File.Open("4.pptx", FileMode.Open);

      var pres = new Presentation(fs);

      var slides = pres.Slides;

      var sw = new Stopwatch();
      sw.Start();

      var presentationdata = new List<SlideMetaData>();

      foreach (var slide in slides)
      {

        var slidedata = new SlideMetaData();

        CollectText(slide.Shapes, slidedata);

        presentationdata.Add(slidedata);

        /*
        foreach (var shape in slide.Shapes)
        {

          if (!shape.IsTextHolder)
          {
           // continue;
          }
          
          if (shape.GetType().Equals(typeof(GroupShape)))
          {
            Debug.WriteLine(shape.GetType().Name);
          }
          
          var textshape = shape as AutoShape;

          if (textshape == null)
          {
            continue;
          }

          var innertext = textshape.TextFrame.Text;
          string title = null;
          var text = new List<string>();


          if (shape.Name.IndexOf("title", StringComparison.OrdinalIgnoreCase) >= 0)
          {

            if (title != null)
            {
              var thos = "maybe";

            }

            title = innertext;
          }
          else
          {
            text.Add(innertext);
          }
        



        }
        */

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

      sw.Stop();
      Debug.WriteLine("elapsed: " + sw.Elapsed.TotalMilliseconds);

      

      while (true)
      {
        Console.WriteLine("enter a key to search on: ");
        var input = Console.ReadLine().ToLowerInvariant();

        sw.Restart();

        List<Tuple<int,int, int>> result;

        if (!wordkey.TryGetValue(input, out result))
        {
          continue;
        }

                
        foreach (var text in result)
        {

          var slideindex = text.Item1;
          var textindex = text.Item2;
          var wordindex = text.Item3;

          //Debug.WriteLine("word: " + input + ", slide: " + slideindex + ", line: " + textindex + ", word: " + wordindex);

          var line = presentationdata[slideindex].text[textindex];
          //var words = presentationdata[slideindex].text[textindex].Split('\x20');

          var min = Math.Max(wordindex - 25, 0);
          var max = Math.Min(line.Length-1, wordindex + 25);

          var context = line.Substring(min, max - min);
          /*
          for (var i = min; i < max; i++)
          {
            if (i>0){
              context += " ";
            }

            context += words[i];
          }*/

          
          //Console.WriteLine(context);

        }

        sw.Stop();
        Debug.WriteLine("results: " + result.Count + ", search time : " + sw.Elapsed.TotalMilliseconds);

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
          Debug.WriteLine(shape.GetType().Name);

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
      
       // var text = new List<string>();


        if (shape.Name.IndexOf("title", StringComparison.OrdinalIgnoreCase) >= 0)
        {

          if (slidedata.title != null)
          {

            var issue = true;
          }
          else
          {

            slidedata.title = innertext;
          }
        }
        else
        {
          slidedata.text.Add(innertext);
        }



      }


    }


    static void ParseDocument()
    {

      //Console.WriteLine("opening " + args[0]);
      var liscense = File.OpenRead("Aspose.Total.lic");
      var mylicense = new Aspose.Slides.License();
      mylicense.SetLicense(liscense);

      Aspose.Slides.Export.PdfOptions PdfOptions = new Aspose.Slides.Export.PdfOptions();
      PdfOptions.SufficientResolution = 800;
      PdfOptions.SaveMetafilesAsPng = true;
      PdfOptions.JpegQuality = 100;

      var fs = File.Open("4.pptx", FileMode.Open);

      var pres = new Presentation(fs);

      var slides = pres.Slides;
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


        var ms = new MemoryStream();
        var bmp = slide.GetThumbnail(size);
        bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
        PdfDocument document = new PdfDocument();
        //document.Info.Title = "Created with PDFsharp";

        PdfPage page = document.AddPage();
        XGraphics gfx = XGraphics.FromPdfPage(page);

        XFont font = new XFont("Verdana", 16);


        XImage image = XImage.FromGdiPlusImage(bmp);

        double height = (page.Width * image.PixelHeight) / image.PixelWidth;

        gfx.DrawImage(image, 0, 0, page.Width, height);

        //gfx.DrawImage(image, (dx - width) / 2, (dy - height) / 2, width, height);
        //      new XRect(0, height+40, page.Width, page.Height),

        var rect = new XRect(40, height + 40, page.Width - 40, page.Height - (height + 40));
        XTextFormatter tf = new XTextFormatter(gfx);
        tf.DrawString(note, font, XBrushes.Black, rect, XStringFormats.TopLeft);

        document.Save("images/" + i + ".pdf");


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