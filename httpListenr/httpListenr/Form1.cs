using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using Microsoft.Office.Interop.Word;
using System.Web;
using Spire.Pdf;
using Spire.Pdf.Graphics;
using System.Web.Script.Serialization;

namespace httpListenr
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Control.CheckForIllegalCrossThreadCalls = false;

            Thread Thread = new Thread(SimpleListenerExample);
            Thread.IsBackground = true;
            Thread.Start();


         
            
        }
        private void SimpleListenerExample()
        {
            string prefixes = "http://192.168.1.18:12000/";

            HttpListener listener = new HttpListener();
            listener.Prefixes.Add(prefixes);
            listener.Start();


        
            while (true)
            {
                Thread.Sleep(1000);
                listener.BeginGetContext(new AsyncCallback(HandleRequest), listener);
            }
                return;
            while (true)
            {

                textBox1.AppendText("Listening..." + System.Environment.NewLine);
                // Note: The GetContext method blocks while waiting for a request.
                HttpListenerContext context = listener.GetContext(); //阻塞
                HttpListenerRequest request = context.Request;

                string operation = HttpUtility.UrlDecode(request.QueryString["operation"]);
                string filePath = HttpUtility.UrlDecode(request.QueryString["filePath"]);


                string responseString = "<HTML><BODY> Hello world111!</BODY></HTML>";



                textBox1.AppendText(filePath + System.Environment.NewLine);

                try
                {
                    // WORDZH==============================================================
                    if (operation == "wordToPdf")
                    {
                    
                        string pdf_filePath = HttpUtility.UrlDecode(request.QueryString["pdf_filePath"]);
                        Wordhandle wordhandle = new Wordhandle(filePath);
                        wordhandle.killAllWord();
                        wordhandle.wordToPdf(pdf_filePath);
                        responseString = string.Format(" (code): (0) , (Count): (10) ");

                    }
                    if (operation == "MergeFiles")
                    {
                        string pdf_filePath = HttpUtility.UrlDecode(request.QueryString["pdf_filePath"]);
                        string filePaths = HttpUtility.UrlDecode(request.QueryString["filePaths"]);
                        var serializer = new JavaScriptSerializer();
                        String[] files = serializer.Deserialize<String[]>(filePaths);

                        PdfDocumentBase doc = PdfDocument.MergeFiles(files);
                        doc.Save(pdf_filePath, FileFormat.PDF);
                        responseString = string.Format(" (code): (0) , (Count): (10) ");
                    }
                    // WORDZH==============================================================
                    if (operation == "getPdfPageCount")
                    {

                        Pdfhandle Pdfhandle = new Pdfhandle(filePath);

                        int pageCount = Pdfhandle.getPdfPageCount();
                        responseString = string.Format(" (pageCount): ({0}) , (code): (0) ", pageCount);
                    }

                    // WORDZH==============================================================
                    if (operation == "getWordPageCount")
                    {

                        Wordhandle wordhandle = new Wordhandle(filePath);
                        wordhandle.killAllWord();
                        int pageCount = wordhandle.getWordPageCount();
                        textBox1.AppendText("word文档页数：" + pageCount + System.Environment.NewLine);
                        responseString = string.Format(" (pageCount): ({0}) , (code): (0) ", pageCount);
                    }
                    if (operation == "imgToPdf")
                    {
                        string pdf_filePath = HttpUtility.UrlDecode(request.QueryString["pdf_filePath"]);

                        Pdfhandle wordhandle = new Pdfhandle(filePath);
                        wordhandle.imgToPdf(pdf_filePath);
                        textBox1.AppendText(filePath + System.Environment.NewLine);
                        textBox1.AppendText(pdf_filePath + System.Environment.NewLine);
                        textBox1.AppendText("图片转pdf执行："  + System.Environment.NewLine);
                        responseString = string.Format(" (pageCount): (0) , (code): (0) ");
                    }
                    // WORDZH==============================================================
                    if (operation == "getPPTPageCount")
                    {
                        Microsoft.Office.Interop.PowerPoint.Application pptApp = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
                        Microsoft.Office.Interop.PowerPoint.Presentation ppt = pptApp.Presentations.Open(filePath, Microsoft.Office.Core.MsoTriState.msoTrue);
                        int pageCount = ppt.Slides.Count;
                        ppt.Close();
                        GC.Collect();
                        responseString = string.Format(" (pageCount): ({0}) , (code): (0) ", pageCount);
                    }

                }
                catch (Exception exception)
                {
                    responseString = string.Format(" (pageCount): (null) , (code): (-1) , (message): ({0}) ", exception.Message);


                    textBox1.AppendText(exception.Message + System.Environment.NewLine);
                }



                // Obtain a response object.
                HttpListenerResponse response = context.Response;//响应
                                                                 // Construct a response.

                byte[] buffer = System.Text.Encoding.UTF8.GetBytes(responseString);
                // Get a response stream and write the response to it.
                response.ContentLength64 = buffer.Length;
                System.IO.Stream output = response.OutputStream;
                output.Write(buffer, 0, buffer.Length);
                // You must close the output stream.
                output.Close();


            }


        }
        public void HandleRequest(IAsyncResult ar)
        {

            HttpListener listener = ar.AsyncState as HttpListener;
            HttpListenerContext context = listener.EndGetContext(ar);

            // Note: The GetContext method blocks while waiting for a request.

            HttpListenerRequest request = context.Request;

            string operation = HttpUtility.UrlDecode(request.QueryString["operation"]);
            string filePath = HttpUtility.UrlDecode(request.QueryString["filePath"]);


            string responseString = "<HTML><BODY> Hello world111!</BODY></HTML>";



            textBox1.AppendText(filePath + System.Environment.NewLine);

            try
            {
                // WORDZH==============================================================
                if (operation == "wordToPdf")
                {

                    string pdf_filePath = HttpUtility.UrlDecode(request.QueryString["pdf_filePath"]);
                    Wordhandle wordhandle = new Wordhandle(filePath);
                //    wordhandle.killAllWord();
                    wordhandle.wordToPdf(pdf_filePath);
                    responseString = string.Format(" (code): (0) , (Count): (10) ");

                }
                if (operation == "pptToPdf")
                {

                    string pdf_filePath = HttpUtility.UrlDecode(request.QueryString["pdf_filePath"]);
                    PPThandle ppthandle = new PPThandle(filePath);
                    ppthandle.pptToPdf(pdf_filePath);
                    responseString = string.Format(" (code): (0) , (Count): (10) ");

                }
                if (operation == "MergeFiles")
                {
                    string pdf_filePath = HttpUtility.UrlDecode(request.QueryString["pdf_filePath"]);
                    string filePaths = HttpUtility.UrlDecode(request.QueryString["filePaths"]);
                    var serializer = new JavaScriptSerializer();
                    String[] files = serializer.Deserialize<String[]>(filePaths);

                    PdfDocumentBase doc = PdfDocument.MergeFiles(files);
                    doc.Save(pdf_filePath, FileFormat.PDF);
                    doc.Close();
                    responseString = string.Format(" (code): (0) , (Count): (10) ");
                }
                // WORDZH==============================================================
                if (operation == "getPdfPageCount")
                {

                    Pdfhandle Pdfhandle = new Pdfhandle(filePath);

                    int pageCount = Pdfhandle.getPdfPageCount();
                    responseString = string.Format(" (pageCount): ({0}) , (code): (0) ", pageCount);
                }

                // WORDZH==============================================================
                if (operation == "getWordPageCount")
                {

                    Wordhandle wordhandle = new Wordhandle(filePath);
                //    wordhandle.killAllWord();
                    int pageCount = wordhandle.getWordPageCount();
                    textBox1.AppendText("word文档页数：" + pageCount + System.Environment.NewLine);
                    responseString = string.Format(" (pageCount): ({0}) , (code): (0) ", pageCount);
                }
                if (operation == "imgToPdf")
                {
                    string pdf_filePath = HttpUtility.UrlDecode(request.QueryString["pdf_filePath"]);

                    Pdfhandle wordhandle = new Pdfhandle(filePath);
                    wordhandle.imgToPdf(pdf_filePath);
                    textBox1.AppendText(filePath + System.Environment.NewLine);
                    textBox1.AppendText(pdf_filePath + System.Environment.NewLine);
                    textBox1.AppendText("图片转pdf执行：" + System.Environment.NewLine);
                    responseString = string.Format(" (pageCount): (0) , (code): (0) ");
                }
                // WORDZH==============================================================
                if (operation == "getPPTPageCount")
                {
                    Microsoft.Office.Interop.PowerPoint.Application pptApp = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
                    Microsoft.Office.Interop.PowerPoint.Presentation ppt = pptApp.Presentations.Open(filePath, Microsoft.Office.Core.MsoTriState.msoTrue);
                    int pageCount = ppt.Slides.Count;
                    ppt.Close();
                    GC.Collect();
                    responseString = string.Format(" (pageCount): ({0}) , (code): (0) ", pageCount);
                }

            }
            catch (Exception exception)
            {
                responseString = string.Format(" (pageCount): (null) , (code): (-1) , (message): ({0}) ", exception.Message);


                textBox1.AppendText(exception.Message + System.Environment.NewLine);
            }



            HttpListenerResponse response = context.Response;//响应
                                                             // Construct a response.

            byte[] buffer = System.Text.Encoding.UTF8.GetBytes(responseString);
            // Get a response stream and write the response to it.
            response.ContentLength64 = buffer.Length;
            System.IO.Stream output = response.OutputStream;
            output.Write(buffer, 0, buffer.Length);
            // You must close the output stream.
            output.Close();
        }
    }



    //1.Ctrl + A选中要整理的代码
    //2.Ctrl + K
    //3.Ctrl + F

    public class Filehanlde
    {
        public string filePath;
        public Filehanlde(string FilePath)
        {
            filePath = FilePath;
        }
    }


    public class Pdfhandle : Filehanlde
    {
        public Pdfhandle(string FilePath) : base(FilePath)
        {
            filePath = FilePath;
        }

        public int getPdfPageCount()
        {
            Spire.Pdf.PdfDocument pdf = new Spire.Pdf.PdfDocument();
            pdf.LoadFromFile(filePath);
            int pageCount = pdf.Pages.Count;
            pdf.Close();
            GC.Collect();
            return pageCount;
        }

        public void imgToPdf(string pdf_filePath)
        {
            //创建一个PdfDocument对象
            Spire.Pdf.PdfDocument doc = new Spire.Pdf.PdfDocument();
            doc.PageSettings.SetMargins(0);

            Image image = Image.FromFile(filePath);
            float width = image.PhysicalDimension.Width;
            float height = image.PhysicalDimension.Height;
            //添加与图片大小相同的页面
            PdfPageBase page = doc.Pages.Add(new SizeF(width, height));
            //声明一个 PdfImage 变量
            PdfImage pdfImage;
            //如果图片宽度大于页面宽度
            if (width > page.Canvas.ClientSize.Width)
            {
                //调整图片大小以适合页面宽度
                float widthFitRate = width / page.Canvas.ClientSize.Width;
                Size size = new Size((int)(width / widthFitRate), (int)(height / widthFitRate));
                Bitmap scaledImage = new Bitmap(image, size);
                //将缩放后的图片加载到 PdfImage 对象
                pdfImage = PdfImage.FromImage(scaledImage);
            }
            else
            {
                pdfImage = PdfImage.FromImage(image);
            }
            page.Canvas.DrawImage(pdfImage, 0, 0, pdfImage.Width, pdfImage.Height);
            doc.SaveToFile(pdf_filePath);
        }
     
    }
    public class Wordhandle : Filehanlde
    {
        public Wordhandle(string FilePath) : base(FilePath)
        {
            filePath = FilePath;
        }

        public bool wordToPdf(string pdf_filePath)
        {
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.ApplicationClass();

            Microsoft.Office.Interop.Word.Document doc = wordApp.Documents.Open(filePath);
            doc.ExportAsFixedFormat(pdf_filePath, WdExportFormat.wdExportFormatPDF);
            doc.Close(WdSaveOptions.wdDoNotSaveChanges);
            wordApp.Quit();

            return true;
        }
        public int getWordPageCount()
        {

            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.ApplicationClass();
            Microsoft.Office.Interop.Word.Document doc = wordApp.Documents.Open(filePath);
            int pageCount = doc.ComputeStatistics(WdStatistic.wdStatisticPages);

            doc.Close(WdSaveOptions.wdDoNotSaveChanges);
            wordApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
            GC.Collect();
            return pageCount; 
        }
        public void killAllWord()
        {
            foreach (System.Diagnostics.Process p in System.Diagnostics.Process.GetProcessesByName("wps")) { p.Kill(); }
            foreach (System.Diagnostics.Process p in System.Diagnostics.Process.GetProcessesByName("winword")) { p.Kill(); }
        }
    }
    public class PPThandle : Filehanlde
    {
        public PPThandle(string FilePath) : base(FilePath)
        {
            filePath = FilePath;
        }
        public void pptToPdf(string pdf_filePath)
        {
            Microsoft.Office.Interop.PowerPoint.Application pptApp = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
            Microsoft.Office.Interop.PowerPoint.Presentation ppt = pptApp.Presentations.Open(filePath, Microsoft.Office.Core.MsoTriState.msoTrue);

            ppt.ExportAsFixedFormat(pdf_filePath, Microsoft.Office.Interop.PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF);
            ppt.Close();
            pptApp.Quit();
        }
    }

}

