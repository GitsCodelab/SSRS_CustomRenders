using Microsoft.ReportingServices.Interfaces;
using Microsoft.ReportingServices.OnDemandReportRendering;
using Microsoft.ReportingServices.Rendering.DataRenderer;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CustomRenders.CSV
{
    public class CSVRender: IRenderingExtension
    {
        public string LocalizedName => "CSVRender";


        private Encoding _encoding = Encoding.UTF8;
        
        private string _extension = "csv";
        private string _mimeType = "text/csv";
        string _name;
        

        private string _recordDelimiter = "\r\n";

        private string _qualifier = "\"";
        private bool _suppressLineBreaks;
        private string _fieldDelimiter = ",";

        private bool m_excelMode = true;


        //Stream to which the intermediate report will be rendered
        private Stream intermediateStream;
        private bool _willSeek;
        private StreamOper _operation;
        private bool m_noHeader;
        private bool m_suppressLineBreaks;

        public Stream IntermediateCreateAndRegisterStream(string name, string extension, Encoding encoding, string mimeType, bool willSeek,
            Microsoft.ReportingServices.Interfaces.StreamOper operation)
        {
            _name = name;
            _encoding = encoding;
            _extension = extension;
            _mimeType = mimeType;
            _operation = operation;
            _willSeek = willSeek;
            intermediateStream = new MemoryStream();
            return intermediateStream;
        }
        public bool Render(Microsoft.ReportingServices.OnDemandReportRendering.Report report, NameValueCollection reportServerParameters, NameValueCollection deviceInfo,
            NameValueCollection clientCapabilities, ref Hashtable renderProperties, CreateAndRegisterStream createAndRegisterStream)
        {
            _name = report.Name;
            ParseDeviceInfo(deviceInfo);
            Microsoft.ReportingServices.Rendering.DataRenderer.CsvReport c = new CsvReport();
            c.Render(report, reportServerParameters, deviceInfo, clientCapabilities, ref renderProperties, new CreateAndRegisterStream(IntermediateCreateAndRegisterStream));


            intermediateStream.Position = 0;
            //Register stream for new rendering extension
            Stream outputStream = createAndRegisterStream(_name, _extension, _encoding, _mimeType, _willSeek, _operation);

            //put stream update code here

            //Copy the stream to the outout stream
            byte[] buffer = new byte[32768];

            string textFromMemoryStream;

            // Method 1: Using StreamReader (Simplest, often preferred)
            using (StreamReader reader = new StreamReader(intermediateStream, Encoding.UTF8)) // Specify encoding if needed
            {
                textFromMemoryStream = reader.ReadToEnd();
            }
            var textFromMemoryStream2 = textFromMemoryStream.Replace("\r\n\r\n", "\r\n");




            var o = Encoding.UTF8.GetBytes(textFromMemoryStream2);
            outputStream.Write(o, 0, o.Length);

            //while (true)
            //{
            //    int read = intermediateStream.Read(buffer, 0, buffer.Length);

            //    if (read <= 0) break;
            //    outputStream.Write(buffer, 0, read);
            //    //sameh
            //    StreamReader reader = new StreamReader(intermediateStream);
            //    string text = reader.ReadToEnd();
            //}


            intermediateStream.Close();


            return true;
            //throw new NotImplementedException();
        }

        public void GetRenderingResource(CreateAndRegisterStream createAndRegisterStreamCallback, NameValueCollection deviceInfo)
        {
            // throw new NotImplementedException();
        }

        public bool RenderStream(string streamName, Microsoft.ReportingServices.OnDemandReportRendering.Report report, NameValueCollection reportServerParameters, NameValueCollection deviceInfo, NameValueCollection clientCapabilities, ref Hashtable renderProperties, CreateAndRegisterStream createAndRegisterStream)
        {
            return true;
            //throw new NotImplementedException();
        }

        public void SetConfiguration(string configuration)
        {
            //throw new NotImplementedException();
        }

        private void ParseDeviceInfo(NameValueCollection deviceInfo)
        {
            if (deviceInfo == null)
            {
                return;
            }
             
            if (deviceInfo["ExcelMode"] != null && !bool.TryParse(deviceInfo["ExcelMode"], out m_excelMode))
            {
                m_excelMode = true;
            }

            if (deviceInfo["NoHeader"] != null && !bool.TryParse(deviceInfo["NoHeader"], out m_noHeader))
            {
                m_noHeader = false;
            }
            if (deviceInfo["SuppressLineBreaks"] != null && !bool.TryParse(deviceInfo["SuppressLineBreaks"], out m_suppressLineBreaks))
            {
                m_suppressLineBreaks = false;
            }

            _extension = deviceInfo["FileExtension"];
            if (_extension == null)
            {
                _extension = deviceInfo["Extension"];
                if (_extension == null)
                {
                    _extension = "csv";
                }
            }
            _qualifier = deviceInfo["Qualifier"];
            if (_qualifier == null)
            {
                _qualifier = "\"";
            }
            _recordDelimiter = deviceInfo["RecordDelimiter"];
            if (_recordDelimiter == null)
            {
                _recordDelimiter = "\r\n";
            }
            _fieldDelimiter = deviceInfo["FieldDelimiter"];
            if (_fieldDelimiter == null)
            {
                _fieldDelimiter = ",";
            }
        }
    }
}
