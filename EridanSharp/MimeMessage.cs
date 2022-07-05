using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EridanSharp
{
    public class MimeMessage : ICrypt
    {
        private bool isAttachment;
        private string messageBody;

        public Dictionary<string, string> types;

        private string contentType = "text/plain";
        private string encode = "utf-8";
        private string bodyText;
        private string subject;
        private string toEmail;
        private string toName;
        private string fromName;
        private string fromEmail;

        private List<Tuple<FileInfo, string>> attachments;

        private struct FileInfo
        {
            public string Name;
            public string Extension;
            public string Path;
            public FileInfo(string name, string extension, string path)
            {
                Name = name;
                Extension = extension;
                Path = path;
            }
        }

        public MimeMessage()
        {
            attachments = new List<Tuple<FileInfo, string>>();
            types = new Dictionary<string, string>();

            types.Add("aac", "audio/aac");
            types.Add("abw", "application/x-abiword");
            types.Add("arc", "application/x-freearc");
            types.Add("avi", "video/x-msvideo");
            types.Add("azw", "application/vnd.amazon.ebook");
            types.Add("bin", "application/octet-stream");
            types.Add("bmp", "image/bmp");
            types.Add("bz", "application/x-bzip");
            types.Add("bz2", "application/x-bzip2");
            types.Add("csh", "application/x-csh");
            types.Add("css", "text/css");
            types.Add("csv", "text/csv");
            types.Add("doc", "application/msword");
            types.Add("docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
            types.Add("eot", "application/vnd.ms-fontobject");
            types.Add("epub", "application/epub+zip");
            types.Add("gz", "application/gzip");
            types.Add("gif", "image/gif");
            types.Add("html", "text/html");
            types.Add("ico", "image/vnd.microsoft.icon");
            types.Add("ics", "text/calendar");
            types.Add("jar", "application/java-archive");
            types.Add("jpeg", "image/jpeg");
            types.Add("jpg", "image/jpeg");
            types.Add("js", "text/javascript");
            types.Add("json", "application/json");
            types.Add("jsonld", "application/ld+json");
            types.Add("mjs", "text/javascript");
            types.Add("mp3", "audio/mpeg");
            types.Add("mpeg", "video/mpeg");
            types.Add("mpkg", "application/vnd.apple.installer+xml");
            types.Add("odp", "application/vnd.oasis.opendocument.presentation");
            types.Add("ods", "application/vnd.oasis.opendocument.spreadsheet");
            types.Add("odt", "application/vnd.oasis.opendocument.text");
            types.Add("oga", "audio/ogg");
            types.Add("ogv", "video/ogg");
            types.Add("ogx", "application/ogg");
            types.Add("opus", "audio/opus");
            types.Add("otf", "font/otf");
            types.Add("png", "image/png");
            types.Add("pdf", "application/pdf");
            types.Add("php", "application/x-httpd-php");
            types.Add("ppt", "application/vnd.ms-powerpoint");
            types.Add("pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
            types.Add("rar", "application/vnd.rar");
            types.Add("rtf", "application/rtf");
            types.Add("sh", "application/x-sh");
            types.Add("svg", "image/svg+xml");
            types.Add("swf", "application/x-shockwave-flash");
            types.Add("tar", "application/x-tar");
            types.Add("tif", "image/tiff");
            types.Add("tiff", "image/tiff");
            types.Add("ts", "video/mp2t");
            types.Add("ttf", "font/ttf");
            types.Add("txt", "text/plain");
            types.Add("vsd", "application/vnd.visio");
            types.Add("wav", "audio/wav");
            types.Add("weba", "audio/webm");
            types.Add("webm", "video/webm");
            types.Add("webp", "image/webp");
            types.Add("woff", "font/woff");
            types.Add("woff2", "font/woff2");
            types.Add("xhtml", "application/xhtml+xml");
            types.Add("xls", "application/vnd.ms-excel");
            types.Add("xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            types.Add("xml", "application/xml");
            types.Add("xul", "application/vnd.mozilla.xul+xml");
            types.Add("zip", "application/zip");
            types.Add("3gp", "video/3gpp");
            types.Add("7z", "application/x-7z-compressed");
        }


        [DefaultValue("text/plain")]
        public string ContentType
        {
            get
            {
                return contentType;
            }
            set
            {
                contentType = value;
            }
        }
        public string BodyText
        {
            get
            {
                return bodyText;
            }
            set
            {
                bodyText = value;
            }
        }
        public string Subject
        {
            get
            {
                return subject;
            }
            set
            {
                subject = value;
            }
        }
        public string ToEmail
        {
            get
            {
                return toEmail;
            }
            set
            {
                toEmail = value;
            }
        }
        public string ToName
        {
            get
            {
                return toName;
            }
            set
            {
                toName = value;
            }
        }
        public string FromName
        {
            get
            {
                return fromName;
            }
            set
            {
                fromName = value;
            }
        }
        public string FromEmail
        {
            get
            {
                return fromEmail;
            }
            set
            {
                fromEmail = value;
            }
        }

        [DefaultValue("utf-8")]
        public string Encode
        {
            get
            {
                return encode;
            }
            set
            {
                encode = value;
            }
        }



        public void AddAttachment(string path)
        {
            string base64data = "";
            try
            {
                using (FileStream fstream = new FileStream(path, FileMode.OpenOrCreate))
                {
                    Byte[] bytes = new Byte[fstream.Length];
                    fstream.Read(bytes, 0, bytes.Length);
                    base64data = Convert.ToBase64String(bytes);
                }
                //bytes = File.ReadAllBytes(path);
                //String file = Convert.ToBase64String(bytes);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }
            attachments.Add(Tuple.Create(
                new FileInfo(Path.GetFileName(path), Path.GetExtension(path), path), base64data)
                );
        }

        public void ClearAttachments()
        {
            if (attachments.Count > 0)
            {
                attachments.Clear();
            }
        }


        public string GetMessage()
        {
            messageBody =
$@"From: {FromEmail}
To: {ToEmail}
Subject: {Subject}
MIME-Version: 1.0 
Content-Type: multipart/mixed; boundary=""B164240059B29C0E4EFEC397""

--B164240059B29C0E4EFEC397
Content-Type: multipart/alternative; boundary = ""B164240059B29C0E4EFEC397""

--B164240059B29C0E4EFEC397
Content-Type: {ContentType}; charset=""{Encode}""
Content-Transfer-Encoding: quoted-printable

{BodyText}


";

            if (attachments.Count > 0)
            {
                messageBody += "--B164240059B29C0E4EFEC397\n";

                foreach (var attachment in attachments)
                {
                    string contentTypeForAttachment;

                    if (types.ContainsKey(attachment.Item1.Extension))
                    {
                        contentTypeForAttachment = types[attachment.Item1.Extension];
                    }
                    else
                    {
                        contentTypeForAttachment = "application/octet-stream";
                    }

                    messageBody +=
$@"Content-Type: {contentTypeForAttachment}; name=""{attachment.Item1.Name}""
Content-Transfer-Encoding: base64
Content-Disposition: attachment; filename = ""{attachment.Item1.Name}""
Content-Description: {attachment.Item1.Name}

";
                    messageBody += attachment.Item2.ToString();
                }
            }

            messageBody += "\n\n--B164240059B29C0E4EFEC397--";

            Debug.WriteLine(messageBody);

            return messageBody;
        }

        public string GetMessageBase64()
        {

            Debug.WriteLine(Base64Encode(GetMessage()));
            return Base64Encode(GetMessage());
        }


        public string Base64Encode(string data)
        {
            var dataBytes = Encoding.UTF8.GetBytes(data);
            return Convert.ToBase64String(dataBytes);
        }

        public string Base64Decode(string data)
        {
            var dataBytes = Convert.FromBase64String(data);
            return Encoding.UTF8.GetString(dataBytes);
        }
    }
}
