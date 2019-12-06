using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace New_Employee_Mail_Sender
{
    public partial class Form1 : Form
    {
        private string imageStr;
        private string body;
        public Form1()
        {
            InitializeComponent();
            imageStr = string.Empty;
            body = string.Empty;
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("ka-GE");
            //System.Globalization.CultureInfo ci = new System.Globalization.CultureInfo("ka-GE");
            //inp_startDate.Format = DateTimePickerFormat.Custom;
            //inp_startDate.CustomFormat = ci.DateTimeFormat.ShortDatePattern;
        }

        public void FillHtml()
        {
            if (string.IsNullOrEmpty(imageStr))
            {
                MessageBox.Show("სურათის ატვირთვა სავალდებულოა");
                return;
            }

            using (StreamReader reader = new StreamReader("index.html"))
            {
                body = reader.ReadToEnd();
            }
            body = body.Replace("{fullname}", inp_fullname.Text.Trim());
            body = body.Replace("{position}", inp_position.Text.Trim());
            body = body.Replace("{department}", inp_department.Text.Trim());
            body = body.Replace("{gankofileba}", inp_ganyofileba.Text.Trim());
            body = body.Replace("{birthDay}", inp_birthDay.Value.ToString("d MMMM"));
            body = body.Replace("{boss}", inp_boss.Text.Trim());
            body = body.Replace("{startDate}", inp_startDate.Value.ToString("d MMMM"));
            body = body.Replace("{email}", inp_email.Text.Trim());
            body = body.Replace("{welcomeText}", inp_welcome_text.Text.Trim());
            body = body.Replace("{welcomeText}", inp_welcome_text.Text.Trim());
            body = body.Replace("{image}", imageStr);
        }

        private void btn_submit_Click(object sender, EventArgs e)
        {
            FillHtml();
            SendMail();
        }

        private void btn_upload_Click(object sender, EventArgs e)
        {
            DialogResult dialog = openFileDialog1.ShowDialog();
            if (dialog == DialogResult.OK)
            {
                var fileName = openFileDialog1.FileName;
                ImageTobase64(fileName);
            }
        }

        public void ImageTobase64(string imgPath)
        {
            byte[] imageBytes = System.IO.File.ReadAllBytes(imgPath);
            string base64String = Convert.ToBase64String(imageBytes);
            ResizeImage(Convert.FromBase64String(base64String));
        }

        //public void Base64ToImage(string image64)
        //{
        //    byte[] imageBytes = Convert.FromBase64String(image64);
        //    ResizeImage(imageBytes);
        //    using (var ms = new MemoryStream(imageBytes, 0, imageBytes.Length))
        //    {
        //        pictureBox1.Image = Image.FromStream(ms, true);
        //    }
        //}

        private void btn_view_template_Click(object sender, EventArgs e)
        {
            FillHtml();
            string path = Path.GetTempPath();
            string fileName = Path.GetRandomFileName();
            File.WriteAllText(path + fileName + ".html", body);
            System.Diagnostics.Process.Start(path + fileName + ".html");
        }

        public void ResizeImage(byte[] imagebase64)
        {
            Image image;
            using (MemoryStream ms = new MemoryStream(imagebase64))
            {
                image = Image.FromStream(ms);
            }
            Bitmap b = new Bitmap(300, 300);
            Graphics g = Graphics.FromImage((Image)b);
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;

            g.DrawImage(image, 0, 0, 300, 300);
            g.Dispose();
            image = (Image)b;

            using (MemoryStream ms = new MemoryStream())
            {
                image.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                byte[] imageBytes = ms.ToArray();
                imageStr = Convert.ToBase64String(imageBytes);
                pictureBox1.Image = Image.FromStream(ms, true);
            }
        }

        public void SendMail()
        {
            try
            {
                Microsoft.Office.Interop.Outlook.Application app = new Microsoft.Office.Interop.Outlook.Application();
                Microsoft.Office.Interop.Outlook.MailItem mailItem = app.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
                mailItem.Subject = $"ახალი თანამშრომელი {inp_fullname.Text.Trim()}";
                mailItem.To = "albert.giulbaziani@lb.ge";
                mailItem.HTMLBody = body;                
                mailItem.Importance = Microsoft.Office.Interop.Outlook.OlImportance.olImportanceHigh;
                mailItem.Display(false);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
