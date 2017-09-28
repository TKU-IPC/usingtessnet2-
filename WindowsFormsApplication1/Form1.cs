using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using tessnet2;
using mshtml;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 開始辨識
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            webBrowser1.Navigate(textBox1.Text.Trim());
        }

        /// <summary>
        /// 需配合被辨識的網頁做調整
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            var forms = webBrowser1.Document.GetElementsByTagName("form");
            // 取得整個網頁的DOM
            var doc = webBrowser1.Document.DomDocument as HTMLDocument;
            var body = doc.body as HTMLBody;
            IHTMLControlRange range = (IHTMLControlRange)body.createControlRange();

            // 取得驗證碼圖片
            var imgEle = (forms[0].GetElementsByTagName("img"))[1].DomElement as IHTMLControlElement;
            range.add(imgEle);
            range.execCommand("copy", false, Type.Missing);
            Application.DoEvents();

            // 將圖片轉換成.Net物件
            Image img = Clipboard.GetImage();
            Clipboard.Clear();
            pictureBox1.Image = img;
            Application.DoEvents();

            // 將驗證碼圖片灰階化以提高辨識的準確性
            pictureBox2.Image = convert2GrayScale(pictureBox1.Image);
            Application.DoEvents();

            // 使用Tessnet Library辨識驗證碼
            textBox2.Text = parseCaptchaStr(pictureBox2.Image);
            Application.DoEvents();

        }

        /// <summary>
        /// 將驗證碼圖片灰階化以提高辨識的準確性
        /// </summary>
        /// <param name="img"></param>
        /// <returns></returns>
        private Image convert2GrayScale(Image img)
        {
            Bitmap bitmap = new Bitmap(img);

            for (int i = 0; i < bitmap.Width; i++)
            {
                for (int j = 0; j < bitmap.Height; j++)
                {
                    Color pixelColor = bitmap.GetPixel(i, j);
                    byte r = pixelColor.R;
                    byte g = pixelColor.G;
                    byte b = pixelColor.B;

                    byte gray = (byte)(0.299 * (float)r + 0.587 * (float)g + 0.114 * (float)b);
                    r = g = b = gray;
                    pixelColor = Color.FromArgb(r, g, b);

                    bitmap.SetPixel(i, j, pixelColor);
                }
            }

            Image grayImage = Image.FromHbitmap(bitmap.GetHbitmap());
            bitmap.Dispose();
            return grayImage;
        }

        /// <summary>
        /// 使用Tessnet Library辨識驗證碼
        /// </summary>
        /// <param name="image"></param>
        /// <returns></returns>
        private string parseCaptchaStr(Image image)
        {
            var sb = new StringBuilder();
            using (var ocr = new Tesseract())
            {
                // 設定辨識的相關設定(以下設定為辨識數字)
                if (checkBox1.Checked == true)
                {
                    ocr.SetVariable("tessedit_char_whitelist", "0123456789");
                }
                else
                {
                    ocr.SetVariable("tessedit_char_whitelist", "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789");
                }
                // 樣本資料夾位置
                String path = Application.StartupPath;
                path = path.Substring(0, path.LastIndexOf("bin") - 1) + @"\tessdata";
                ocr.Init(path, "eng", true);
                var result = ocr.DoOCR(new Bitmap(image), Rectangle.Empty);
                // 將辨識結果轉換成字串
                result.ForEach(c =>
                {
                    sb.Append(c.Text);
                });
            }
            return sb.ToString();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

    }
}
