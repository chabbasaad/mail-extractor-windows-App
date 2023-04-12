using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Security.Cryptography;

namespace WindowsFormsApplication1
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            System.Text.Encoding utf_8 = System.Text.Encoding.UTF8;
            string myString;
            myString = "Hello World";
            byte[] data = Encoding.ASCII.GetBytes(myString);
           // textBox1.Text = data.ToString();
            // This is our Unicode string:
            string s_unicode = "abcéabc";

            // Convert a string to utf-8 bytes.
            byte[] utf8Bytes = System.Text.Encoding.UTF8.GetBytes(s_unicode);

            // Convert utf-8 bytes to a string.
            string s_unicode2 = System.Text.Encoding.UTF8.GetString(utf8Bytes);

          //  MessageBox.Show(Convert.ToString(utf8Bytes));
            string codage1 = BitConverter.ToString(data).Replace("-", " ");
            string codage64 = Convert.ToBase64String(data);
           // MessageBox.Show(codage64);
            string myData = "Here is a string to encode.";
          
            textBox3.Clear();
            if (comboBox1.SelectedItem == "Base64")
            {
                string myDataEncoded = EncodeTo64(textBox1.Text);
                textBox3.Text = myDataEncoded;

                string myDataUnencoded = DecodeFrom64(myDataEncoded);
            }
            if (comboBox1.SelectedItem == "MD5")
            {
                string myDataUnencodedmd5 = MD5Hash(textBox1.Text);
                // MessageBox.Show(myDataUnencoded);
                textBox3.Text = myDataUnencodedmd5;
            }
            if (comboBox1.SelectedItem == "Ascii")
            {
                Encoding iso8859_1 = Encoding.GetEncoding("iso-8859-1");
                Encoding unicode = Encoding.Unicode;

                string sample = "abcéabc";

                byte[] isoBytes = iso8859_1.GetBytes(textBox1.Text);
                string isoAsUnicode = unicode.GetString(isoBytes);
                foreach (var b in isoBytes)
                    //MessageBox.Show(b + " ");
                    //MessageBox.Show(isoAsUnicode);
                    textBox3.Text += isoAsUnicode;

                byte[] uniBytes = Encoding.Convert(iso8859_1, unicode, isoBytes);
                string uniAsUnicode = unicode.GetString(uniBytes);

                /*  foreach (var b in uniBytes)
                      //MessageBox.Show(b + " ");
                  MessageBox.Show(uniAsUnicode);*/

            }

          
        }
        static public string EncodeTo64(string toEncode)
        {

            byte[] toEncodeAsBytes

                  = System.Text.ASCIIEncoding.ASCII.GetBytes(toEncode);

            string returnValue

                  = System.Convert.ToBase64String(toEncodeAsBytes);

            return returnValue;

        }
        static public string DecodeFrom64(string encodedData)
        {

            byte[] encodedDataAsBytes  = System.Convert.FromBase64String(encodedData);

            string returnValue = System.Text.ASCIIEncoding.ASCII.GetString(encodedDataAsBytes);

            return returnValue;

        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox3.Clear();
            if (comboBox1.SelectedItem == "Base64")
            {
                string myDataEncoded = EncodeTo64(textBox1.Text);
                string myDataUnencoded = DecodeFrom64(myDataEncoded);
               // MessageBox.Show(myDataUnencoded);
                textBox3.Text = myDataUnencoded;
            }



            if (comboBox1.SelectedItem == "MD5")
            {
               
            }
        }
        public static string MD5Hash(string input)
        {
            StringBuilder hash = new StringBuilder();
            MD5CryptoServiceProvider md5provider = new MD5CryptoServiceProvider();
            byte[] bytes = md5provider.ComputeHash(new UTF8Encoding().GetBytes(input));

            for (int i = 0; i < bytes.Length; i++)
            {
                hash.Append(bytes[i].ToString("x2"));
            }
            return hash.ToString();
        }
      
        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
          
        }
        static public string EncodeToutf8(string toEncode)
        {

            System.Text.Encoding utf_8 = System.Text.Encoding.UTF8;

          
            //string to utf
            byte[] utf = System.Text.Encoding.UTF8.GetBytes(toEncode);
          

            //utf to string
            string s2 = System.Text.Encoding.UTF8.GetString(utf);

       

            return s2;

        }

        private void button4_Click(object sender, EventArgs e)
        {
           /* string[] city1 = { "Vancouver", "Richmond", "Delta", "Surrey", "Burnaby" };
            string[] city2 = { "Vancouver", "Richmond", "Delta", "Surrey", "Burnaby", "Coquitlam", "Mission" };
            listBox1.Items.AddRange(city1);
            listBox2.Items.AddRange(city2);
           // richTextBox1.Lines = city1.ToArray();
          //  richTextBox2.Lines = city2.ToArray();
            foreach (var objLI in listBox2.Items)
            {
                if (listBox1.Items.Contains(objLI))
                    listBox3.Items.Add(objLI); //Adding Shared List
                else
                    listBox4.Items.Add(objLI); // Different List
              //  richTextBox3.Lines = 
            }
            */
            foreach (var objLI1 in richTextBox2.Lines)
            {
                
                if (richTextBox1.Lines.Contains(objLI1))
                    richTextBox4.Text += objLI1 + "\n";
                // listBox3.Items.Add(objLI); //Adding Shared List
                else
                    // listBox4.Items.Add(objLI); // Different List
                   
                richTextBox3.Text += objLI1 +"\n";
            }

        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            foreach (var objLI1 in richTextBox2.Lines)
            {

                if (richTextBox1.Lines.Contains(objLI1))
                    richTextBox4.Text += objLI1 + "\n";
                // listBox3.Items.Add(objLI); //Adding Shared List
                else
                    // listBox4.Items.Add(objLI); // Different List

                    richTextBox3.Text += objLI1 + "\n";
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            System.Text.Encoding utf_8 = System.Text.Encoding.UTF8;
            string myString;
            myString = "Hello World";
            byte[] data = Encoding.ASCII.GetBytes(myString);
            // textBox1.Text = data.ToString();
            // This is our Unicode string:
            string s_unicode = "abcéabc";

            // Convert a string to utf-8 bytes.
            byte[] utf8Bytes = System.Text.Encoding.UTF8.GetBytes(s_unicode);

            // Convert utf-8 bytes to a string.
            string s_unicode2 = System.Text.Encoding.UTF8.GetString(utf8Bytes);

            //  MessageBox.Show(Convert.ToString(utf8Bytes));
            string codage1 = BitConverter.ToString(data).Replace("-", " ");
            string codage64 = Convert.ToBase64String(data);
            // MessageBox.Show(codage64);
            string myData = "Here is a string to encode.";

            textBox3.Clear();
            if (comboBox1.SelectedItem == "Base64")
            {
                string myDataEncoded = EncodeTo64(textBox1.Text);
                textBox3.Text = myDataEncoded;

                string myDataUnencoded = DecodeFrom64(myDataEncoded);
            }
            if (comboBox1.SelectedItem == "MD5")
            {
                string myDataUnencodedmd5 = MD5Hash(textBox1.Text);
                // MessageBox.Show(myDataUnencoded);
                textBox3.Text = myDataUnencodedmd5;
            }
            if (comboBox1.SelectedItem == "Ascii")
            {
                Encoding iso8859_1 = Encoding.GetEncoding("iso-8859-1");
                Encoding unicode = Encoding.Unicode;

                string sample = "abcéabc";

                byte[] isoBytes = iso8859_1.GetBytes(textBox1.Text);
                string isoAsUnicode = unicode.GetString(isoBytes);
                foreach (var b in isoBytes)
                    //MessageBox.Show(b + " ");
                    //MessageBox.Show(isoAsUnicode);
                    textBox3.Text += isoAsUnicode;

                byte[] uniBytes = Encoding.Convert(iso8859_1, unicode, isoBytes);
                string uniAsUnicode = unicode.GetString(uniBytes);

                /*  foreach (var b in uniBytes)
                      //MessageBox.Show(b + " ");
                  MessageBox.Show(uniAsUnicode);*/

            }

        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            textBox3.Clear();
            if (comboBox1.SelectedItem == "Base64")
            {
                string myDataEncoded = EncodeTo64(textBox1.Text);
                string myDataUnencoded = DecodeFrom64(myDataEncoded);
                // MessageBox.Show(myDataUnencoded);
                textBox3.Text = myDataUnencoded;
            }



            if (comboBox1.SelectedItem == "MD5")
            {

            }
        }

        private void button4_Click_2(object sender, EventArgs e)
        {
            richTextBox3.Text = "";
            richTextBox4.Text = "";
            foreach (var objLI1 in richTextBox2.Lines)
            {

                if (richTextBox1.Lines.Contains(objLI1))
                    richTextBox4.Text += objLI1 + "\n";
                // listBox3.Items.Add(objLI); //Adding Shared List
                else
                    // listBox4.Items.Add(objLI); // Different List

                    richTextBox3.Text += objLI1 + "\n";
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {

        }

    
    }
}
