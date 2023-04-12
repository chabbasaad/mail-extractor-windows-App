using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using S22.Imap;
using System.Net.Mail;
using MailKit.Net.Imap;
using System.Threading;
using System.Text.RegularExpressions;
using System.Net.NetworkInformation;
using MailKit.Search;
using MailKit;
using System.IO;
using System.Net.Sockets;
using System.Diagnostics;
using System.Net;
using MailKit.Security;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Interactions;
using Microsoft.Win32;
using OpenQA.Selenium.PhantomJS;
using OpenQA.Selenium.Remote;
using Google.GData.Client;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Google.Spreadsheets;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Google.Apis.Auth.OAuth2.Flows;
using Google.Apis.Auth.OAuth2.Responses;
using Data = Google.Apis.Sheets.v4.Data;


namespace WindowsFormsApplication1
{

    public partial class Form1 : Form
    {

        public Form1()
        {

            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //string username = "daciaall2016@hotmail.com";
            //   string password = "lola1994";

            /*
             
  ImapClient Client = new ImapClient("imap-mail.outlook.com", 993,username,password, AuthMethod.Login,false,null);
                // IEnumerable<uint> uids = Client.Search(SearchCondition.All(),"INBOX");
                // IEnumerable<MailMessage> allMails = Client.GetMessage(uids,false, "INBOX");
                 // Download each message but skip message parts that are larger than 1 Megabyte.
                 IEnumerable<uint> uids = Client.Search(SearchCondition.From("FQFp2nfXBXVgwt3kQ5@earthdatings.com"));
                 MessageBox.Show("hello");
                 // Fetch the first message and print it's subject and body.
                 if (uids.Count() > 0)
                 {
                     MessageBox.Show("start");
                     MailMessage msg = Client.GetMessage(uids.First());

                     MessageBox.Show("Subject: " + msg.Subject);
                    
                 }

                 Client.Dispose();
             */

            /*
            IEnumerable<string> mailbox = Client.ListMailboxes();
            IEnumerable<uint> allMailsID = Client.Search(SearchCondition.All(), "INBOX");

            //IEnumerable<MailMessage> allMails = imap.GetMessages(allMailsID,,false, "INBOX");

            MailMessage msg = Client.GetMessage(allMailsID.First(), FetchOptions.HeadersOnly);
            MessageBox.Show("Subject: " + msg.Subject);
                        */
        }
        public int proc = 0;
        private void button2_Click(object sender, EventArgs e)
        {
            label5.Text = "";
            // proc = 0;
            progressBar1.Value = 0;
            //  progressBar1.ResetText();
            // MessageBox.Show("start");
            if (comboBox1.Text == "" && comboBox2.Text == "")
            {
                MessageBox.Show("Please Enter Your Login and Password");
            }
            else
            {
                using (var client = new MailKit.Net.Imap.ImapClient())
                {
                    // For demo-purposes, accept all SSL certificates
                    client.ServerCertificateValidationCallback = (s, c, h, v) => true;
                    //"imap.gmail.com"
                    comboBox1.Items.Add(comboBox1.Text);
                    comboBox2.Items.Add(comboBox2.Text);
                    if (comboBox1.Text.Contains("gmail.com"))
                    {
                        client.Connect("imap.gmail.com", 993, true);

                        client.Authenticate(comboBox1.Text, comboBox2.Text);
                    }
                    if (comboBox1.Text.Contains("hotmail.com"))
                    {
                        client.Connect("imap-mail.outlook.com", 993, true);

                        client.Authenticate(comboBox1.Text, comboBox2.Text);

                    }

                    if (comboBox1.Text.Contains("yahoo.com"))
                    {
                        client.Connect("imap.mail.yahoo.com", 993, true);
                        client.Authenticate(comboBox1.Text, comboBox2.Text);
                    }
                    // Get the first personal namespace and list the toplevel folders under it.
                    /*   var personal = client.GetFolder(client.PersonalNamespaces[0]);
                       foreach (var folder in personal.GetSubfolders(false))
                           MessageBox.Show("[folder] {0}", folder.Name);*/
                    var connected = client.IsAuthenticated;
                    // The Inbox folder is always available on all IMAP servers...

                    var inbox = client.Inbox;
                    inbox.Open(FolderAccess.ReadOnly);
                    /*
                    Console.WriteLine("Total messages: {0}", inbox.Count);
                    Console.WriteLine("Recent messages: {0}", inbox.Recent);*/
                    label8.Text = inbox.Count.ToString();
                    for (int i = 0; i < inbox.Count; i++)
                    {
                        this.timer1.Start();
                        var message = inbox.GetMessage(i);
                        // MessageBox.Show("Subject: {0}", message.Subject);
                        listBox1.Items.Add(message.Subject);
                        //  progressBar1.Minimum = 0;
                        // progressBar1.Maximum = i++;
                        //progressBar1.Step = i;
                        // ThreadPool.QueueUserWorkItem(new WaitCallback(AnimateScroll));

                        // here's the change:

                    }
                    client.Disconnect(true);
                    // MessageBox.Show("Done !");
                    progressBar1.Value = 100;

                }





                string sPath = "C:/Users/Public/passandlog.txt";
                if (File.Exists("C:/Users/Public/passandlog.txt"))
                {

                    System.IO.StreamWriter SaveFile = new System.IO.StreamWriter(sPath);
                    foreach (var item in comboBox1.Items)
                    {
                        SaveFile.WriteLine(item);

                    }
                    foreach (var x in comboBox2.Items)
                    {
                        SaveFile.WriteLine(x);
                    }



                    SaveFile.Close();
                }
                else
                {
                    System.IO.StreamWriter SaveFile = new System.IO.StreamWriter(sPath);
                    foreach (var item in comboBox1.Items)
                    {
                        SaveFile.WriteLine(item);
                    }
                    foreach (var item in comboBox2.Items)
                    {
                        SaveFile.WriteLine(item);
                    }

                    SaveFile.Close();
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            label8.Text = "";
            label5.Text = "";
            //proc = 0;
            progressBar1.Value = 0;
            //progressBar1.ResetText();
            if (comboBox1.Text == "" && comboBox2.Text == "")
            {
                MessageBox.Show("Please Enter Your Login and Password");
            }
            else
            {

                using (var client = new MailKit.Net.Imap.ImapClient())
                {
                    // For demo-purposes, accept all SSL certificates
                    client.ServerCertificateValidationCallback = (s, c, h, v) => true;
                    comboBox1.Items.Add(comboBox1.Text);
                    comboBox2.Items.Add(comboBox2.Text);


                    if (comboBox1.Text.Contains("gmail.com"))
                    {
                        client.Connect("imap.gmail.com", 993, true);

                        client.Authenticate(comboBox1.Text, comboBox2.Text);
                        // MessageBox.Show("here");
                        var test = client.IsAuthenticated;
                        //  MessageBox.Show(test.ToString());
                        var spam = client.GetFolder("[Gmail]/Spam");
                        spam.Open(FolderAccess.ReadOnly);
                        label8.Text = spam.Count.ToString();
                        for (int i = 0; i < spam.Count; i++)
                        {
                            // MessageBox.Show("Start");
                            this.timer1.Start();
                            var messagespam = spam.GetMessage(i);
                            // MessageBox.Show("Subject: {0}", messagespam.Subject);
                            listBox1.Items.Add(messagespam.Subject);

                            //  progressBar1.Minimum = 0;
                            // progressBar1.Maximum = 100;
                            // progressBar1.Step = 1;
                            // ThreadPool.QueueUserWorkItem(new WaitCallback(AnimateScroll));

                            // here's the change:

                        }
                    }
                    if (comboBox1.Text.Contains("hotmail.com"))
                    {
                        client.Connect("imap-mail.outlook.com", 993, true);
                        client.Authenticate(comboBox1.Text, comboBox2.Text);
                        var spam = client.GetFolder("junk");
                        spam.Open(FolderAccess.ReadOnly);
                        label8.Text = spam.Count.ToString();
                        for (int i = 0; i < spam.Count; i++)
                        {
                            // MessageBox.Show("Start");
                            this.timer1.Start();
                            var messagespam = spam.GetMessage(i);
                            // MessageBox.Show("Subject: {0}", messagespam.Subject);
                            listBox1.Items.Add(messagespam.Subject);

                            //  progressBar1.Minimum = 0;
                            // progressBar1.Maximum = 100;
                            // progressBar1.Step = 1;
                            // ThreadPool.QueueUserWorkItem(new WaitCallback(AnimateScroll));

                            // here's the change:

                        }
                    }
                    if (comboBox1.Text.Contains("yahoo.com"))
                    {
                        client.Connect("imap.mail.yahoo.com", 993, true);

                        client.Authenticate(comboBox1.Text, comboBox2.Text);
                        var spam = client.GetFolder("Bulk Mail");
                        spam.Open(FolderAccess.ReadOnly);
                        label8.Text = spam.Count.ToString();
                        for (int i = 0; i < spam.Count; i++)
                        {
                            // MessageBox.Show("Start");
                            this.timer1.Start();
                            var messagespam = spam.GetMessage(i);
                            // MessageBox.Show("Subject: {0}", messagespam.Subject);
                            listBox1.Items.Add(messagespam.Subject);
                            //  progressBar1.Minimum = 0;
                            // progressBar1.Maximum = 100;
                            // progressBar1.Step = 1;
                            // ThreadPool.QueueUserWorkItem(new WaitCallback(AnimateScroll));

                            // here's the change:

                        }
                    }

                    // Get the first personal namespace and list the toplevel folders under it.
                    /* var personal = client.GetFolder(client.PersonalNamespaces[0]);
                     foreach (var folder in personal.GetSubfolders(false))
                         MessageBox.Show("[folder] {0}", folder.Name);*/
                    // MessageBox.Show("start");
                    // The Inbox folder is always available on all IMAP servers...


                    client.Disconnect(true);
                    //MessageBox.Show("Done !");
                    progressBar1.Value = 100;

                }

                string sPath = "C:/Users/Public/passandlog.txt";
                if (File.Exists("C:/Users/Public/passandlog.txt"))
                {
                    System.IO.StreamWriter SaveFile = new System.IO.StreamWriter(sPath);
                    foreach (var item in comboBox1.Items)
                    {
                        SaveFile.WriteLine(item);
                    }
                    foreach (var item in comboBox2.Items)
                    {
                        SaveFile.WriteLine(item);
                    }

                    SaveFile.Close();
                }
                else
                {
                    System.IO.StreamWriter SaveFile = new System.IO.StreamWriter(sPath);
                    foreach (var item in comboBox1.Items)
                    {
                        SaveFile.WriteLine(item);
                    }
                    foreach (var item in comboBox2.Items)
                    {
                        SaveFile.WriteLine(item);
                    }

                    SaveFile.Close();
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            this.progressBar1.Increment(1);
            label5.Text = progressBar1.Value.ToString() + " %";
        }

        public static void EnableTab(TabPage page, bool enable)
        {
            EnableControls(page.Controls, enable);
        }
        private static void EnableControls(Control.ControlCollection ctls, bool enable)
        {
            foreach (Control ctl in ctls)
            {
                ctl.Enabled = enable;
                EnableControls(ctl.Controls, enable);
            }
        }


        public static string ipextacrt = "";

        private void Form1_Load(object sender, EventArgs e)
        {
            // Start the BackgroundWorker.
        /*    string whatIsMyIp = "http://ip.c.la/";
            string getIpRegex = @"(?<=<TITLE>.*)\d*\.\d*\.\d*\.\d*(?=</TITLE>)";
            WebClient wc = new WebClient();
            UTF8Encoding utf8 = new UTF8Encoding();
            string requestHtml = "";
            try
            {
                requestHtml = utf8.GetString(wc.DownloadData(whatIsMyIp));
            }

            catch (WebException we)
            {
                // do something with exception
                //Console.Write(we.ToString());
            }
            // MessageBox.Show(requestHtml);
            string ipextract = requestHtml;
            string contentbl2 = ipextract;
            string content2bl2 = Regex.Replace(contentbl2, "_", " ");
            string content1bl2 = ipextract;


            var matchesbl2 = Regex.Matches(content2bl2, @"(\d{0,9}\.\d{0,9}\.\d{0,9}\.\d{0,9})");
            foreach (var matchbl in matchesbl2)
            {
                ipextacrt = matchbl.ToString();
                //  MessageBox.Show(matchbl.ToString());
            }
            string allips = "81.192.142.89" + "81.192.142.176" + "81.192.142.92" + "173.212.201.187 " + "81.192.142.93" + "80.241.213.157";
            string[] array1 = { "81.192.142.89", "81.192.142.176", "81.192.142.92", "173.212.201.187", "81.192.142.93" };
            /*foreach (string lang in array1)
             {*/
           /* if (allips.Contains(ipextacrt))
            {
                //      MessageBox.Show("find it");

            }
            if (!allips.Contains(ipextacrt))
            {
                MessageBox.Show("You don't have permission to access / on this App. Please contact THE Owner.");
                System.Windows.Forms.Application.Exit();
            }*/
            
           // int SerialsCounter = listBox2.Items.Count;
           // label40.Text = SerialsCounter.ToString();
            // link to enable first page and disabled
          //  EnableTab(tabControl1.TabPages[tabControl1.SelectedIndex = 0], true);
          //  EnableTab(tabControl1.TabPages[tabControl1.SelectedIndex = 3], false);
           // EnableTab(tabControl1.TabPages[tabControl1.SelectedIndex = 2],false);

            backgroundWorker1.RunWorkerAsync();



            List<string> liness = new List<string>();
            string path = @"C:/Users/Public/warmuptest.txt";
            using (StreamReader r = new StreamReader(path))
            {
                string line;
                while ((line = r.ReadLine()) != null)
                {
                    liness.Add(line);
                    //listBox2.Items.Add(line);
                }
            }
            // listView1.View = View.Details;
            // listView1.Columns.Add("", 250);

            timer2.Start();
            timer3.Start();
            timer4.Start();
            timer5.Start();
            //  MessageBox.Show("if you are Using yahoo.com and Gmail.com You Need To Turn on Allow apps that use less secure sign-in " + "\n yahoo : "+"https://login.yahoo.com/account/security#other-apps" + "\n gmail :"+"https://myaccount.google.com/lesssecureapps");

            // Start the BackgroundWorker.
            /* string path = "C:/Users/Public/passandlog.txt";
            if (File.Exists("C:/Users/Public/passandlog.txt"))
             {
        
                 StreamReader sr = new StreamReader(path);
               while (sr.Peek() >= 0)
                 {
                   //  comboBox2.Items.Add(sr.ReadLine());
                     comboBox1.Items.Add(sr.ReadLine());
                    
                 }
               
               //  MessageBox.Show(sr.ReadLine());
                 comboBox1.SelectedIndex = -1;
                 comboBox2.SelectedIndex = -1;
                 int val = 0;
                 sr.Close();
            }
           */

            timer2.Start();
            timer3.Start();
            timer4.Start();
            timer5.Start();
            var lines = File.ReadAllLines("C:/Users/Public/passandlog.txt");
            File.WriteAllLines("C:/Users/Public/passandlog.txt", lines.Distinct().ToArray());
            var firstFound = false;
            for (int index = 0; index < lines.Count(); index++)
            {
                /*  if (!firstFound && lines[index].Contains("hotmail.com"))
                  {
                     // MessageBox.Show("true");
                      //MessageBox.Show(lines[index]);
                      comboBox1.Items.Add(lines[index]);
                  }
                  if (!firstFound && lines[index].Contains("gmail.com"))
                  {
                      // MessageBox.Show("true");
                      //MessageBox.Show(lines[index]);
                      comboBox1.Items.Add(lines[index]);
                  }*/
                if (!firstFound && lines[index].Contains("yahoo.com") || lines[index].Contains("gmail.com") || lines[index].Contains("hotmail.com"))
                {
                    // MessageBox.Show("true");
                    //MessageBox.Show(lines[index]);
                    comboBox1.Items.Add(lines[index]);
                }
                else
                {
                    comboBox2.Items.Add(lines[index]);
                }


            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {

            /*proc = 0;
           progressBar1.Value = 0;*/
            if (comboBox1.Text == "" && comboBox2.Text == "")
            {
                MessageBox.Show("Please Enter Your Login and Password");
            }
            else
            {
                using (var client = new MailKit.Net.Imap.ImapClient())
                {
                    // For demo-purposes, accept all SSL certificates
                    client.ServerCertificateValidationCallback = (s, c, h, v) => true;
                    if (comboBox1.Text.Contains("gmail.com"))
                    {
                        client.Connect("imap.gmail.com", 993, true);

                        client.Authenticate(comboBox1.Text, comboBox2.Text);
                        // MessageBox.Show("here");
                        var test = client.IsAuthenticated;
                        //  MessageBox.Show(test.ToString());
                        var spam = client.GetFolder("[Gmail]/Spam");
                        spam.Open(FolderAccess.ReadWrite);
                        int ie = spam.Count;

                        MessageBox.Show("" + spam.Count);
                        for (int i = 0; i < (spam.Count); i++)
                        {
                            //var messagespam = spam.GetMessage(i);
                            var matchFolder = client.GetFolder("inbox");
                            if (i >= 0)
                            {
                                spam.MoveTo(i, matchFolder);
                            }
                            if (i == 0)
                            {
                                spam.MoveTo(i, matchFolder);

                                //  spam.MoveTo(IList<UniqueId> uids,matchFolder,CancellationToken cancellationToken = null);

                            }
                            else
                            {
                                MessageBox.Show("Nothing to move");
                            }


                        }


                        client.Disconnect(true);
                        MessageBox.Show("all emails was Moved to Inbox Folder!");
                    }
                    if (comboBox1.Text.Contains("hotmail.com"))
                    {
                        client.Connect("imap-mail.outlook.com", 993, true);

                        client.Authenticate(comboBox1.Text, comboBox2.Text);
                        var spam = client.GetFolder("junk");
                        //spam.Open(FolderAccess.ReadOnly);
                        spam.Open(FolderAccess.ReadWrite);
                        int ie = spam.Count;
                        // int iz = 0;
                        MessageBox.Show("" + spam.Count);
                        for (int i = 0; i < (spam.Count); i++)
                        {
                            //var messagespam = spam.GetMessage(i);
                            var matchFolder = client.GetFolder("inbox");
                            if (i >= 0)
                            {
                                spam.MoveTo(i, matchFolder);
                            }
                            if (i == 0)
                            {
                                spam.MoveTo(i, matchFolder);
                            }
                            else
                            {
                                MessageBox.Show("Nothing to move");
                            }


                        }


                        client.Disconnect(true);
                        MessageBox.Show("all emails was Moved to Inbox Folder!");
                    }
                    if (comboBox1.Text.Contains("yahoo.com"))
                    {
                        client.Connect("imap.mail.yahoo.com", 993, true);

                        client.Authenticate(comboBox1.Text, comboBox2.Text);
                        var spam = client.GetFolder("Bulk Mail");
                        //spam.Open(FolderAccess.ReadOnly);
                        spam.Open(FolderAccess.ReadWrite);
                        int ie = spam.Count;
                        // int iz = 0;
                        MessageBox.Show("" + spam.Count);
                        for (int i = 0; i < (spam.Count); i++)
                        {
                            //var messagespam = spam.GetMessage(i);
                            var matchFolder = client.GetFolder("inbox");
                            if (i >= 0)
                            {
                                spam.MoveTo(i, matchFolder);
                            }
                            if (i == 0)
                            {
                                spam.MoveTo(i, matchFolder);

                                //  spam.MoveTo(IList<UniqueId> uids,matchFolder,CancellationToken cancellationToken = null);

                            }
                            else
                            {
                                MessageBox.Show("Nothing to move");
                            }


                        }


                        client.Disconnect(true);
                        MessageBox.Show("all emails was Moved to Inbox Folder!");
                    }


                }


                MessageBox.Show("Sorry !! this Operation Still On Procces");
            }
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            using (var client = new MailKit.Net.Imap.ImapClient())
            {
                // For demo-purposes, accept all SSL certificates
                client.ServerCertificateValidationCallback = (s, c, h, v) => true;

                client.Connect("imap-mail.outlook.com", 993, true);

                client.Authenticate(comboBox1.Text, comboBox2.Text);

                var spam = client.GetFolder("junk");
                //spam.Open(FolderAccess.ReadOnly);
                spam.Open(FolderAccess.ReadWrite);

                /*for (int i = 0; i < spam.Count; i++)
                {
                    var messagespam = spam.GetMessage(i);
                    var matchFolder = client.GetFolder("inbox");
                    if (matchFolder != null && spam.Count != 0)
                    {

                        spam.MoveTo(i, matchFolder);

                        MessageBox.Show("" + spam.Count);
                    }

                }*/
                for (int i = 0; i < spam.Count; i++)
                {
                    var messagespam = spam.GetMessage(i);
                    var matchFolder = client.GetFolder("inbox");
                    spam.MoveTo(i, matchFolder);
                }
                client.Disconnect(true);
                MessageBox.Show("all emails was Moved to Inbox Folder!");
            }

        }

        private void button6_Click_1(object sender, EventArgs e)
        {

            if (textBox3.Text == "")
            {
                MessageBox.Show("Chose Your File Name Please!");
            }
            if (textBox3.Text != "")
            {
                string fileName = @"C:/Users/Public/" + textBox3.Text + ".txt";
                //  StreamReader reader = File.OpenText(fileName);
                /*  using (StreamReader sr = File.OpenText(fileName))
                  {
                      string s = "";
                      while ((s = sr.ReadLine()) != null)
                      {
                          MessageBox.Show(s);
                      }
                  }*/
                Process.Start("notepad.exe", @"C:/Users/Public/" + textBox3.Text + ".txt");
                // listBox1.Items.Clear();
                //listBox1.DataSource = File.ReadAllLines("C:/Users/Public/SubjectSaved.txt");
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (textBox3.Text == "")
            {
                MessageBox.Show("Chose Your File Name Please!");
            }
            if (textBox3.Text != "")
            {
                string[] lines = File.ReadAllLines("C:/Users/Public/" + textBox3.Text + ".txt");
                File.WriteAllLines("C:/Users/Public/" + textBox3.Text + ".txt", lines.Distinct().ToArray());

                MessageBox.Show("All Duplicate Lignes are Removed from your file \n" + "C:/Users/Public/" + textBox3.Text + ".txt");
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            SKGL.Validate ValidateAKey = new SKGL.Validate();// create an object
            ValidateAKey.secretPhase = "My$ecretPa$$W0rd"; // the passsword
            ValidateAKey.Key = "LMFME-OTQAF-JVBUP-OKFGP"; // enter a valid key
            Console.WriteLine(ValidateAKey.IsValid); // check whether the key has been modified or not
        }

        private void button9_Click(object sender, EventArgs e)
        {
            File.WriteAllText(@"C:/Users/Public/passandlog.txt", "");
            MessageBox.Show("Your Login and Password Are Cleaned with Success !");
            comboBox1.Text = "";
            comboBox2.Text = "";
        }
        private static string ExtractJist(string freeText)
        {
            StringBuilder patternBuilder = new StringBuilder();
            patternBuilder.Append(@"First name: (?<fn>.*$)\n")
                .Append("Last name: (?<ln>.*$)\n")
                .Append("Address: (?<address>.*$)\n")
                .Append("City: (?<city>.*$)\n")
                .Append("State: (?<state>.*$)\n")
                .Append("Zip: (?<zip>.*$)");
            Match match = Regex.Match(freeText, patternBuilder.ToString(), RegexOptions.Multiline | RegexOptions.IgnoreCase);
            string fullname = string.Concat(match.Groups["fn"], " ", match.Groups["ln"]);
            string address = match.Groups["address"].ToString();
            string city = match.Groups["city"].ToString();
            string state = match.Groups["state"].ToString();
            string zip = match.Groups["zip"].ToString();
            return string.Concat(fullname, "|", address, "|", city, "|", state, "|", zip);
        }
        private void button10_Click(object sender, EventArgs e)
        {
            /* string source = @"First name: Elvis
 Last name: Presley
 Address: 1 Heaven Street
 City: Memphis
 State: TN
 Zip: 12345
 ";
             MessageBox.Show(ExtractJist(source));*/
            // string sourceString = "sid [1544764] srv [CFT256] remip [10.0.128.31] fwf []...remip [10.0.128.41] fwf []... ";
            string content = textBox1.Text;
            string content2 = Regex.Replace(content, "_", " ");
            //   MessageBox.Show(content2);
            var matches = Regex.Matches(content2, @"(\d{0,9}\.\d{0,9}\.\d{0,9}\.\d{0,9})");
            // var matches = Regex.Matches(content2, @"\b(\d{0,9}\.\d{0,9}\.\d{0,9}\.\d{0,9})\b");
            textBox1.Clear();
            foreach (var match in matches)
            {

                richTextBox5.Text += match.ToString();
                textBox1.AppendText(Environment.NewLine);
            }
            textBox1.Text = string.Join(Environment.NewLine, textBox1.Lines.Distinct());
            label22.Text = matches.Count.ToString();
            MessageBox.Show("Done!");

            /*  SautinSoft.HtmlToRtf h = new SautinSoft.HtmlToRtf();           
              // string htmlString = @"<b>Hello World!</b>";
              string htmlString = @"\b(\d{0,9}\.\d{0,9}\.\d{0,9}\.\d{0,9})\b";
                 h.OutputFormat = HtmlToRtf.eOutputFormat.TextUnicode;
               string textString = h.ConvertString(htmlString);
               MessageBox.Show(textString);*/
        }

        private void button11_Click(object sender, EventArgs e)
        {


        }

        private void button11_Click_1(object sender, EventArgs e)
        {
            textBox1.Text = string.Join(Environment.NewLine, textBox1.Lines.Distinct());
            label22.Text = textBox1.Lines.Count().ToString();
            MessageBox.Show("Done!");
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }



        private void timer2_Tick(object sender, EventArgs e)
        {
            textBox1.Text = DateTime.Now.ToString();
        }

        private void timer2_Tick_1(object sender, EventArgs e)
        {
            label12.Text = DateTime.Now.ToLongTimeString();
            timer2.Start();
        }

        private void timer3_Tick(object sender, EventArgs e)
        {
            label14.Text = TimeZoneInfo.ConvertTimeBySystemTimeZoneId(DateTime.UtcNow, "Central Europe Standard Time").ToLongTimeString();
            timer3.Start();
        }

        private void timer4_Tick(object sender, EventArgs e)
        {
            label16.Text = TimeZoneInfo.ConvertTimeBySystemTimeZoneId(DateTime.UtcNow, "AUS Eastern Standard Time").ToLongTimeString();
            timer4.Start();
        }

        private void timer5_Tick(object sender, EventArgs e)
        {
            label18.Text = TimeZoneInfo.ConvertTimeBySystemTimeZoneId(DateTime.UtcNow, "Atlantic Standard Time").ToLongTimeString();
            timer5.Start();
        }
        // method cntl a
        private void TbUsername_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == System.Windows.Forms.Keys.A)
            {
                textBox1.SelectAll();
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            int SerialsCounter = textBox1.Lines.Length;
            label22.Text = SerialsCounter.ToString();
            //TextBox tex = new TextBox();
            textBox1.Focus();



        }


        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int SerialsCounter = listBox1.Items.Count;
            label23.Text = SerialsCounter.ToString();
        }

        private void button10_Click_1(object sender, EventArgs e)
        {
            richTextBox7.Clear();



            /* string source = @"First name: Elvis
Last name: Presley
Address: 1 Heaven Street
City: Memphis
State: TN
Zip: 12345
";
           MessageBox.Show(ExtractJist(source));*/
            // string sourceString = "sid [1544764] srv [CFT256] remip [10.0.128.31] fwf []...remip [10.0.128.41] fwf []... ";
            string content = richTextBox5.Text;
            string content2 = Regex.Replace(content, "_", " ");
            string content1 = richTextBox5.Text;
            /*string[] content3 = Regex.Split(content1, @"\W");
               MessageBox.Show(content3[0]);*/

            var matches = Regex.Matches(content2, @"(\d{0,9}\.\d{0,9}\.\d{0,9}\.\d{0,9})");
            Regex regex = new Regex(@"[A-Za-z]+Story[0-9]*");
            var mat = Regex.Matches(richTextBox5.Text, @"[A-Za-z]+[0-9]*");
            //var matches1 = Regex.Match(content3[0],@"^\d");
            foreach (var match in mat)
            {
                // MessageBox.Show(match.ToString());
                //  richTextBox7.Text += richTextBox5.Lines[match].Split('_').First() + "\n";
                /* richTextBox5.Text += match.ToString();
                 richTextBox5.AppendText(Environment.NewLine);*/
                richTextBox7.Text += match.ToString();
                richTextBox7.AppendText(Environment.NewLine);

            }
            /*  for (int n = 0; n < richTextBox5.Lines.Length; ++n)
              {
                  richTextBox7.Text += richTextBox5.Lines[n].Split('_').First() + "\n";
                  //richTextBox7.Text += richTextBox5.Lines[n].Split('_').Last() + "\n";
              }*/

            // var matches = Regex.Matches(content2, @"\b(\d{0,9}\.\d{0,9}\.\d{0,9}\.\d{0,9})\b");
            richTextBox5.Clear();
            foreach (var match in matches)
            {

                richTextBox5.Text += match.ToString();
                richTextBox5.AppendText(Environment.NewLine);

            }

            richTextBox7.Text = string.Join(Environment.NewLine, richTextBox7.Lines.Distinct());
            richTextBox5.Text = string.Join(Environment.NewLine, richTextBox5.Lines.Distinct());
            label22.Text = matches.Count.ToString();
            //label28.Text = matches1.Count.ToString();
            /* string str = "Title = Car Promotion, Model = BMW 323";
             Debug.Write(str.Split('=').Last());*/

            MessageBox.Show("Done!");

            /*  SautinSoft.HtmlToRtf h = new SautinSoft.HtmlToRtf();           
              // string htmlString = @"<b>Hello World!</b>";
              string htmlString = @"\b(\d{0,9}\.\d{0,9}\.\d{0,9}\.\d{0,9})\b";
                 h.OutputFormat = HtmlToRtf.eOutputFormat.TextUnicode;
               string textString = h.ConvertString(htmlString);
               MessageBox.Show(textString);*/
        }

        private void button11_Click_2(object sender, EventArgs e)
        {
            richTextBox5.Text = string.Join(Environment.NewLine, richTextBox5.Lines.Distinct());
            label22.Text = richTextBox5.Lines.Count().ToString();
            MessageBox.Show("Done!");
        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            richTextBox6.Clear();
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            if (richTextBox6.Text == null || textBox3.Text == null)
            {
                MessageBox.Show("Nothing in the List to Save");
            }
            else
            {
                string sPath = "C:/Users/Public/" + textBox3.Text + ".txt";

                System.IO.StreamWriter SaveFile = new System.IO.StreamWriter(sPath);

                for (int n = 0; n < richTextBox6.Lines.Length; ++n)
                {
                    SaveFile.WriteLine(richTextBox6.Lines[n]);
                }

                SaveFile.Close();

                MessageBox.Show("Lists Saved!" + "in this folder " + " C:/Users/Public/" + textBox3.Text + ".txt");
            }
            /*
             *  if (richTextBox5.Text == null || textBox3.Text == null)
            {
                MessageBox.Show("Nothing in the List to Save");
            }
            else
            {
                string sPath = "C:/Users/Public/" + textBox3.Text + ".txt";

                System.IO.StreamWriter SaveFile = new System.IO.StreamWriter(sPath);
                foreach (var item in listBox1.Items)
                {
                    SaveFile.WriteLine(item);
                }

                SaveFile.Close();

                MessageBox.Show("Lists Saved!" + "in this folder " + " C:/Users/Public/" + textBox3.Text + ".txt");
            }
             * */
        }

        private void button7_Click_1(object sender, EventArgs e)
        {
            if (textBox3.Text == "")
            {
                MessageBox.Show("Chose Your File Name Please!");
            }
            if (textBox3.Text != "")
            {
                string[] lines = File.ReadAllLines("C:/Users/Public/" + textBox3.Text + ".txt");
                File.WriteAllLines("C:/Users/Public/" + textBox3.Text + ".txt", lines.Distinct().ToArray());

                MessageBox.Show("All Duplicate Lignes are Removed from your file \n" + "C:/Users/Public/" + textBox3.Text + ".txt");
            }
        }

        private void button6_Click_2(object sender, EventArgs e)
        {
            if (textBox3.Text == "")
            {
                MessageBox.Show("Chose Your File Name Please!");
            }
            if (textBox3.Text != "")
            {
                string fileName = @"C:/Users/Public/" + textBox3.Text + ".txt";
                //  StreamReader reader = File.OpenText(fileName);
                /*  using (StreamReader sr = File.OpenText(fileName))
                  {
                      string s = "";
                      while ((s = sr.ReadLine()) != null)
                      {
                          MessageBox.Show(s);
                      }
                  }*/
                Process.Start("notepad.exe", @"C:/Users/Public/" + textBox3.Text + ".txt");
                // listBox1.Items.Clear();
                //listBox1.DataSource = File.ReadAllLines("C:/Users/Public/SubjectSaved.txt");
            }
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {
            timer2.Start();
            timer3.Start();
            timer4.Start();
            timer5.Start();
        }

        private void button1_Click_2(object sender, EventArgs e)
        {
            /*proc = 0;
          progressBar1.Value = 0;*/
            if (comboBox1.Text == "" && comboBox2.Text == "")
            {
                MessageBox.Show("Please Enter Your Login and Password");
            }
            else
            {
                using (var client = new MailKit.Net.Imap.ImapClient())
                {
                    // For demo-purposes, accept all SSL certificates
                    client.ServerCertificateValidationCallback = (s, c, h, v) => true;
                    if (comboBox1.Text.Contains("gmail.com"))
                    {
                        client.Connect("imap.gmail.com", 993, true);

                        client.Authenticate(comboBox1.Text, comboBox2.Text);
                        // MessageBox.Show("here");
                        var test = client.IsAuthenticated;
                        //  MessageBox.Show(test.ToString());
                        var spam = client.GetFolder("[Gmail]/Spam");
                        spam.Open(FolderAccess.ReadWrite);
                        int ie = spam.Count;

                        int iz = 0;
                        // MessageBox.Show("" + spam.Count);
                        if (spam.Count == 0)
                        {
                            MessageBox.Show("Spam Empty");
                        }
                        else
                        {
                            while ((spam.Count) > 0)
                            {
                                var matchFolder = client.GetFolder("inbox");
                                spam.MoveTo(iz, matchFolder);
                            }
                           


                            client.Disconnect(true);
                            MessageBox.Show("all emails was Moved to Inbox Folder!");
                        }


                    }
                    if (comboBox1.Text.Contains("hotmail.com"))
                    {
                        client.Connect("imap-mail.outlook.com", 993, true);

                        client.Authenticate(comboBox1.Text, comboBox2.Text);
                        var spam = client.GetFolder("junk");

                        //spam.Open(FolderAccess.ReadOnly);
                        spam.Open(FolderAccess.ReadWrite);
                        int ie = spam.Count;
                        int iz = 0;
                      
                        if (spam.Count == 0)
                        {
                            MessageBox.Show("spam Empty");
                        }
                        else
                        {
                            while ((spam.Count) > 0)
                            {
                                var matchFolder = client.GetFolder("inbox");
                                spam.MoveTo(iz, matchFolder);
                                // code moving from spam to archive
                                /* var archive = client.GetFolder("inbox/chabbatest");
                        MessageBox.Show(archive.Count.ToString());
                                spam.MoveTo(iz, archive);*/
                                // MessageBox.Show(spam.Count.ToString());

                            }
                            


                            client.Disconnect(true);
                            MessageBox.Show("all emails was Moved to Inbox Folder!");
                        }
                    }
                    if (comboBox1.Text.Contains("yahoo.com"))
                    {

                        client.Connect("imap.mail.yahoo.com", 993, true);

                        client.Authenticate(comboBox1.Text, comboBox2.Text);
                        var spam = client.GetFolder("Bulk Mail");
                        //spam.Open(FolderAccess.ReadOnly);
                        spam.Open(FolderAccess.ReadWrite);
                        int ie = spam.Count;

                        int iz = 0;
                        // MessageBox.Show("" + spam.Count);
                        if (spam.Count == 0)
                        {
                            MessageBox.Show("Spam Empty");
                        }
                        else
                        {
                            while ((spam.Count) > 0)
                            {
                                var matchFolder = client.GetFolder("inbox");
                                spam.MoveTo(iz, matchFolder);
                                // code moving from spam to archive
                                /* var archive = client.GetFolder("inbox/chabbatest");
                        MessageBox.Show(archive.Count.ToString());
                                spam.MoveTo(iz, archive);*/
                                // MessageBox.Show(spam.Count.ToString());

                            }
                            // boucle for 
                            /* for (int i = 0; i < (spam.Count); i++)
                             {
                                 //var messagespam = spam.GetMessage(i);
                                 var matchFolder = client.GetFolder("inbox");
                                 if (i >= 0)
                                 {
                                     spam.MoveTo(i, matchFolder);
                                 }
                               /*  if (i == 0)
                                 {
                                     spam.MoveTo(i, matchFolder);
                                 }*/
                            /*   else
                               {
                                   MessageBox.Show("Nothing to move");
                               }*/


                            //}
                            // end boucle for


                            client.Disconnect(true);
                            MessageBox.Show("all emails was Moved to Inbox Folder!");
                        }
                    }
                    if (comboBox1.Text.Contains("aol.com"))
                    {

                        client.Connect("imap.aol.com", 993, true);

                        client.Authenticate(comboBox1.Text, comboBox2.Text);
                        var spam = client.GetFolder("Bulk Mail");
                        //spam.Open(FolderAccess.ReadOnly);
                        spam.Open(FolderAccess.ReadWrite);
                        int ie = spam.Count;

                        int iz = 0;
                        // MessageBox.Show("" + spam.Count);
                        if (spam.Count == 0)
                        {
                            MessageBox.Show("Spam Empty");
                        }
                        else
                        {
                            while ((spam.Count) > 0)
                            {
                                var matchFolder = client.GetFolder("inbox");
                                spam.MoveTo(iz, matchFolder);
                                // code moving from spam to archive
                                /* var archive = client.GetFolder("inbox/chabbatest");
                        MessageBox.Show(archive.Count.ToString());
                                spam.MoveTo(iz, archive);*/
                                // MessageBox.Show(spam.Count.ToString());

                            }
                            // boucle for 
                            /* for (int i = 0; i < (spam.Count); i++)
                             {
                                 //var messagespam = spam.GetMessage(i);
                                 var matchFolder = client.GetFolder("inbox");
                                 if (i >= 0)
                                 {
                                     spam.MoveTo(i, matchFolder);
                                 }
                               /*  if (i == 0)
                                 {
                                     spam.MoveTo(i, matchFolder);
                                 }*/
                            /*   else
                               {
                                   MessageBox.Show("Nothing to move");
                               }*/


                            //}
                            // end boucle for


                            client.Disconnect(true);
                            MessageBox.Show("all emails was Moved to Inbox Folder!");
                        }
                    }


                }


                //MessageBox.Show("Sorry !! this Operation Still On Procces");
            }
        }

        private void button9_Click_1(object sender, EventArgs e)
        {
            File.WriteAllText(@"C:/Users/Public/passandlog.txt", "");
            MessageBox.Show("Your Login and Password Are Cleaned with Success !");
            comboBox1.Text = "";
            comboBox2.Text = "";
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            label8.Text = "";
            label5.Text = "";
            //proc = 0;
            progressBar1.Value = 0;
            //progressBar1.ResetText();
            if (comboBox1.Text == "" && comboBox2.Text == "")
            {
                MessageBox.Show("Please Enter Your Login and Password");
            }
            else
            {

                using (var client = new MailKit.Net.Imap.ImapClient())
                {
                    // For demo-purposes, accept all SSL certificates
                    client.ServerCertificateValidationCallback = (s, c, h, v) => true;
                    comboBox1.Items.Add(comboBox1.Text);
                    comboBox2.Items.Add(comboBox2.Text);


                    if (comboBox1.Text.Contains("gmail.com"))
                    {
                        client.Connect("imap.gmail.com", 993, true);

                        client.Authenticate(comboBox1.Text, comboBox2.Text);
                        // MessageBox.Show("here");
                        var test = client.IsAuthenticated;
                        //  MessageBox.Show(test.ToString());
                        var spam = client.GetFolder("[Gmail]/Spam");
                        spam.Open(FolderAccess.ReadOnly);
                        label8.Text = spam.Count.ToString();
                        for (int i = 0; i < spam.Count; i++)
                        {
                            // MessageBox.Show("Start");
                            this.timer1.Start();
                            var messagespam = spam.GetMessage(i);
                            // MessageBox.Show("Subject: {0}", messagespam.Subject);
                            listBox1.Items.Add(messagespam.Subject);
                            richTextBox6.Text += messagespam.Subject + "\n";
                            //  progressBar1.Minimum = 0;
                            // progressBar1.Maximum = 100;
                            // progressBar1.Step = 1;
                            // ThreadPool.QueueUserWorkItem(new WaitCallback(AnimateScroll));

                            // here's the change:

                        }
                    }
                    if (comboBox1.Text.Contains("hotmail.com"))
                    {
                        client.Connect("imap-mail.outlook.com", 993, true);
                        client.Authenticate(comboBox1.Text, comboBox2.Text);
                        var spam = client.GetFolder("junk");
                        spam.Open(FolderAccess.ReadOnly);
                        label8.Text = spam.Count.ToString();
                        for (int i = 0; i < spam.Count; i++)
                        {
                            // MessageBox.Show("Start");
                            this.timer1.Start();
                            var messagespam = spam.GetMessage(i);
                            // MessageBox.Show("Subject: {0}", messagespam.Subject);
                            listBox1.Items.Add(messagespam.Subject);
                            richTextBox6.Text += messagespam.Subject + "\n";
                            //  progressBar1.Minimum = 0;
                            // progressBar1.Maximum = 100;
                            // progressBar1.Step = 1;
                            // ThreadPool.QueueUserWorkItem(new WaitCallback(AnimateScroll));

                            // here's the change:

                        }
                    }
                    if (comboBox1.Text.Contains("outlook.com"))
                    {
                        client.Connect("imap-mail.outlook.com", 993, true);
                        client.Authenticate(comboBox1.Text, comboBox2.Text);
                        var spam = client.GetFolder("junk");
                        spam.Open(FolderAccess.ReadOnly);
                        label8.Text = spam.Count.ToString();
                        for (int i = 0; i < spam.Count; i++)
                        {
                            // MessageBox.Show("Start");
                            this.timer1.Start();
                            var messagespam = spam.GetMessage(i);
                            // MessageBox.Show("Subject: {0}", messagespam.Subject);
                            listBox1.Items.Add(messagespam.Subject);
                            richTextBox6.Text += messagespam.Subject + "\n";
                            //  progressBar1.Minimum = 0;
                            // progressBar1.Maximum = 100;
                            // progressBar1.Step = 1;
                            // ThreadPool.QueueUserWorkItem(new WaitCallback(AnimateScroll));

                            // here's the change:

                        }
                    }
                    if (comboBox1.Text.Contains("yahoo.com"))
                    {
                        client.Connect("imap.mail.yahoo.com", 993, true);

                        client.Authenticate(comboBox1.Text, comboBox2.Text);
                        var spam = client.GetFolder("Bulk Mail");
                        spam.Open(FolderAccess.ReadOnly);
                        label8.Text = spam.Count.ToString();
                        for (int i = 0; i < spam.Count; i++)
                        {
                            // MessageBox.Show("Start");
                            this.timer1.Start();
                            var messagespam = spam.GetMessage(i);
                            // MessageBox.Show("Subject: {0}", messagespam.Subject);
                            listBox1.Items.Add(messagespam.Subject);
                            richTextBox6.Text += messagespam.Subject + "\n";
                            //  progressBar1.Minimum = 0;
                            // progressBar1.Maximum = 100;
                            // progressBar1.Step = 1;
                            // ThreadPool.QueueUserWorkItem(new WaitCallback(AnimateScroll));

                            // here's the change:

                        }
                    }
                    if (comboBox1.Text.Contains("aol.com"))
                    {
                        client.Connect("imap.aol.com", 993, true);

                        client.Authenticate(comboBox1.Text, comboBox2.Text);
                        var spam = client.GetFolder("Bulk Mail");
                        spam.Open(FolderAccess.ReadOnly);
                        label8.Text = spam.Count.ToString();
                        for (int i = 0; i < spam.Count; i++)
                        {
                            // MessageBox.Show("Start");
                            this.timer1.Start();
                            var messagespam = spam.GetMessage(i);
                            // MessageBox.Show("Subject: {0}", messagespam.Subject);
                            listBox1.Items.Add(messagespam.Subject);
                            richTextBox6.Text += messagespam.Subject + "\n";
                            //  progressBar1.Minimum = 0;
                            // progressBar1.Maximum = 100;
                            // progressBar1.Step = 1;
                            // ThreadPool.QueueUserWorkItem(new WaitCallback(AnimateScroll));

                            // here's the change:

                        }
                    }

                    // Get the first personal namespace and list the toplevel folders under it.
                    /* var personal = client.GetFolder(client.PersonalNamespaces[0]);
                     foreach (var folder in personal.GetSubfolders(false))
                         MessageBox.Show("[folder] {0}", folder.Name);*/
                    // MessageBox.Show("start");
                    // The Inbox folder is always available on all IMAP servers...


                    client.Disconnect(true);
                    //MessageBox.Show("Done !");
                    progressBar1.Value = 100;

                }

                string sPath = "C:/Users/Public/passandlog.txt";
                if (File.Exists("C:/Users/Public/passandlog.txt"))
                {
                    System.IO.StreamWriter SaveFile = new System.IO.StreamWriter(sPath);
                    foreach (var item in comboBox1.Items)
                    {
                        SaveFile.WriteLine(item);
                    }
                    foreach (var item in comboBox2.Items)
                    {
                        SaveFile.WriteLine(item);
                    }

                    SaveFile.Close();
                }
                else
                {
                    System.IO.StreamWriter SaveFile = new System.IO.StreamWriter(sPath);
                    foreach (var item in comboBox1.Items)
                    {
                        SaveFile.WriteLine(item);
                    }
                    foreach (var item in comboBox2.Items)
                    {
                        SaveFile.WriteLine(item);
                    }

                    SaveFile.Close();
                }
            }
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            label5.Text = "";
            // proc = 0;
            progressBar1.Value = 0;
            //  progressBar1.ResetText();
            // MessageBox.Show("start");
            if (comboBox1.Text == "" && comboBox2.Text == "")
            {
                MessageBox.Show("Please Enter Your Login and Password");
            }
            else
            {
                using (var client = new MailKit.Net.Imap.ImapClient())
                {
                    // For demo-purposes, accept all SSL certificates
                    client.ServerCertificateValidationCallback = (s, c, h, v) => true;
                    //"imap.gmail.com"
                    comboBox1.Items.Add(comboBox1.Text);
                    comboBox2.Items.Add(comboBox2.Text);

                    if (comboBox3.Text.Contains("body"))
                    {
                        if (comboBox1.Text.Contains("gmail.com"))
                        {
                            client.Connect("imap.gmail.com", 993, true);

                            client.Authenticate(comboBox1.Text, comboBox2.Text);
                        }
                        if (comboBox1.Text.Contains("hotmail.com"))
                        {
                            client.Connect("imap-mail.outlook.com", 993, true);

                            client.Authenticate(comboBox1.Text, comboBox2.Text);

                        }

                        if (comboBox1.Text.Contains("yahoo.com"))
                        {
                            client.Connect("imap.mail.yahoo.com", 993, true);
                            client.Authenticate(comboBox1.Text, comboBox2.Text);
                        }
                        if (comboBox1.Text.Contains("aol.com"))
                        {
                            client.Connect("imap.aol.com", 993, true);
                            client.Authenticate(comboBox1.Text, comboBox2.Text);
                        }

                        // Get the first personal namespace and list the toplevel folders under it.
                        /*   var personal = client.GetFolder(client.PersonalNamespaces[0]);
                           foreach (var folder in personal.GetSubfolders(false))
                               MessageBox.Show("[folder] {0}", folder.Name);*/
                        var connected = client.IsAuthenticated;
                        // The Inbox folder is always available on all IMAP servers...

                        var inbox = client.Inbox;
                        label8.Text = inbox.Count.ToString();
                        inbox.Open(FolderAccess.ReadOnly);
                        /*
                        Console.WriteLine("Total messages: {0}", inbox.Count);
                        Console.WriteLine("Recent messages: {0}", inbox.Recent);*/
                        label8.Text = inbox.Count.ToString();
                        for (int i = 0; i < inbox.Count; i++)
                        {
                            this.timer1.Start();
                            var message = inbox.GetMessage(i);
                            // MessageBox.Show("Subject: {0}", message.Subject);
                            listBox1.Items.Add(message.Subject);
                            richTextBox6.Text += message.HtmlBody + "\n";
                            //  progressBar1.Minimum = 0;
                            // progressBar1.Maximum = i++;
                            //progressBar1.Step = i;
                            // ThreadPool.QueueUserWorkItem(new WaitCallback(AnimateScroll));

                            // here's the change:

                        }
                    }
                    if (comboBox3.Text.Contains("to"))
                    {
                        if (comboBox1.Text.Contains("gmail.com"))
                        {
                            client.Connect("imap.gmail.com", 993, true);

                            client.Authenticate(comboBox1.Text, comboBox2.Text);
                        }
                        if (comboBox1.Text.Contains("hotmail.com"))
                        {
                            client.Connect("imap-mail.outlook.com", 993, true);

                            client.Authenticate(comboBox1.Text, comboBox2.Text);

                        }

                        if (comboBox1.Text.Contains("yahoo.com"))
                        {
                            client.Connect("imap.mail.yahoo.com", 993, true);
                            client.Authenticate(comboBox1.Text, comboBox2.Text);
                        }
                        if (comboBox1.Text.Contains("aol.com"))
                        {
                            client.Connect("imap.aol.com", 993, true);
                            client.Authenticate(comboBox1.Text, comboBox2.Text);
                        }

                        // Get the first personal namespace and list the toplevel folders under it.
                        /*   var personal = client.GetFolder(client.PersonalNamespaces[0]);
                           foreach (var folder in personal.GetSubfolders(false))
                               MessageBox.Show("[folder] {0}", folder.Name);*/
                        var connected = client.IsAuthenticated;
                        // The Inbox folder is always available on all IMAP servers...

                        var inbox = client.Inbox;
                        label8.Text = inbox.Count.ToString();
                        inbox.Open(FolderAccess.ReadOnly);
                        /*
                        Console.WriteLine("Total messages: {0}", inbox.Count);
                        Console.WriteLine("Recent messages: {0}", inbox.Recent);*/
                        label8.Text = inbox.Count.ToString();
                        for (int i = 0; i < inbox.Count; i++)
                        {
                            this.timer1.Start();
                            var message = inbox.GetMessage(i);
                            // MessageBox.Show("Subject: {0}", message.Subject);
                            listBox1.Items.Add(message.Subject);
                            richTextBox6.Text += message.To + "\n";
                            //  progressBar1.Minimum = 0;
                            // progressBar1.Maximum = i++;
                            //progressBar1.Step = i;
                            // ThreadPool.QueueUserWorkItem(new WaitCallback(AnimateScroll));

                            // here's the change:

                        }
                    }
                    if (comboBox3.Text.Contains("subject"))
                    {
                        if (comboBox1.Text.Contains("gmail.com"))
                        {
                            client.Connect("imap.gmail.com", 993, true);

                            client.Authenticate(comboBox1.Text, comboBox2.Text);
                        }
                        if (comboBox1.Text.Contains("hotmail.com"))
                        {
                            client.Connect("imap-mail.outlook.com", 993, true);

                            client.Authenticate(comboBox1.Text, comboBox2.Text);

                        }

                        if (comboBox1.Text.Contains("yahoo.com"))
                        {
                            client.Connect("imap.mail.yahoo.com", 993, true);
                            client.Authenticate(comboBox1.Text, comboBox2.Text);
                        }
                        if (comboBox1.Text.Contains("aol.com"))
                        {
                            client.Connect("imap.aol.com", 993, true);
                            client.Authenticate(comboBox1.Text, comboBox2.Text);
                        }

                        // Get the first personal namespace and list the toplevel folders under it.
                        /*   var personal = client.GetFolder(client.PersonalNamespaces[0]);
                           foreach (var folder in personal.GetSubfolders(false))
                               MessageBox.Show("[folder] {0}", folder.Name);*/
                        var connected = client.IsAuthenticated;
                        // The Inbox folder is always available on all IMAP servers...

                        var inbox = client.Inbox;
                        label8.Text = inbox.Count.ToString();
                        inbox.Open(FolderAccess.ReadOnly);
                        /*
                        Console.WriteLine("Total messages: {0}", inbox.Count);
                        Console.WriteLine("Recent messages: {0}", inbox.Recent);*/
                        label8.Text = inbox.Count.ToString();
                        for (int i = 0; i < inbox.Count; i++)
                        {
                            this.timer1.Start();
                            var message = inbox.GetMessage(i);
                            // MessageBox.Show("Subject: {0}", message.Subject);
                            listBox1.Items.Add(message.Subject);
                            richTextBox6.Text += message.Subject + "\n";
                            //  progressBar1.Minimum = 0;
                            // progressBar1.Maximum = i++;
                            //progressBar1.Step = i;
                            // ThreadPool.QueueUserWorkItem(new WaitCallback(AnimateScroll));

                            // here's the change:

                        }
                    }
                  
                    client.Disconnect(true);
                    // MessageBox.Show("Done !");
                    progressBar1.Value = 100;

                }





                string sPath = "C:/Users/Public/passandlog.txt";
                if (File.Exists("C:/Users/Public/passandlog.txt"))
                {

                    System.IO.StreamWriter SaveFile = new System.IO.StreamWriter(sPath);
                    foreach (var item in comboBox1.Items)
                    {
                        SaveFile.WriteLine(item);

                    }
                    foreach (var x in comboBox2.Items)
                    {
                        SaveFile.WriteLine(x);
                    }



                    SaveFile.Close();
                }
                else
                {
                    System.IO.StreamWriter SaveFile = new System.IO.StreamWriter(sPath);
                    foreach (var item in comboBox1.Items)
                    {
                        SaveFile.WriteLine(item);
                    }
                    foreach (var item in comboBox2.Items)
                    {
                        SaveFile.WriteLine(item);
                    }

                    SaveFile.Close();
                }
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {

        }

        private void timer4_Tick_1(object sender, EventArgs e)
        {
            label16.Text = TimeZoneInfo.ConvertTimeBySystemTimeZoneId(DateTime.UtcNow, "AUS Eastern Standard Time").ToLongTimeString();
            timer4.Start();
        }

        private void timer5_Tick_1(object sender, EventArgs e)
        {
            label18.Text = TimeZoneInfo.ConvertTimeBySystemTimeZoneId(DateTime.UtcNow, "Atlantic Standard Time").ToLongTimeString();
            timer5.Start();
        }

        private void timer3_Tick_1(object sender, EventArgs e)
        {
            label14.Text = TimeZoneInfo.ConvertTimeBySystemTimeZoneId(DateTime.UtcNow, "Central Europe Standard Time").ToLongTimeString();
            timer3.Start();
        }

        private void timer2_Tick_2(object sender, EventArgs e)
        {
            label12.Text = DateTime.Now.ToLongTimeString();
            timer2.Start();
        }

        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {
            int SerialsCounter = textBox1.Lines.Length;
            label22.Text = SerialsCounter.ToString();
            //TextBox tex = new TextBox();
            textBox1.Focus();
        }

        private void listBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            int SerialsCounter = listBox1.Items.Count;
            label23.Text = SerialsCounter.ToString();
        }

        private void button22_Click(object sender, EventArgs e)
        {

        }

        private void button21_Click(object sender, EventArgs e)
        {

        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }

        private void groupBox4_Enter(object sender, EventArgs e)
        {

        }

        private void button15_Click(object sender, EventArgs e)
        {

        }

        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void button14_Click(object sender, EventArgs e)
        {

        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void button12_Click(object sender, EventArgs e)
        {

        }

        private void label14_Click(object sender, EventArgs e)
        {

        }

        private void label18_Click(object sender, EventArgs e)
        {

        }

        private void label15_Click(object sender, EventArgs e)
        {

        }

        private void label17_Click(object sender, EventArgs e)
        {

        }

        private void label16_Click(object sender, EventArgs e)
        {

        }

        private void label22_Click(object sender, EventArgs e)
        {

        }

        private void label23_Click(object sender, EventArgs e)
        {

        }

        private void button8_Click_1(object sender, EventArgs e)
        {

        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter_1(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void progressBar1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void timer1_Tick_1(object sender, EventArgs e)
        {

        }

        private void label21_Click(object sender, EventArgs e)
        {

        }

        private void label20_Click(object sender, EventArgs e)
        {

        }

        private void label19_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click_1(object sender, EventArgs e)
        {

        }

        private void timer6_Tick(object sender, EventArgs e)
        {

        }

        private void timer7_Tick(object sender, EventArgs e)
        {

        }

        private void timer8_Tick(object sender, EventArgs e)
        {

        }

        private void timer9_Tick(object sender, EventArgs e)
        {

        }

        private void timer10_Tick(object sender, EventArgs e)
        {

        }

        private void button17_Click(object sender, EventArgs e)
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

        private void button16_Click(object sender, EventArgs e)
        {
            richTextBox1.Text = "";
            richTextBox2.Text = "";
            richTextBox3.Text = "";
            richTextBox4.Text = "";
        }

        private void richTextBox5_TextChanged(object sender, EventArgs e)
        {
            int SerialsCounter = richTextBox5.Lines.Length;
            label22.Text = SerialsCounter.ToString();
            //TextBox tex = new TextBox();
            // textBox1.Focus();
        }

        private void richTextBox6_TextChanged(object sender, EventArgs e)
        {
            int SerialsCounter = richTextBox6.Lines.Length;
            label23.Text = SerialsCounter.ToString();
        }

        private void richTextBox7_TextChanged(object sender, EventArgs e)
        {
            int SerialsCounter = richTextBox7.Lines.Length;
            label28.Text = SerialsCounter.ToString();
        }

        private void button18_Click(object sender, EventArgs e)
        {
            /*proc = 0;
         progressBar1.Value = 0;*/
            if (comboBox1.Text == "" && comboBox2.Text == "")
            {
                MessageBox.Show("Please Enter Your Login and Password");
            }
            else
            {
                using (var client = new MailKit.Net.Imap.ImapClient())
                {
                    // For demo-purposes, accept all SSL certificates
                    client.ServerCertificateValidationCallback = (s, c, h, v) => true;

                    if (comboBox1.Text.Contains("gmail.com"))
                    {
                        client.Connect("imap.gmail.com", 993, true);

                        client.Authenticate(comboBox1.Text, comboBox2.Text);
                        // MessageBox.Show("here");
                        var test = client.IsAuthenticated;
                        //  MessageBox.Show(test.ToString());
                        var inbox = client.Inbox;
                        //  label8.Text = inbox.Count.ToString();
                        //var spam = client.GetFolder("[Gmail]/Spam");
                        inbox.Open(FolderAccess.ReadWrite);
                        int ie = inbox.Count;

                        int iz = 0;
                        // MessageBox.Show("" + spam.Count);
                        if (inbox.Count == 0)
                        {
                            MessageBox.Show("Spam Empty");
                        }
                        else
                        {
                            while ((inbox.Count) > 0)
                            {
                                var matchFolder = client.GetFolder("[Gmail]/Important");
                                inbox.MoveTo(iz, matchFolder);
                                // code moving from spam to archive
                                /* var archive = client.GetFolder("inbox/chabbatest");
                        MessageBox.Show(archive.Count.ToString());
                                spam.MoveTo(iz, archive);*/
                                // MessageBox.Show(spam.Count.ToString());

                            }
                            // boucle for 
                            /* for (int i = 0; i < (spam.Count); i++)
                             {
                                 //var messagespam = spam.GetMessage(i);
                                 var matchFolder = client.GetFolder("inbox");
                                 if (i >= 0)
                                 {
                                     spam.MoveTo(i, matchFolder);
                                 }
                               /*  if (i == 0)
                                 {
                                     spam.MoveTo(i, matchFolder);
                                 }*/
                            /*   else
                               {
                                   MessageBox.Show("Nothing to move");
                               }*/


                            //}
                            // end boucle for


                            client.Disconnect(true);
                            MessageBox.Show("all emails was Moved to Inbox Folder!");
                        }


                    }
                    if (comboBox1.Text.Contains("hotmail.com"))
                    {
                        client.Connect("imap-mail.outlook.com", 993, true);

                        client.Authenticate(comboBox1.Text, comboBox2.Text);
                        var spam = client.GetFolder("junk");
                        var inbox = client.Inbox;

                        //spam.Open(FolderAccess.ReadOnly);
                        // spam.Open(FolderAccess.ReadWrite);
                        inbox.Open(FolderAccess.ReadWrite);
                        int ie = spam.Count;
                        int iz = 0;
                        // MessageBox.Show("" + spam.Count);
                        //MessageBox.Show(inbox.Count.ToString());

                        if (inbox.Count == 0)
                        {
                            MessageBox.Show("Inbox Empty");
                        }
                        else
                        {
                            while ((inbox.Count) > 0)
                            {
                                /* var matchFolder = client.GetFolder("inbox");
                                 spam.MoveTo(iz, matchFolder);*/
                                // code moving from spam to archive
                                var archive = client.GetFolder("inbox/test");
                                //  MessageBox.Show(archive.Count.ToString());
                                var matchFolder = client.GetFolder("inbox");
                                matchFolder.MoveTo(iz, archive);

                                // MessageBox.Show(spam.Count.ToString());

                            }
                            // boucle for 
                            /* for (int i = 0; i < (spam.Count); i++)
                             {
                                 //var messagespam = spam.GetMessage(i);
                                 var matchFolder = client.GetFolder("inbox");
                                 if (i >= 0)
                                 {
                                     spam.MoveTo(i, matchFolder);
                                 }
                               /*  if (i == 0)
                                 {
                                     spam.MoveTo(i, matchFolder);
                                 }*/
                            /*   else
                               {
                                   MessageBox.Show("Nothing to move");
                               }*/


                            //}
                            // end boucle for


                            client.Disconnect(true);
                            MessageBox.Show("all emails was Moved to Inbox Folder!");
                        }
                    }
                    if (comboBox1.Text.Contains("yahoo.com"))
                    {
                        client.Connect("imap.mail.yahoo.com", 993, true);

                        client.Authenticate(comboBox1.Text, comboBox2.Text);
                        var spam = client.GetFolder("Bulk Mail");
                        //spam.Open(FolderAccess.ReadOnly);
                        spam.Open(FolderAccess.ReadWrite);
                        int ie = spam.Count;
                        //  int iz = 0;
                        MessageBox.Show("" + spam.Count);
                        for (int i = 0; i < (spam.Count); i++)
                        {
                            //var messagespam = spam.GetMessage(i);
                            var matchFolder = client.GetFolder("inbox");
                            if (i >= 0)
                            {
                                spam.MoveTo(i, matchFolder);
                            }
                            if (i == 0)
                            {
                                spam.MoveTo(i, matchFolder);

                                //  spam.MoveTo(IList<UniqueId> uids,matchFolder,CancellationToken cancellationToken = null);

                            }
                            else
                            {
                                MessageBox.Show("Nothing to move");
                            }


                        }


                        client.Disconnect(true);
                        MessageBox.Show("all emails was Moved to Inbox Folder!");
                    }


                }


                //MessageBox.Show("Sorry !! this Operation Still On Procces");
            }
        }

        private void comboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            // if(comboBox1.sele
            bool tr = comboBox1.Items.Contains("job.test@yahoo.com");
            // MessageBox.Show(tr.ToString());
            /* if (tr== true)
             {
              
                 button18.Visible = false;
             }*/

        }

        private void button19_Click(object sender, EventArgs e)
        {


            using (var sfd = new SaveFileDialog())
            {
                sfd.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
                sfd.FilterIndex = 2;

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    File.WriteAllText(sfd.FileName, "");
                }
            }
          /*  string path = @"C:/Users/Public/warmuptest.txt";
            if (!File.Exists(path))
            {
                // Create a file to write to.
                using (StreamWriter sw = File.CreateText(path))
                {

                }
            }*/


        }
        private static ManualResetEvent connectDone = new ManualResetEvent(false);
        private static void ConnectCallback(IAsyncResult ar)
        {
            Socket client = (Socket)ar.AsyncState;

            // Complete the connection.
            client.EndConnect(ar);

            // Signal that the connection has been made.
            connectDone.Set();
        }


        private void button20_Click(object sender, EventArgs e)
        {
            using (var client = new MailKit.Net.Imap.ImapClient())
            {
                // For demo-purposes, accept all SSL certificates
                client.ServerCertificateValidationCallback = (s, c, h, v) => true;


                //   client.Connect(socket, "outlook.office365.com", 995, SecureSocketOptions.SslOnConnect);

                /* System.Net.Sockets.TcpClient clientSocket = new System.Net.Sockets.TcpClient();
             clientSocket.Connect("46.105.168.248", 3838);*/
                //   string host = "188.165.250.78";
                Socket sock = ConnectSocket("188.165.250.78", 3836);
                WebClient webcli = new WebClient();


                /* TcpListener server = new TcpListener(IPAddress.Parse("188.165.250.78"),3836);
                 server.Start();
                 MessageBox.Show("Server has started on 127.0.0.1:80.{0}Waiting for a connection...", Environment.NewLine);

                 TcpClient client4 = server.AcceptTcpClient();*/

                MessageBox.Show("A client connected.");
                // client.Connect("imap-mail.outlook.com", 993, true);
                //saadsocket.BeginConnect(new IPEndPoint(IPAddress.Parse("188.165.250.78"),3836), new AsyncCallback(ConnectCallback), null);
                //client.Connect("imap-mail.outlook.com", 993, SecureSocketOptions.SslOnConnect);

                // IPAddress ipAddress;
                // ipAddress = IPAddress.Parse("188.165.250.78");
                //    IPAddress ipaddress = IPAddress.Parse("188.165.250.78");
                // IPEndPoint localEndp = new IPEndPoint(ipAddress,3836);
                //   var uri = new Uri("imaps://imap-mail.outlook.com:993");
                // proxy test
                /*   HttpWebRequest request = (HttpWebRequest)WebRequest.Create(@"http://www.outlook.com");
                   WebProxy myproxy = new WebProxy("178.33.154.173",3838);
                  myproxy.BypassProxyOnLocal = false;
                  request.Proxy = myproxy;
                 request.Method = "GET";
                  HttpWebResponse response = (HttpWebResponse) request.GetResponse();
                  MessageBox.Show(response.IsMutuallyAuthenticated.ToString());
                  if (myproxy.Address != null)
                  {
                      
                  }*/
                string ipproxy = "178.33.154.173";
                int port = 3838;
                Uri url = new Uri("http://www.outlook.com");
                IWebProxy proxy = new WebProxy(ipproxy, port);
                HttpWebRequest myWebRequest = (HttpWebRequest)WebRequest.Create(url);
                myWebRequest.Proxy = proxy;
                //  MessageBox.Show(CanPing("178.33.154.173").ToString());
                /* HttpWebRequest request = WebRequest.Create(postUrl) as HttpWebRequest;
                 request.Proxy = new WebProxy(MyProxyHostString, MyProxyPort);*/
                // WebProxy myproxy = new WebProxy("178.33.154.173", 3838);
                //  byte[] data;






                System.Net.IPAddress ipaddress = System.Net.IPAddress.Parse("188.165.250.78");  //127.0.0.1 as an example
                Socket ssocket = new Socket(ipaddress.AddressFamily, SocketType.Stream, ProtocolType.Tcp);
                ssocket.Connect(ipaddress, 3836);

                client.Connect(ssocket, "imap-mail.outlook.com", 993, SecureSocketOptions.SslOnConnect);
                client.Authenticate("daciaall2016@hotmail.com", "masteralex1994");
                var spam = client.GetFolder("junk");


                //spam.Open(FolderAccess.ReadOnly);
                spam.Open(FolderAccess.ReadWrite);
                int ie = spam.Count;
                int iz = 0;
                // MessageBox.Show("" + spam.Count);
                if (spam.Count == 0)
                {
                    MessageBox.Show("spam Empty");
                }
                else
                {
                    while ((spam.Count) > 0)
                    {
                        var matchFolder = client.GetFolder("inbox");
                        spam.AddFlags(iz, MessageFlags.Seen, true);

                        spam.MoveTo(iz, matchFolder);

                    }



                    client.Disconnect(true);
                    MessageBox.Show("all emails was Moved to Inbox Folder!");
                }


            }


        }

        public object ipAddress { get; set; }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            richTextBox7.Clear();



            /* string source = @"First name: Elvis
Last name: Presley
Address: 1 Heaven Street
City: Memphis
State: TN
Zip: 12345
";
           MessageBox.Show(ExtractJist(source));*/
            // string sourceString = "sid [1544764] srv [CFT256] remip [10.0.128.31] fwf []...remip [10.0.128.41] fwf []... ";
            string content = richTextBox5.Text;
            string content2 = Regex.Replace(content, "_", " ");
            string content1 = richTextBox5.Text;
            /*string[] content3 = Regex.Split(content1, @"\W");
               MessageBox.Show(content3[0]);*/

            var matches = Regex.Matches(content2, @"(\d{0,9}\.\d{0,9}\.\d{0,9}\.\d{0,9})");
            Regex regex = new Regex(@"[A-Za-z]+Story[0-9]*");
            var mat = Regex.Matches(richTextBox5.Text, @"[A-Za-z]+[0-9]*");
            //var matches1 = Regex.Match(content3[0],@"^\d");
            foreach (var match in mat)
            {
                // MessageBox.Show(match.ToString());
                //  richTextBox7.Text += richTextBox5.Lines[match].Split('_').First() + "\n";
                /* richTextBox5.Text += match.ToString();
                 richTextBox5.AppendText(Environment.NewLine);*/
                richTextBox7.Text += match.ToString();
                richTextBox7.AppendText(Environment.NewLine);

            }
            /*  for (int n = 0; n < richTextBox5.Lines.Length; ++n)
              {
                  richTextBox7.Text += richTextBox5.Lines[n].Split('_').First() + "\n";
                  //richTextBox7.Text += richTextBox5.Lines[n].Split('_').Last() + "\n";
              }*/

            // var matches = Regex.Matches(content2, @"\b(\d{0,9}\.\d{0,9}\.\d{0,9}\.\d{0,9})\b");
            richTextBox5.Clear();
            foreach (var match in matches)
            {

                richTextBox5.Text += match.ToString();
                richTextBox5.AppendText(Environment.NewLine);

            }

            richTextBox7.Text = string.Join(Environment.NewLine, richTextBox7.Lines.Distinct());
            richTextBox5.Text = string.Join(Environment.NewLine, richTextBox5.Lines.Distinct());
            label22.Text = matches.Count.ToString();
            //label28.Text = matches1.Count.ToString();
            /* string str = "Title = Car Promotion, Model = BMW 323";
             Debug.Write(str.Split('=').Last());*/

            MessageBox.Show("Done!");

            /*  SautinSoft.HtmlToRtf h = new SautinSoft.HtmlToRtf();           
              // string htmlString = @"<b>Hello World!</b>";
              string htmlString = @"\b(\d{0,9}\.\d{0,9}\.\d{0,9}\.\d{0,9})\b";
                 h.OutputFormat = HtmlToRtf.eOutputFormat.TextUnicode;
               string textString = h.ConvertString(htmlString);
               MessageBox.Show(textString);*/
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {

        }

        private void button21_Click_1(object sender, EventArgs e)
        {

            string host = "188.165.250.78";
            //host = Dns.GetHostName();
            string result = SocketSendReceive(host, 3836);

            MessageBox.Show(result);
        }
        //method connect socket
        private static Socket ConnectSocket(string server, int port)
        {
            Socket s = null;
            IPHostEntry hostEntry = null;

            // Get host related information.
            hostEntry = Dns.GetHostEntry(server);

            // Loop through the AddressList to obtain the supported AddressFamily. This is to avoid
            // an exception that occurs when the host IP Address is not compatible with the address family
            // (typical in the IPv6 case).
            foreach (IPAddress address in hostEntry.AddressList)
            {
                IPEndPoint ipe = new IPEndPoint(address, port);
                Socket tempSocket =
                    new Socket(ipe.AddressFamily, SocketType.Stream, ProtocolType.Tcp);

                tempSocket.Connect(ipe);

                if (tempSocket.Connected)
                {
                    s = tempSocket;
                    break;
                }
                else
                {
                    continue;
                }
            }
            return s;
        }

        // method socketsendrecive
        private static string SocketSendReceive(string server, int port)
        {
            string request = "GET / HTTP/1.1\r\nHost: " + server +
                "\r\nConnection: Close\r\n\r\n";
            Byte[] bytesSent = Encoding.ASCII.GetBytes(request);
            Byte[] bytesReceived = new Byte[256];

            // Create a socket connection with the specified server and port.
            Socket s = ConnectSocket(server, port);

            if (s == null)
                return ("Connection failed");

            // Send request to the server.
            s.Send(bytesSent, bytesSent.Length, 0);

            // Receive the server home page content.
            int bytes = 0;
            string page = "Default HTML page on " + server + ":\r\n";

            // The following will block until te page is transmitted.
            do
            {
                bytes = s.Receive(bytesReceived, bytesReceived.Length, 0);
                page = page + Encoding.ASCII.GetString(bytesReceived, 0, bytes);
            }
            while (bytes > 0);

            return page;
        }
        public string versionofmail = "";
        public static string BaseUrl = "https://outlook.live.com/owa/";
        private void button22_Click_1(object sender, EventArgs e)
        {
            //*****************************
            /* int inc = 0;
             ImageList imgs = new ImageList();
             imgs.ImageSize = new Size(20, 20);
             string[] paths = { };
             paths = Directory.GetFiles("C:/Users/Public/imgs");
             foreach (string path in paths)
             {
                 imgs.Images.Add(Image.FromFile(path));
             }*/


            using (StreamReader sr = new StreamReader("C:/Users/Public/warmuptest.txt"))
            {
                while (sr.Peek() >= 0)
                {
                    string str;
                    string[] strArray;
                    str = sr.ReadLine();

                    strArray = str.Split(',');

                    string ips = strArray[0];
                    string boite = strArray[1];
                    string pass = strArray[2];

                    var chromeOptions = new ChromeOptions();
                    //Create a new proxy object
                    var proxy = new Proxy();
                    //Set the http proxy value, host and port.
                    MessageBox.Show(ips.ToString());
                    chromeOptions.AddArgument("--lang=fr");
                    proxy.HttpProxy = ips;
                    //Set the proxy to the Chrome options
                    chromeOptions.Proxy = proxy;
                    // chromeOptions.AddArgument("--proxy-server=38.89.140.134:99");
                    //Then create a new ChromeDriver passing in the options
                    //ChromeDriver path isn't required if its on your path
                    //If it now downloaded it and put the path here

                    var driver1 = new ChromeDriver(@"C:\webdrivers", chromeOptions);
                    //***********work without proxy******************
                    // var driver1 = new ChromeDriver(@"C:\webdrivers");

                    //Navigation to a url and a look at the traffic logged in fiddler

                    driver1.Navigate().GoToUrl(BaseUrl);
                    driver1.FindElementByLinkText("Se connecter").Click();


                    var loginBox = driver1.FindElement(By.XPath("//input[contains(text(),'')]"));
                    System.Threading.Thread.Sleep(2500);
                    loginBox.SendKeys(boite);
                    driver1.FindElement(By.XPath("//input[@type='submit'][@value='Suivant']")).Click();
                    //MessageBox.Show(btn.Text);
                    System.Threading.Thread.Sleep(2500);
                    // driver1.Manage().Window.Maximize();
                    var passwordBox = driver1.FindElement(By.XPath("//input[@id='i0118' and contains(text(),'')]"));
                    // MessageBox.Show(passwordBox.Text);
                    passwordBox.SendKeys(pass);
                    System.Threading.Thread.Sleep(2500);
                    driver1.FindElement(By.XPath("//input[@type='submit'][@value='Se connecter']")).Click();

                    System.Threading.Thread.Sleep(2500);

                    using (var client = new MailKit.Net.Imap.ImapClient())
                    {
                        // For demo-purposes, accept all SSL certificates
                        client.ServerCertificateValidationCallback = (s, c, h, v) => true;
                        client.Connect("imap-mail.outlook.com", 993, SecureSocketOptions.SslOnConnect);
                        client.Authenticate(boite, pass);
                        var spam = client.GetFolder("junk");
                        //spam.Open(FolderAccess.ReadOnly);
                        spam.Open(FolderAccess.ReadOnly);
                        int ie = spam.Count;
                        int iz = 0;
                        // MessageBox.Show("" + spam.Count);
                        if (spam.Count == 0)
                        {
                            //MessageBox.Show("spam Empty");

                            //listView1.SmallImageList = imgs;
                            //   listView1.Items.Add(boite + " Spam Empty", 0);

                        }
                        else
                        {
                            while ((spam.Count) > 0)
                            {
                                //************watch if the mailbox new version or old version*****************
                                driver1.Manage().Window.Maximize();

                                System.Threading.Thread.Sleep(2000);
                                Boolean visible1 = false;  // assume it is invisible 
                                try
                                {
                                    // try to find the element 
                                    var boiteblocked1 = driver1.FindElement(By.XPath("//span[contains(@class,'label')][contains(text(),'Essayer la version bêta')]")).Displayed;
                                    // if I get to here the element exists 
                                    // if it is visible 
                                    if (boiteblocked1)
                                        visible1 = true;
                                    string newboite = driver1.FindElement(By.XPath("//span[contains(@class,'label')][contains(text(),'Essayer la version bêta')]")).Text.ToString();
                                    // MessageBox.Show(newboite);
                                    //start program
                                    if (newboite == "Essayer la version bêta")
                                    {
                                        /*System.Threading.Thread.Sleep(2000);

                                        driver1.FindElement(By.XPath("//span[contains(@class,'_db_m2')][contains(text(),'Filtrer')]")).Click();
                                        //  driver1.FindElementByName("Afficher la boîte de réception Prioritaire").Click();
                                        driver1.FindElement(By.XPath("//div[contains(@class,'_db_m2')][contains(text(),'Afficher la boîte de réception Prioritaire')]")).Click();


                                        //System.Threading.Thread.Sleep(2000);
                                        //driver1.FindElement(By.XPath("//div[@class='_fce_F']/div[@class='_fce_G']/span[contains(@class,'owaimg _fce_H')][contains(text(),'Afficher la boîte de réception Prioritaire')]")).Click();
                                        MessageBox.Show("hey");

                                        /*  IWebElement ddl = driver1.FindElement(By.XPath("//span[contains(@class,'_db_m2')][contains(text(),'Filtrer')]"));
                                          ddl.Click();
                                          IList<IWebElement> lis = driver1.FindElements(By.ClassName("_fce_z ms-font-s ms-fwt-sl"));
                                          foreach (IWebElement li in lis)
                                              try
                                              {
                                                  IWebElement checkBox = li.FindElement(By.ClassName("o365button _fce_B menuItemWithCheckMark ms-fwt-r ms-font-s ms-fwt-sb ms-fcl-nd ms-bgc-tl"));
                                                  //if (checkBox.Selected)
                                                  checkBox.Click();
                                                  break;
                                              }
                                              catch { }
                                          */

                                        // listView1.SmallImageList = imgs;
                                        // listView1.Items.Add(boite + " Done", 1);
                                        /* driver1.Close();
                                         driver1.Quit();*/
                                        //  MessageBox.Show("add contact");

                                    }
                                }
                                catch (Exception g)
                                {
                                    //  Assert.IsFalse(visible1);
                                }






                                client.Disconnect(true);
                            }
                        }



                    }


                }

            }

            //**************************************
            string[] linetest = File.ReadAllLines("C:/Users/Public/warmuptest.txt");
            int count = 0;
            //   foreach (string line in linetest)

            {
                string[] parts = linetest[count].Split(',');
                //MessageBox.Show(line);
                foreach (string part in parts)
                {
                    //MessageBox.Show(part);
                    string ipetport = parts[0];
                    MessageBox.Show(ipetport);
                    string email1 = parts[1];
                    MessageBox.Show(email1);
                    string pass1 = parts[2];
                    MessageBox.Show(pass1);
                }
                count++;
            }





            //    foreach (string line in File.ReadAllLines("C:/Users/Public/warmuptest.txt"))
            {
                /*  progressBar2.Minimum = 0;
                  progressBar2.Maximum = inc;*/
                // string[] backtoligne = line.Split('\n');

                int i = 0;

                System.IO.StreamReader file =
    new System.IO.StreamReader(path);
                string linee;
                while ((linee = file.ReadLine()) != null)
                {
                    MessageBox.Show(linee);
                    string[] parts = linee.Split(',');
                //    listBox2.Items.Add(parts[1] + "_____Done");


                    foreach (string part in parts)
                    {
                        // MessageBox.Show("{0}:{1}",part.ToString());
                        //test 

                   //     progressBar2.Visible = true;
                        string ipetport = parts[0];

                        string email1 = parts[1];
                        string pass1 = parts[2];
                        //MessageBox.Show(email1.ToString());
                        // label5.Text = email1;
                        // label8.Text = inc.ToString();
                        //  label6.Text = ipetport;
                        //  progressBar2.Value = inc;
                        var chromeOptions = new ChromeOptions();
                        //Create a new proxy object
                        var proxy = new Proxy();
                        //Set the http proxy value, host and port.
                        MessageBox.Show(ipetport.ToString());
                        proxy.HttpProxy = ipetport;
                        //Set the proxy to the Chrome options
                        chromeOptions.Proxy = proxy;
                        // chromeOptions.AddArgument("--proxy-server=38.89.140.134:99");
                        //Then create a new ChromeDriver passing in the options
                        //ChromeDriver path isn't required if its on your path
                        //If it now downloaded it and put the path here

                        var driver1 = new ChromeDriver(@"C:\webdrivers", chromeOptions);

                        // var driver1 = new ChromeDriver(@"C:\webdrivers");

                        //Navigation to a url and a look at the traffic logged in fiddler

                        driver1.Navigate().GoToUrl(BaseUrl);
                        //   IWebElement query = driver1.FindElement(By.Id("Email"));
                        //driver1.FindElementByClassName("linkButtonFixedHeader.office-signIn").Click();
                        // driver1.FindElement(By.ClassName("linkButtonFixedHeader office-signIn")).Click();

                        // driver1.FindElement(By.ClassName("linkButtonSigninHeader")).Click();
                        driver1.FindElementByLinkText("Se connecter").Click();


                        var loginBox = driver1.FindElement(By.XPath("//input[contains(text(),'')]"));
                        System.Threading.Thread.Sleep(4000);
                        loginBox.SendKeys(email1);
                        driver1.FindElement(By.XPath("//input[@type='submit'][@value='Suivant']")).Click();
                        //MessageBox.Show(btn.Text);
                        System.Threading.Thread.Sleep(3000);
                        // driver1.Manage().Window.Maximize();
                        var passwordBox = driver1.FindElement(By.XPath("//input[@id='i0118' and contains(text(),'')]"));
                        // MessageBox.Show(passwordBox.Text);
                        passwordBox.SendKeys(pass1);
                        System.Threading.Thread.Sleep(7000);
                        driver1.FindElement(By.XPath("//input[@type='submit'][@value='Se connecter']")).Click();

                        System.Threading.Thread.Sleep(7000);

                        using (var client = new MailKit.Net.Imap.ImapClient())
                        {
                            // For demo-purposes, accept all SSL certificates
                            client.ServerCertificateValidationCallback = (s, c, h, v) => true;



                            client.Connect("imap-mail.outlook.com", 993, SecureSocketOptions.SslOnConnect);
                            client.Authenticate(email1, pass1);
                            var spam = client.GetFolder("junk");


                            //spam.Open(FolderAccess.ReadOnly);
                            spam.Open(FolderAccess.ReadOnly);
                            int ie = spam.Count;
                            int iz = 0;
                            // MessageBox.Show("" + spam.Count);
                            if (spam.Count == 0)
                            {
                                //MessageBox.Show("spam Empty");
                              //  listBox2.Items.Add(email1 + "___Spam Empty");
                            }
                            else
                            {
                                while ((spam.Count) > 0)
                                {
                                    // *********************code for older version with add contact
                                    driver1.FindElement(By.XPath("//span[contains(@class,'_n_A4')][contains(text(),'Courrier indésirable')]")).Click();
                                    System.Threading.Thread.Sleep(2000);
                                    driver1.FindElement(By.ClassName("_lvv_r1")).Click();
                                    System.Threading.Thread.Sleep(2000);
                                    driver1.Manage().Window.Maximize();
                                    System.Threading.Thread.Sleep(2000);
                                    driver1.FindElement(By.XPath("//button[@title='Déplacer un message légitime dans la boîte de réception']")).Click();
                                    System.Threading.Thread.Sleep(2000);
                                    driver1.FindElement(By.XPath("//span[contains(@class,'_n_A4')][contains(text(),'Boîte de réception')]")).Click();
                                    System.Threading.Thread.Sleep(2000);
                                    driver1.FindElement(By.ClassName("_lvv_2")).Click();
                                    System.Threading.Thread.Sleep(2000);
                                    driver1.FindElement(By.ClassName("_pe_l")).Click();
                                    System.Threading.Thread.Sleep(2000);
                                    driver1.FindElement(By.ClassName("_pf_d")).Click();
                                    System.Threading.Thread.Sleep(2000);
                                    driver1.FindElement(By.ClassName("_fce_G")).Click();
                                    System.Threading.Thread.Sleep(2000);
                                    driver1.FindElement(By.XPath("//button[@title='Enregistrer la modification du contact']")).Click();
                                  //  listBox2.Items.Add(email1 + "_____Done");
                                    // driver1.Dispose();

                                }



                                client.Disconnect(true);

                            }


                        }

                        try
                        {
                        }
                        catch (System.Exception ep)
                        {
                            Console.WriteLine(ep.Message);
                        }
                        //}
                    }
                    //  inc++;
                    i++;
                    // For demonstration.
                }
            }
            IWebDriver driver = null;
            // driver = new ChromeDriver(@"C:\Program Files (x86)\Google\Chrome\Application");

            //set the url of the website you want to navigate to.
            //  driver.Url = "http://www.outlook.com";

            // driver.Manage().Window.Minimize();

            //use navigate method to go to particular URL.
            /*  driver.Navigate();
              driver.Navigate().Refresh();*/

            // var driver1 = new ChromeDriver(@"C:\webdrivers");


            // using proxy selnium




            // MessageBox.Show(nameinbox.ToString());
            // driver1.FindElement(By.XPath("//span[contains(@class,'_n_A4 ms-font-m ms-fwt-sl ms-fcl-np ms-font-weight-semibold')][contains(text(),'Courrier indésirable')]")).Click();
            //span[contains(@class,'_n_A4 ms-font-m ms-fwt-sl ms-fcl-np ms-font-weight-semibold')][contains(text(),'Courrier indésirable')]
            //input[@id='_ariaId_34.folder' and contains(text(),'')]
            //  driver.Quit();


        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            int SerialsCounter = richTextBox1.Lines.Length;
            label31.Text = SerialsCounter.ToString();
            string resultString = Regex.Replace(richTextBox1.Text, @"^\s+$[\r\n]*", string.Empty, RegexOptions.Multiline);


        }

        private void richTextBox2_TextChanged(object sender, EventArgs e)
        {
            int SerialsCounter = richTextBox2.Lines.Length;
            label32.Text = SerialsCounter.ToString();
        }

        private void richTextBox3_TextChanged(object sender, EventArgs e)
        {
            int SerialsCounter = richTextBox3.Lines.Length;
            label33.Text = SerialsCounter.ToString();
        }

        private void richTextBox4_TextChanged(object sender, EventArgs e)
        {
            int SerialsCounter = richTextBox4.Lines.Length;
            label34.Text = SerialsCounter.ToString();
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button21_Click_2(object sender, EventArgs e)
        {
            //*****************************
            int incr = 0;
            using (StreamReader sr = new StreamReader("C:/Users/Public/warmuptest.txt"))
            {
                while (sr.Peek() >= 0)
                {
                    string str;
                    string[] strArray;
                    str = sr.ReadLine();

                    strArray = str.Split(',');

                    string ips = strArray[0];
                    string boite = strArray[1];
                    string pass = strArray[2];
                    //*****************************************************

                    //****************************************
                    var chromeOptions = new ChromeOptions();
                    //Create a new proxy object
                    var proxy = new Proxy();
                    //Set the http proxy value, host and port.
                    MessageBox.Show(ips.ToString());
                    chromeOptions.AddArgument("--lang=fr");
                    proxy.HttpProxy = ips;
                    //Set the proxy to the Chrome options
                    chromeOptions.Proxy = proxy;

                    var driver1 = new ChromeDriver(@"C:\webdrivers", chromeOptions);
                    //***********work without proxy******************
                    // var driver1 = new ChromeDriver(@"C:\webdrivers");
                    //Navigation to a url and a look at the traffic logged in fiddler
                    driver1.Navigate().GoToUrl(BaseUrl);
                    driver1.FindElementByLinkText("Se connecter").Click();
                    var loginBox = driver1.FindElement(By.XPath("//input[contains(text(),'')]"));
                    System.Threading.Thread.Sleep(2500);
                    loginBox.SendKeys(boite);
                    driver1.FindElement(By.XPath("//input[@type='submit'][@value='Suivant']")).Click();
                    System.Threading.Thread.Sleep(2500);
                    var passwordBox = driver1.FindElement(By.XPath("//input[@id='i0118' and contains(text(),'')]"));
                    passwordBox.SendKeys(pass);
                    System.Threading.Thread.Sleep(2500);
                    driver1.FindElement(By.XPath("//input[@type='submit'][@value='Se connecter']")).Click();
                    driver1.Manage().Window.Maximize();
                    // code for know if page are blocked or know *************************
                    Boolean visible = false;
                    // 
                    /*  string txtblock1 = driver1.FindElement(By.XPath("//span[contains(@class,'o365cs-nav-brandingText')][contains(text(),'Courrier Outlook')]")).Text.ToString();
                      MessageBox.Show(txtblock1);
                      if (txtblock1 == "Courrier Outlook")
                      {
                          MessageBox.Show("saad");
                      }*/
                    try
                    {
                        // try to find the element 
                        var boiteblocked = driver1.FindElement(By.ClassName("text-title")).Displayed;
                        if (boiteblocked)
                            visible = true;
                        string txtblock3 = driver1.FindElement(By.ClassName("text-title")).Text.ToString();
                        if (txtblock3 == "Aidez-nous à protéger votre compte")
                        {
                            //  listView1.Items.Add(boite + " Boite Blocked");
                            continue;
                        }
                        else continue;
                    }
                    catch (Exception g)
                    {
                    }

                    // try for older boite
                    Boolean visible1 = false;
                    try
                    {
                        // try to find the element 
                        var boiteblocked = driver1.FindElement(By.XPath("//span[contains(@class,'o365cs-nav-brandingText')][contains(text(),'Courrier Outlook')]")).Displayed;
                        if (boiteblocked)
                            visible1 = true;
                        string txtblock4 = driver1.FindElement(By.XPath("//span[contains(@class,'o365cs-nav-brandingText')][contains(text(),'Courrier Outlook')]")).Text.ToString();
                        MessageBox.Show(txtblock4);
                        if (txtblock4 == "Courrier Outlook")
                        {
                            versionofmail = "olderversion";

                        }

                    }
                    catch (Exception g)
                    {
                    }
                    // try for new boite
                    Boolean visible2 = false;
                    try
                    {
                        // try to find the element 
                        var boiteblocked1 = driver1.FindElement(By.CssSelector("i[data-icon-name='Waffle']")).Displayed;
                        // if I get to here the element exists 
                        // if it is visible 
                        if (boiteblocked1)
                            visible2 = true;
                        string txtblock = driver1.FindElement(By.CssSelector("i[data-icon-name='Waffle']")).Text.ToString();
                        MessageBox.Show(txtblock);

                        if (txtblock == "")
                        {
                            versionofmail = "newversion";
                        }
                    }
                    catch (Exception g)
                    {
                    }
                    // for clk in option to enable prioritaire for new version outlook
                    /*   IWebElement ell = driver1.FindElement(By.CssSelector("i[data-icon-name='Settings']"));
                       //MessageBox.Show(ell.Location.ToString());
                       ell.Click();
                       System.Threading.Thread.Sleep(2000);
                       IJavaScriptExecutor js = (IJavaScriptExecutor)driver1;
                       IWebElement Element = driver1.FindElement(By.XPath("//button[@id='Toggle111']"));
                       //This will scroll the page till the element is found		
                   
                       js.ExecuteScript("arguments[0].scrollIntoView();", Element);

                       driver1.FindElement(By.XPath("//button[@id='Toggle111']")).Click();

                       System.Threading.Thread.Sleep(1000);
                       // close page of option
                       IWebElement ellclose = driver1.FindElement(By.CssSelector("i[data-icon-name='Cancel']"));
                       //MessageBox.Show(ell.Location.ToString());
                       ellclose.Click();*/


                    System.Threading.Thread.Sleep(1000);
                    // etape moving spam to inbox
                    using (var client = new MailKit.Net.Imap.ImapClient())
                    {

                        // For demo-purposes, accept all SSL certificates
                        client.ServerCertificateValidationCallback = (s, c, h, v) => true;
                        client.Connect("imap-mail.outlook.com", 993, SecureSocketOptions.SslOnConnect);
                        client.Authenticate(boite, pass);
                        var spam = client.GetFolder("junk");
                        //spam.Open(FolderAccess.ReadOnly);
                        spam.Open(FolderAccess.ReadOnly);
                        int ie = spam.Count;
                        int iz = 0;

                        // MessageBox.Show("" + spam.Count);
                        if (spam.Count == 0)
                        {
                            //MessageBox.Show("spam Empty");

                            //  listView1.SmallImageList = imgs;
                            // listView1.Items.Add(boite + " Spam Empty");

                        }
                        else
                        {

                            while ((spam.Count) > 0)
                            {
                                //System.Threading.Thread.Sleep(2000);
                                /*  try
                                  {*/
                                // MessageBox.Show("3");
                                switch (versionofmail)
                                {

                                    case "olderversion":
                                        //start program
                                        //MessageBox.Show(versionofmail);
                                        //MessageBox.Show("old version in");
                                        // *********************code for older version with add contact
                                        System.Threading.Thread.Sleep(2000);
                                        driver1.FindElement(By.XPath("//span[contains(@class,'_n_A4')][contains(text(),'Courrier indésirable')]")).Click();
                                        System.Threading.Thread.Sleep(2000);
                                        driver1.FindElement(By.ClassName("_lvv_r1")).Click();
                                        System.Threading.Thread.Sleep(2000);
                                        driver1.Manage().Window.Maximize();
                                        System.Threading.Thread.Sleep(2000);

                                        driver1.FindElementByLinkText("Ceci n’est pas du courrier indésirable").Click();
                                        // driver1.FindElement(By.XPath("//button[@title='Déplacer un message légitime dans la boîte de réception']")).Click();
                                        System.Threading.Thread.Sleep(2000);
                                        driver1.FindElement(By.XPath("//span[contains(@class,'_n_A4')][contains(text(),'Boîte de réception')]")).Click();
                                        System.Threading.Thread.Sleep(2000);
                                        driver1.FindElement(By.ClassName("_lvv_2")).Click();
                                        System.Threading.Thread.Sleep(2000);
                                        driver1.FindElement(By.ClassName("_pe_l")).Click();
                                        System.Threading.Thread.Sleep(2000);
                                        driver1.FindElement(By.ClassName("_pf_d")).Click();
                                        System.Threading.Thread.Sleep(2000);
                                        driver1.FindElement(By.ClassName("_fce_G")).Click();
                                        System.Threading.Thread.Sleep(2000);
                                        driver1.FindElement(By.XPath("//button[@title='Enregistrer la modification du contact']")).Click();
                                        //  listView1.Items.Add(boite + " Done");

                                        break;

                                    case "newversion":

                                        // MessageBox.Show("new version in");

                                        System.Threading.Thread.Sleep(2000);
                                        // Boolean visible1 = false;  // assume it is invisible 
                                        IWebElement firststep = driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']"));
                                        firststep.Click();
                                        //*********************code for new version 
                                        System.Threading.Thread.Sleep(2000);

                                        driver1.FindElement(By.XPath("//span[contains(@class,'_3lQ0EN5N3oGHxkBKF6_Ane')][contains(text(),'Courrier indésirable')]")).Click();
                                        System.Threading.Thread.Sleep(2000);

                                        driver1.FindElement(By.ClassName("_2EHjCdO2IEh-zlrH2jOD50")).Click();

                                        System.Threading.Thread.Sleep(2500);
                                        System.Threading.Thread.Sleep(2500);
                                        driver1.FindElementByLinkText("Courrier légitime").Click();
                                        System.Threading.Thread.Sleep(2500);
                                        //  driver1.FindElement(By.XPath("//span[contains(@class,'_3lQ0EN5N3oGHxkBKF6_Ane')][contains(text(),'Boîte de réception')]")).Click();


                                        break;
                                    //  listView1.Items.Add(boite + " Done");

                                    default:

                                        break;
                                }
                                client.Disconnect(true);
                                //
                                System.Threading.Thread.Sleep(2000);
                                // Boolean visible1 = false;  // assume it is invisible 
                                // add this 18.30

                                // add this 18.30

                                //************watch if the mailbox new version or old version*****************

                                /* driver1.Close();
                                 driver1.Quit();*/
                                //try } //first while 
                            }
                        }





                    }
                    incr++;
                }
            }




            //}
            // }
        }
        //  static public string ips ="";
        private void button23_Click(object sender, EventArgs e)
        {
            //**************************************

            string line;
            int number = 0;
            System.IO.StreamReader file =
    new System.IO.StreamReader(path);

            while ((line = file.ReadLine()) != null)
            {


                // show ligne in file 
                //MessageBox.Show(line);
                string[] lines = line.Split(',');

                string ips = lines[0];
                // MessageBox.Show(lines[0]);
                string boite = lines[1];
                string pass = lines[2];

                number++;


                /* RegistryKey registry = Registry.CurrentUser.OpenSubKey
                ("Software\\Microsoft\\Windows\\CurrentVersion\\Internet Settings", true);
                 registry.SetValue("ProxyEnable", 1);
                 registry.SetValue
                 ("ProxyServer",ips);
                 if ((int)registry.GetValue("ProxyEnable", 1) == 1)
                     MessageBox.Show(" enable the proxy.");
                /* else
                     MessageBox.Show("The proxy has been turned on.");*/

                //**************************************

                var chromeOptions = new ChromeOptions();
                //Create a new proxy object
                var proxy = new Proxy();
                //Set the http proxy value, host and port.
                // MessageBox.Show(ips.ToString());
                chromeOptions.AddArgument("--lang=en-ca");
                // proxy.HttpProxy = ips;
                //Set the proxy to the Chrome options
                //  chromeOptions.Proxy = proxy;
                // chromeOptions.AddArguments("--proxy-server=" + ips);
                var driver1 = new ChromeDriver(@"C:\webdrivers", chromeOptions);
                //***********work without proxy******************
                //var driver1 = new ChromeDriver(@"C:\webdrivers");
                //Navigation to a url and a look at the traffic logged in fiddler
                driver1.Navigate().GoToUrl(BaseUrl);
                driver1.FindElementByLinkText("Se connecter").Click();
                var loginBox = driver1.FindElement(By.XPath("//input[contains(text(),'')]"));
                System.Threading.Thread.Sleep(2500);
                loginBox.SendKeys(boite);
                driver1.FindElement(By.XPath("//input[@type='submit'][@value='Suivant']")).Click();
                System.Threading.Thread.Sleep(2500);
                var passwordBox = driver1.FindElement(By.XPath("//input[@id='i0118' and contains(text(),'')]"));
                passwordBox.SendKeys(pass);
                System.Threading.Thread.Sleep(2500);
                driver1.FindElement(By.XPath("//input[@type='submit'][@value='Se connecter']")).Click();
                driver1.Manage().Window.Maximize();

                // code for know if page are blocked or know *************************
                Boolean visible = false;
                // 
                /*  string txtblock1 = driver1.FindElement(By.XPath("//span[contains(@class,'o365cs-nav-brandingText')][contains(text(),'Courrier Outlook')]")).Text.ToString();
                  MessageBox.Show(txtblock1);
                  if (txtblock1 == "Courrier Outlook")
                  {
                      MessageBox.Show("saad");
                  }*/
                try
                {
                    // try to find the element 
                    var boiteblocked = driver1.FindElement(By.ClassName("text-title")).Displayed;
                    if (boiteblocked)
                        visible = true;
                    string txtblock3 = driver1.FindElement(By.ClassName("text-title")).Text.ToString();
                    if (txtblock3 == "Aidez-nous à protéger votre compte")
                    {
                        //listBox3.Items.Add(ips + "," + boite + "," + pass + " : Boite Blocked");
                        System.Threading.Thread.Sleep(1000);
                        driver1.Close();
                        driver1.Dispose();
                        continue;
                    }
                    if (txtblock3 == "Votre compte a été temporairement bloqué")
                    {

                        //listBox3.Items.Add(ips + "," + boite + "," + pass + " : Boite Blocked");
                        driver1.FindElement(By.XPath("//input[@type='submit'][@value='Continuer']")).Click();
                        continue;
                    }
                    else continue;
                }
                catch (Exception g)
                {
                }


                // try for older boite
                Boolean visible1 = false;
                try
                {
                    // try to find the element 
                    var boiteblocked = driver1.FindElement(By.XPath("//span[contains(@class,'o365cs-nav-brandingText')][contains(text(),'Courrier Outlook')]")).Displayed;
                    if (boiteblocked)
                        visible1 = true;
                    string txtblock4 = driver1.FindElement(By.XPath("//span[contains(@class,'o365cs-nav-brandingText')][contains(text(),'Courrier Outlook')]")).Text.ToString();
                    MessageBox.Show(txtblock4);
                    if (txtblock4 == "Courrier Outlook")
                    {
                        versionofmail = "olderversion";

                    }

                }
                catch (Exception g)
                {
                }
                // try for new boite
                Boolean visible2 = false;
                try
                {
                    // try to find the element 
                    System.Threading.Thread.Sleep(2000);
                    var boiteblocked1 = driver1.FindElement(By.CssSelector("i[data-icon-name='Waffle']")).Displayed;
                    // if I get to here the element exists 
                    // if it is visible 
                    if (boiteblocked1)
                        visible2 = true;
                    string txtblock = driver1.FindElement(By.CssSelector("i[data-icon-name='Waffle']")).Text.ToString();
                    //  MessageBox.Show(txtblock);

                    if (txtblock == "")
                    {
                        versionofmail = "newversion";
                    }
                }
                catch (Exception g)
                {
                }

                // for clk in option to enable prioritaire for new version outlook
                /*   IWebElement ell = driver1.FindElement(By.CssSelector("i[data-icon-name='Settings']"));
                   //MessageBox.Show(ell.Location.ToString());
                   ell.Click();
                   System.Threading.Thread.Sleep(2000);
                   IJavaScriptExecutor js = (IJavaScriptExecutor)driver1;
                   IWebElement Element = driver1.FindElement(By.XPath("//button[@id='Toggle111']"));
                   //This will scroll the page till the element is found		
                   
                   js.ExecuteScript("arguments[0].scrollIntoView();", Element);

                   driver1.FindElement(By.XPath("//button[@id='Toggle111']")).Click();

                   System.Threading.Thread.Sleep(1000);
                   // close page of option
                   IWebElement ellclose = driver1.FindElement(By.CssSelector("i[data-icon-name='Cancel']"));
                   //MessageBox.Show(ell.Location.ToString());
                   ellclose.Click();*/


                System.Threading.Thread.Sleep(1000);
                // etape moving spam to inbox
                using (var client = new MailKit.Net.Imap.ImapClient())
                {

                    // For demo-purposes, accept all SSL certificates
                    client.ServerCertificateValidationCallback = (s, c, h, v) => true;
                    client.Connect("imap-mail.outlook.com", 993, SecureSocketOptions.SslOnConnect);
                    client.Authenticate(boite, pass);
                    var spam = client.GetFolder("junk");
                    //spam.Open(FolderAccess.ReadOnly);
                    spam.Open(FolderAccess.ReadOnly);
                    int ie = spam.Count;
                    int iz = 0;

                    // MessageBox.Show("" + spam.Count);
                    if (spam.Count == 0)
                    {
                        //listView1.Items.Add(boite + " Spam Empty");
                        System.Threading.Thread.Sleep(1000);
                        driver1.Close();
                        driver1.Dispose();
                    }
                    else
                    {
                        Boolean visible6 = false;
                        try
                        {
                            System.Threading.Thread.Sleep(2000);
                            IWebElement firststep = driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']"));
                            firststep.Click();
                            driver1.FindElement(By.XPath("//span[contains(@class,'_3lQ0EN5N3oGHxkBKF6_Ane')][contains(text(),'Courrier indésirable')]")).Click();

                        }
                        catch
                        {
                        }
                        int startwp = spam.Count();

                        do
                        {
                            //startwp=spam.Count(); 
                            System.Threading.Thread.Sleep(2000);
                            //MessageBox.Show(startwp.ToString());

                            // new boite version 
                            try
                            {

                                // driver1.FindElement(By.XPath("//span[contains(@class,'_3lQ0EN5N3oGHxkBKF6_Ane')][contains(text(),'Courrier indésirable')]")).Click();
                                var boiteblocked6 = driver1.FindElement(By.ClassName("_2EHjCdO2IEh-zlrH2jOD50")).Displayed;
                                System.Threading.Thread.Sleep(2000);
                                //MessageBox.Show(boiteblocked6.ToString());
                                // Ce dossier est vide
                                if (boiteblocked6)
                                {

                                    // Boolean visible1 = false;  // assume it is invisible 
                                    /* IWebElement firststep = driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']"));
                                     firststep.Click();*/
                                    //*********************code for new version 
                                    System.Threading.Thread.Sleep(2000);
                                    driver1.FindElement(By.XPath("//span[contains(@class,'_3lQ0EN5N3oGHxkBKF6_Ane')][contains(text(),'Courrier indésirable')]")).Click();
                                    System.Threading.Thread.Sleep(2000);
                                    driver1.FindElement(By.ClassName("_2EHjCdO2IEh-zlrH2jOD50")).Click();
                                    System.Threading.Thread.Sleep(2500);
                                    driver1.FindElementByLinkText("Courrier légitime").Click();
                                    System.Threading.Thread.Sleep(2500);
                                    visible6 = false;

                                }
                                else
                                {

                                    visible6 = false;
                                    /* System.Threading.Thread.Sleep(1000);
                                     driver1.Close();
                                     driver1.Dispose();*/
                                    break;

                                    /*var boiteblocked7 = driver1.FindElement(By.XPath("//span[contains(@class,'__25o_l9GBJ9r8kRewkYoBfm _2oemjr1JNV5H9BUBUobbw2')][contains(text(),'Ce dossier est vide')]")).Displayed;

                                    if (boiteblocked7 == true && boiteblocked6 == false)
                                    {
                                        listView1.Items.Add(boite + " Done");
                                        break;
                                    }*/

                                }

                            }
                            catch
                            {
                                /* MessageBox.Show("done2");
                                 break;*/
                            }

                            driver1.FindElement(By.XPath("//span[contains(@class,'_3lQ0EN5N3oGHxkBKF6_Ane')][contains(text(),'Courrier indésirable')]")).Click();
                            System.Threading.Thread.Sleep(2000);

                            /*   else
                                   visible6 = false;
                               break;
                                           //  driver1.FindElement(By.XPath("//span[contains(@class,'_3lQ0EN5N3oGHxkBKF6_Ane')][contains(text(),'Boîte de réception')]")).Click();
                                       }
                           
                   
                      catch { 
                          continue;
                      }  */

                            // new boite version 
                            // old boite version                  
                            /*     try
                                      {
                                          // try to find the element 
                                          MessageBox.Show("start try");
                                          driver1.FindElement(By.XPath("//span[contains(@class,'_n_A4')][contains(text(),'Courrier indésirable')]")).Click();
                                          var boiteblocked6 = driver1.FindElement(By.ClassName("_lvv_r1")).Displayed;
                                          if (boiteblocked6==true)
                                          {
                                              //MessageBox.Show("spam not vide yet");
                                              visible6 = true;
                                              System.Threading.Thread.Sleep(2000);
                                              driver1.FindElement(By.XPath("//span[contains(@class,'_n_A4')][contains(text(),'Courrier indésirable')]")).Click();
                                              System.Threading.Thread.Sleep(2000);
                                              MessageBox.Show("clk");
                                              driver1.FindElement(By.ClassName("_lvv_r1")).Click();
                                              System.Threading.Thread.Sleep(2000);
                                              driver1.Manage().Window.Maximize();
                                              System.Threading.Thread.Sleep(2000);

                                              driver1.FindElementByLinkText("Ceci n’est pas du courrier indésirable").Click();
                                              // driver1.FindElement(By.XPath("//button[@title='Déplacer un message légitime dans la boîte de réception']")).Click();
                                              System.Threading.Thread.Sleep(2000);
                                              driver1.FindElement(By.XPath("//span[contains(@class,'_n_A4')][contains(text(),'Boîte de réception')]")).Click();
                                              System.Threading.Thread.Sleep(2000);
                                              // witout prioraitaire
                                              driver1.FindElement(By.ClassName("_lvv_c")).Click();
                                              // with prioraitarire
                                              //driver1.FindElement(By.ClassName("_lvv_2")).Click();
                                              System.Threading.Thread.Sleep(2000);
                                              driver1.FindElement(By.ClassName("_pe_l")).Click();
                                              System.Threading.Thread.Sleep(2000);
                                              driver1.FindElement(By.ClassName("_pf_d")).Click();
                                              System.Threading.Thread.Sleep(2000);
                                              driver1.FindElement(By.ClassName("_fce_G")).Click();
                                              System.Threading.Thread.Sleep(2000);
                                              driver1.FindElement(By.XPath("//button[@title='Enregistrer la modification du contact']")).Click();
                           
                                          }
                                          else visible6 = false;
                                          break;

                                              }
                                 catch { continue; }    */
                            // end old boite version



                            // MessageBox.Show((startwp).ToString());
                            // client.Disconnect(true);
                            startwp--;
                            if (startwp == 0)
                            {
                                //listView1.Items.Add(boite + " Done");
                            }
                        } while ((startwp) >= 0);
                        // do while

                    }

                }


                // }
                //close of while

            }
            //progressBar2.Value = 100;
            MessageBox.Show("Reporting Done!");

        }

        private void progressBar2_Click(object sender, EventArgs e)
        {

        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            backgroundWorker1.WorkerReportsProgress = true;

            System.IO.StreamReader file =
    new System.IO.StreamReader(@"C:/Users/Public/warmuptest.txt");
            string line;
            int i = 1;
            while ((line = file.ReadLine()) != null)
            {
                // Wait 100 milliseconds.
                Thread.Sleep(100);
                // Report progress.

                backgroundWorker1.ReportProgress(i);
                i++;
            }

        }
        private void backgroundWorker1_ProgressChanged(object sender,
           ProgressChangedEventArgs e)
        {
            // Change the value of the ProgressBar to the BackgroundWorker progress.
            //progressBar2.Value = e.ProgressPercentage;
            // Set the text.
            this.Text = e.ProgressPercentage.ToString();
        }
        // code added

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        static extern bool SetCursorPos(int x, int y);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);

        public const int MOUSEEVENTF_LEFTDOWN = 0x02;
        public const int MOUSEEVENTF_LEFTUP = 0x04;

        //This simulates a left mouse click
        public static void LeftMouseClick(int xpos, int ypos)
        {
            SetCursorPos(xpos, ypos);
            mouse_event(MOUSEEVENTF_LEFTDOWN, xpos, ypos, 0, 0);
            mouse_event(MOUSEEVENTF_LEFTUP, xpos, ypos, 0, 0);
        }

        private void button24_Click(object sender, EventArgs e)
        {
            ImageList imgs = new ImageList();
            imgs.ImageSize = new Size(50, 50);
            String[] pathss = { };
            pathss = Directory.GetFiles("C:/imgs");


            // buttun test sans imap
            string line;
            int number = 0;

            System.IO.StreamReader file =
    new System.IO.StreamReader(path);

            while ((line = file.ReadLine()) != null)
            {
                string[] lines = line.Split(',');

                string ips = lines[0];
                // MessageBox.Show(lines[0]);
                string boite = lines[1];
                string pass = lines[2];
                string useragent = lines[3];
                number++;
                string[] ipall = ips.Split(':');
                int port = Convert.ToInt32(ipall[1]);
                bool tt = PingHost(ipall[0], port);
                if (tt == true)
                {

                    // MessageBox.Show("tt True");


                    //**************************************

                    var chromeOptions = new ChromeOptions();
                    //Create a new proxy object
                    var proxy = new Proxy();
                    //Set the http proxy value, host and port.
                    // proxy.HttpProxy = ips;
                    //Set the proxy to the Chrome options
                    //  chromeOptions.Proxy = proxy;
                    chromeOptions.AddArguments("--disable-notifications");
                   // chromeOptions.AddArgument("--headless");
                       chromeOptions.AddArgument("--user-agent="+useragent);
                    //
                   /* if (checkBox1.Checked)
                    {

                        chromeOptions.AddArgument("--lang=aus");
                    }
                    else
                    {
                        chromeOptions.AddArguments("--proxy-server=" + ips);
                        chromeOptions.AddArgument("--lang=aus");
                    }*/
                    // disable pop ups 
                    /*DesiredCapabilities capabilities = DesiredCapabilities.Chrome();
                    capabilities.SetCapability(ChromeOptions.Capability, chromeOptions);*/
                    //chromeOptions.AddAdditionalCapability("excludeSwitches", "disable-popup-blocking");
                    //chromeOptions.AddArgument("--disable-popup-blocking");

                    var driver1 = new ChromeDriver(@"C:\webdrivers", chromeOptions);
                    //***********work without proxy******************
                    //var driver1 = new ChromeDriver(@"C:\webdrivers");
                    driver1.Navigate().GoToUrl(BaseUrl);
                    try
                    {
                        var erorcode = driver1.FindElement(By.XPath("//div[contains(@class,'error-code')][contains(text(),'ERR_PROXY_CONNECTION_FAILED')]")).Displayed;
                        if (erorcode)
                        {
                            driver1.Close();
                            driver1.Dispose();
                            continue;
                        }

                    }
                    catch
                    {
                    }

                    driver1.FindElementByLinkText("Se connecter").Click();

                    var loginBox = driver1.FindElement(By.XPath("//input[contains(text(),'')]"));
                    System.Threading.Thread.Sleep(2500);
                    loginBox.SendKeys(boite);
                    driver1.FindElement(By.XPath("//input[@type='submit'][@value='Suivant']")).Click();
                    System.Threading.Thread.Sleep(2500);
                    try
                    {
                        var suiv = driver1.FindElement(By.XPath("//input[@type='submit'][@value='Suivant']")).Displayed;
                        if (suiv)
                        {
                            /*  listView1.SmallImageList = imgs;
                              listView1.Items.Add("Michael Carrick", 1);*/
                            //listBox4.Items.Add(ips + "," + boite + "," + pass + " : User Inccorect ! Please Try Again");
                            // listView1.Items.Add(boite + "User Inccorect ! Please Try Again");
                            driver1.Close();
                            driver1.Dispose();
                            continue;
                        }

                    }
                    catch { }

                    var passwordBox = driver1.FindElement(By.XPath("//input[@id='i0118' and contains(text(),'')]"));
                    passwordBox.SendKeys(pass);
                    System.Threading.Thread.Sleep(2500);
                    driver1.FindElement(By.XPath("//input[@type='submit'][@value='Se connecter']")).Click();
                    //System.Threading.Thread.Sleep(10000);
                    // driver1.FindElement(By.ClassName("_29AA3OWLpLPsgv7R6xoYtp")).Click();
                    System.Threading.Thread.Sleep(12000);
                    driver1.Manage().Window.Maximize();

                    System.Threading.Thread.Sleep(12000);

                    // MessageBox.Show("10 second");
                    // driver1.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(5);
                    /* IWebElement ele = driver1.FindElement(By.ClassName("_3KAweDTgWHSpd5lcACQdA3"));
                       Actions build = new Actions(driver1);
                       //build.ContextClick(ele).Build().Perform();
                   
                       build.MoveToElement(ele,28,28).Click().Build().Perform();*/

                    try
                    {
                        string df = driver1.FindElement(By.ClassName("_36rwQx61abZ9wJjdE6uJPX")).Text.ToString();

                        //  MessageBox.Show(df);Sans doute pas
                        if (df == "Not likely")
                        {
                            /*  driver1.FindElement(By.ClassName("_3GrIR9uD5dhhMZHdiJ2GC9")).Click();
                              //driver1.FindElement(By.ClassName("ms-Button-icon")).Click();
                              System.Threading.Thread.Sleep(2500);
                              var msgtosent = driver1.FindElement(By.XPath("//textarea[contains(text(),'')]"));
                              System.Threading.Thread.Sleep(2500);
                              msgtosent.SendKeys("good");
                               System.Threading.Thread.Sleep(2500);
                            // driver1.FindElement(By.XPath("//div[contains(@class,'ms-Button-label')][contains(text(),'Envoyer')]")).Click();
                              
                               driver1.FindElement(By.XPath("//div[contains(@class,'ms-Button-label')][contains(text(),'Send')]")).Click();
                               System.Threading.Thread.Sleep(3500);*/

                            driver1.FindElement(By.XPath("//button[contains(@class, 'ms-Button') and contains(@aria-label,'Close')]")).Click();
                            /* IWebElement ele = driver1.FindElement(By.CssSelector("i[data-icon-name='Cancel']"));
                    Actions build = new Actions(driver1);
                      build.MoveToElement(ele,16,16).Click().Build().Perform();*/

                        }
                        if (df == "Sans doute pas")
                        {
                            driver1.FindElement(By.XPath("//button[contains(@class, 'ms-Button') and contains(@aria-label,'Fermer')]")).Click();

                        }
                    }
                    catch { }

                    try
                    {
                        var condition = driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']")).Displayed;
                        if (condition)
                        {
                            //  MessageBox.Show("true");  
                        }
                        else
                        {
                            //  MessageBox.Show("false");
                            driver1.FindElement(By.ClassName("_29AA3OWLpLPsgv7R6xoYtp")).Click();
                        }




                        /*    MessageBox.Show("he 1");
                            IJavaScriptExecutor js1 = (IJavaScriptExecutor)driver1;
                            js1.ExecuteScript("window.scrollTo(0, document.body.scrollHeight)");
                         string df = driver1.FindElement(By.CssSelector("i[data-icon-name='Cancel']")).Text.ToString();
                   
                       if (df == "")
                       {
                           MessageBox.Show("egale");
                           System.Threading.Thread.Sleep(2000);
                           driver1.FindElement(By.CssSelector("i[data-icon-name='Cancel']")).Click();
                           System.Threading.Thread.Sleep(2000);
                             IJavaScriptExecutor js1 = (IJavaScriptExecutor)driver1;
                              js.executeScript("window.scrollTo(0, document.body.scrollHeight)");
                             IWebElement Element14 = driver1.FindElement(By.CssSelector("i[data-icon-name='Cancel']"));
                             var tst = driver1.FindElement(By.CssSelector("i[data-icon-name='Cancel']")).Displayed;
                             System.Threading.Thread.Sleep(2000);
                             //   MessageBox.Show(tst.ToString());

                             js1.ExecuteScript("arguments[0].scrollIntoView();", Element14);
                             Element14.Click();
                             System.Threading.Thread.Sleep(1000);
                             // driver1.FindElement(By.ClassName("_3QWMcW9VDbQi11Tl_MkRml")).Click();*/



                    }

                    catch { }
                    //driver1.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(6);




                    System.Threading.Thread.Sleep(2000);
                    try
                    {

                        System.Threading.Thread.Sleep(1500);
                        var suiv1 = driver1.FindElement(By.XPath("//input[@id='i0118' and contains(text(),'')]")).Displayed;
                        // string fde = driver1.FindElement(By.ClassName("alert alert-error")).Text.ToString();
                        System.Threading.Thread.Sleep(1000);
                        //   MessageBox.Show(fde);
                        if (suiv1)
                        {
                            // MessageBox.Show("");
                         
                        }

                    }
                    catch { }
                    //Nous avons bien enregistré votre demande
                    Boolean visibledemande = false;

                    try
                    {
                        // try to find the element Nous avons bien enregistré votre demande

                        var boiteblocked4 = driver1.FindElement(By.ClassName("text-title")).Displayed;
                        if (boiteblocked4)
                        {
                            visibledemande = true;
                            string txtblock3 = driver1.FindElement(By.ClassName("text-title")).Text.ToString();
                            if (txtblock3 == "Nous avons bien enregistré votre demande")
                            {

                                //  MessageBox.Show("suivant");
                                System.Threading.Thread.Sleep(1000);
                                driver1.FindElement(By.XPath("//input[@type='button'][@value='Suivant']")).Click();
                                System.Threading.Thread.Sleep(1000);
                                driver1.Close();
                                driver1.Dispose();
                                // continue;
                            }
                        }

                    }
                    catch (Exception g)
                    {
                    }
                    //
                    try
                    {
                        IWebElement mv2 = driver1.FindElement(By.CssSelector("i[data-icon-name='Cancel']"));
                        Actions st1 = new Actions(driver1);
                        var tto = driver1.FindElement(By.CssSelector("i[data-icon-name='Cancel']")).Displayed;
                        if (tto)
                        {
                            // MessageBox.Show("ok ! cancel");
                            st1.MoveToElement(mv2, 16, 16).Click().Build().Perform();
                        }

                    }
                    catch { };
                    try
                    {
                        // open fisrt boite step
                        //ms-Icon ms-Icon--ChevronRight
                        var classfirstopen = driver1.FindElement(By.CssSelector("i[data-icon-name='ChevronRight']")).Displayed;
                        if (classfirstopen)
                        {
                            System.Threading.Thread.Sleep(3000);
                            //  MessageBox.Show("done fisrt");
                            //1
                            //iconButton nextButton lowerButton
                            IWebElement mv1 = driver1.FindElement(By.CssSelector("i[data-icon-name='ChevronRight']"));
                            Actions st10 = new Actions(driver1);
                            st10.MoveToElement(mv1, 40, 16).Click().Build().Perform();

                            System.Threading.Thread.Sleep(3000);
                            //2
                            //iconButton nextButton lowerButton
                            IWebElement mv2 = driver1.FindElement(By.CssSelector("i[data-icon-name='ChevronRight']"));
                            Actions st1 = new Actions(driver1);
                            st1.MoveToElement(mv2, 40, 16).Click().Build().Perform();
                            System.Threading.Thread.Sleep(3000);
                            //3
                            //ms-Icon ms-Icon--ChevronRight
                            IWebElement mv8 = driver1.FindElement(By.CssSelector("i[data-icon-name='ChevronRight']"));
                            Actions st2 = new Actions(driver1);
                            st2.MoveToElement(mv8, 40, 16).Click().Build().Perform();
                            System.Threading.Thread.Sleep(3000);
                            //4
                            //iconButton nextButton
                            IWebElement mv4 = driver1.FindElement(By.CssSelector("i[data-icon-name='ChevronRight']"));
                            Actions st3 = new Actions(driver1);
                            st3.MoveToElement(mv4, 40, 16).Click().Build().Perform();
                            System.Threading.Thread.Sleep(3000);


                            // btn ok 
                            //
                            //driver1.FindElement(By.ClassName("ms-Button-label")).Click();
                              driver1.FindElement(By.XPath("//div[contains(@class,'ms-Button-label')][contains(text(),'Got it')]")).Click();
                            System.Threading.Thread.Sleep(3000);


                            System.Threading.Thread.Sleep(2000);
                       
                            driver1.Close();
                            driver1.Dispose();
                            continue;

                        }

                    }
                    catch { }


                    try
                    {
                        // open fisrt boite step
                        //ms-Icon ms-Icon--ChevronRight
                        var classfirstopen = driver1.FindElement(By.CssSelector("i[class='ms-Icon ms-Icon--ChevronRight']")).Displayed;
                        if (classfirstopen)
                        {
                            System.Threading.Thread.Sleep(3000);
                            //  MessageBox.Show("done fisrt");
                            //1
                            //iconButton nextButton lowerButton
                            IWebElement mv1 = driver1.FindElement(By.CssSelector("i[class='ms-Icon ms-Icon--ChevronRight']"));
                            Actions st10 = new Actions(driver1);
                            st10.MoveToElement(mv1, 30, 30).Click().Build().Perform();

                            System.Threading.Thread.Sleep(3000);
                            //2
                            //iconButton nextButton lowerButton
                            IWebElement mv2 = driver1.FindElement(By.CssSelector("i[class='ms-Icon ms-Icon--ChevronRight']"));
                            Actions st1 = new Actions(driver1);
                            st1.MoveToElement(mv2, 30, 30).Click().Build().Perform();
                            System.Threading.Thread.Sleep(3000);
                            //3
                            //ms-Icon ms-Icon--ChevronRight
                            IWebElement mv8 = driver1.FindElement(By.CssSelector("i[class='ms-Icon ms-Icon--ChevronRight']"));
                            Actions st2 = new Actions(driver1);
                            st2.MoveToElement(mv8, 30, 30).Click().Build().Perform();
                            System.Threading.Thread.Sleep(3000);
                            //4
                            //iconButton nextButton
                            IWebElement mv4 = driver1.FindElement(By.CssSelector("i[class='ms-Icon ms-Icon--ChevronRight']"));
                            Actions st3 = new Actions(driver1);
                            st3.MoveToElement(mv4, 30, 30).Click().Build().Perform();
                            System.Threading.Thread.Sleep(3000);


                            // btn ok 
                            //
                            driver1.FindElement(By.ClassName("primaryButton")).Click();
                            System.Threading.Thread.Sleep(3000);


                            System.Threading.Thread.Sleep(2000);
                          
                            driver1.Close();
                            driver1.Dispose();
                            continue;

                        }

                    }
                    catch { }

                    //btn enregistrer
                    try
                    {

                        var enregistrer = driver1.FindElement(By.XPath("//span[contains(@class,'signinTxt')][contains(text(),'Enregistrer')]")).Displayed;

                        if (enregistrer)
                        {

                            System.Threading.Thread.Sleep(3000);
                            var select = driver1.FindElement(By.Id("selTz")).Displayed;
                            if (select)
                            {
                                // MessageBox.Show("here");

                            }

                            driver1.FindElement(By.Id("selTz")).FindElement(By.XPath(".//option[contains(text(),'‎(UTC-12:00)‎ Ligne de changement de date internationale ‎(Ouest)‎')]")).Click();
                            System.Threading.Thread.Sleep(3000);
                            driver1.FindElement(By.XPath("//span[contains(@class,'signinTxt')][contains(text(),'Enregistrer')]")).Click();
                            System.Threading.Thread.Sleep(3000);
                            //listBox3.Items.Add(ips + "," + boite + "," + pass + " First Use Time zone Updated");
                            driver1.Close();
                            driver1.Dispose();
                            continue;


                        }

                    }
                    catch { }


                    Boolean visible = false;

                    try
                    {
                        // try to find the element 
                        var boiteblocked = driver1.FindElement(By.ClassName("text-title")).Displayed;
                        if (boiteblocked)
                            visible = true;
                        string txtblock3 = driver1.FindElement(By.ClassName("text-title")).Text.ToString();
                        if (txtblock3 == "Aidez-nous à protéger votre compte")
                        {
                               System.Threading.Thread.Sleep(1000);
                            driver1.Close();
                            driver1.Dispose();
                            continue;
                        }
                        if (txtblock3 == "Votre compte a été temporairement bloqué")
                        {
                            driver1.FindElement(By.XPath("//input[@type='submit'][@value='Continuer']")).Click();
                            driver1.Close();
                            driver1.Dispose();
                            continue;
                        }
                        else continue;
                    }
                    catch (Exception g)
                    {
                    }
                    // try for older boite
                    Boolean visible1 = false;
                    try
                    {
                        // try to find the element 
                        var boiteblocked = driver1.FindElement(By.XPath("//span[contains(@class,'o365cs-nav-brandingText')][contains(text(),'Courrier Outlook')]")).Displayed;
                        if (boiteblocked)
                            visible1 = true;
                        string txtblock4 = driver1.FindElement(By.XPath("//span[contains(@class,'o365cs-nav-brandingText')][contains(text(),'Courrier Outlook')]")).Text.ToString();
                        //  MessageBox.Show(txtblock4);
                        if (txtblock4 == "Courrier Outlook")
                        {
                            versionofmail = "olderversion";

                        }

                    }
                    catch (Exception g)
                    {
                    }
                    // try for new boite
                    Boolean visible2 = false;
                    try
                    {
                        // try to find the element 
                        System.Threading.Thread.Sleep(2000);
                        var boiteblocked1 = driver1.FindElement(By.CssSelector("i[data-icon-name='Waffle']")).Displayed;
                        // if I get to here the element exists 
                        // if it is visible 
                        if (boiteblocked1)
                            visible2 = true;
                        string txtblock = driver1.FindElement(By.CssSelector("i[data-icon-name='Waffle']")).Text.ToString();
                        //  MessageBox.Show(txtblock);

                        if (txtblock == "")
                        {
                            versionofmail = "newversion";
                        }
                    }
                    catch (Exception g)
                    {
                    }

                    System.Threading.Thread.Sleep(1000);


                    Boolean visible6 = false;
                    try
                    {
                        /*
                         *   System.Threading.Thread.Sleep(3000);
                           IWebElement firststep = driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']"));
                           System.Threading.Thread.Sleep(2000);
                           var openlistspam = driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']")).Displayed;
                           System.Threading.Thread.Sleep(2000);
                           firststep.Click();
                           System.Threading.Thread.Sleep(2000);
                         * */
                        // try this code if work
                        try
                        {

                            var clickenglich = driver1.FindElement(By.XPath("//span[contains(@class,'_3lQ0EN5N3oGHxkBKF6_Ane')][contains(text(),'Junk Email')]")).Displayed;
                            if (clickenglich)
                            {
                                System.Threading.Thread.Sleep(3000);
                                IWebElement firststep = driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']"));
                                System.Threading.Thread.Sleep(2000);
                                var openlistspam = driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']")).Displayed;
                                System.Threading.Thread.Sleep(2000);
                                firststep.Click();
                                System.Threading.Thread.Sleep(2000);
                            }

                        }
                        catch { }
                        try
                        {

                            var clickfrench = driver1.FindElement(By.XPath("//span[contains(@class,'_3lQ0EN5N3oGHxkBKF6_Ane')][contains(text(),'Courrier indésirable')]")).Displayed;
                            if (clickfrench)
                            {
                                System.Threading.Thread.Sleep(3000);
                                IWebElement firststep = driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']"));
                                System.Threading.Thread.Sleep(2000);
                                var openlistspam = driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']")).Displayed;
                                System.Threading.Thread.Sleep(2000);
                                firststep.Click();
                                System.Threading.Thread.Sleep(2000);
                            }

                        }
                        catch { }


                        System.Threading.Thread.Sleep(3000);
                        IWebElement firststep2 = driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']"));
                        System.Threading.Thread.Sleep(2000);
                        var openlistspam2 = driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']")).Displayed;
                        System.Threading.Thread.Sleep(2000);
                        firststep2.Click();
                        System.Threading.Thread.Sleep(2000);

                        driver1.FindElement(By.CssSelector("i[data-icon-name='Blocked']")).Click();
                        //driver1.FindElement(By.XPath("//span[contains(@class,'_3lQ0EN5N3oGHxkBKF6_Ane')][contains(text(),'Courrier indésirable')]")).Click();
                        var spamshow = driver1.FindElement(By.CssSelector("i[data-icon-name='Blocked']")).Displayed;
                        //  var spamshow = driver1.FindElement(By.XPath("//span[contains(@class,'_3lQ0EN5N3oGHxkBKF6_Ane')][contains(text(),'Courrier indésirable')]")).Displayed;
                        if (spamshow)
                        {
                            System.Threading.Thread.Sleep(2000);
                            // last try
                            try
                            {
                                try
                                {

                                    var dossiervide = driver1.FindElement(By.XPath("//div[contains(@class,'_25o_l9GBJ9r8kRewkYoBfm _2oemjr1JNV5H9BUBUobbw2')][contains(text(),'Ce dossier est vide')]")).Text.ToString();
                                    if (dossiervide == "Ce dossier est vide")
                                    {
                                        // MessageBox.Show("vide");
                                         driver1.Close();
                                        driver1.Dispose();
                                        continue;
                                    }

                                }
                                catch { }


                                try
                                {
                                    var dossiervide1 = driver1.FindElement(By.XPath("//div[contains(@class,'_25o_l9GBJ9r8kRewkYoBfm _2oemjr1JNV5H9BUBUobbw2')][contains(text(),'This folder is empty')]")).Text.ToString();
                                    if (dossiervide1 == "This folder is empty")
                                    {
                                        // MessageBox.Show("vide");

                                      driver1.Close();
                                        driver1.Dispose();
                                        continue;
                                    }

                                }
                                catch { }
                                // try to find the element 
                                System.Threading.Thread.Sleep(2000);
                                var markasread = driver1.FindElement(By.CssSelector("i[data-icon-name='Read']")).Displayed;
                                // if I get to here the element exists 
                                // if it is visible 
                                System.Threading.Thread.Sleep(2000);

                                if (markasread)
                                {
                                    visible2 = true;
                                    string markasread1 = driver1.FindElement(By.CssSelector("i[data-icon-name='Read']")).Text.ToString();
                                    //  MessageBox.Show(txtblock);

                                    if (markasread1 == "")
                                    {
                                        //MessageBox.Show("comme lu");
                                        // try to read spam and move it
                                        try
                                        {
                                            System.Threading.Thread.Sleep(2000);

                                            var boiteblocked6 = driver1.FindElement(By.ClassName("_2EHjCdO2IEh-zlrH2jOD50")).Displayed;
                                            System.Threading.Thread.Sleep(2000);
                                            //   string fe = driver1.FindElement(By.XPath("//span[contains(@class,'ms-Button-label label-55')][contains(text(),'Marquer tout comme lu')]")).Text.ToString();
                                            //  MessageBox.Show(fe.ToString());
                                            // Ce dossier est vide
                                            do
                                            {
                                                if (boiteblocked6)
                                                {

                                                    // Boolean visible1 = false;  // assume it is invisible 
                                                    /* IWebElement firststep = driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']"));
                                                     firststep.Click();*/
                                                    //*********************code for new version 
                                                    System.Threading.Thread.Sleep(2000);
                                                    driver1.FindElement(By.CssSelector("i[data-icon-name='Blocked']")).Click();
                                                    //driver1.FindElement(By.XPath("//span[contains(@class,'_3lQ0EN5N3oGHxkBKF6_Ane')][contains(text(),'Courrier indésirable')]")).Click();
                                                    System.Threading.Thread.Sleep(2000);
                                                    driver1.FindElement(By.ClassName("_2EHjCdO2IEh-zlrH2jOD50")).Click();
                                                    System.Threading.Thread.Sleep(1500);
                                                    /* driver1.FindElement(By.CssSelector("i[data-icon-name='ChevronDown']")).Click();
                                                     IJavaScriptExecutor js1 = (IJavaScriptExecutor)driver1;

                                                     System.Threading.Thread.Sleep(1500);
                                                    // IWebElement Element1 = driver1.FindElement(By.XPath("//span[contains(@class,'ms-ContextualMenu-itemText label-170')][contains(text(),'Courrier légitime')]"));
                                                     IWebElement Elementbtn = driver1.FindElement(By.XPath("//button[contains(@class,'ms-ContextualMenu-link root-146')][contains(@name,'Courrier légitime')]"));
                                                     MessageBox.Show(Elementbtn.ToString());
                                                       
                                                     System.Threading.Thread.Sleep(15000);
                                                    // IWebElement Element = driver1.FindElement(By.XPath("//span[contains(@class,'ms-ContextualMenu-itemText label-170')][contains(text(),'Not spam')]"));
                                                     //This will scroll the page till the element is found		

                                                     js1.ExecuteScript("arguments[0].scrollIntoView();", Elementbtn);
                                                     Elementbtn.Click();*/
                                                    System.Threading.Thread.Sleep(1500);
                                                    Boolean visibleselctbox = false;
                                                    try
                                                    {
                                                        // var selctbox = driver1.FindElement(By.XPath("//span[contains(@class,'ms-Button-label label-55')][contains(text(),'Not spam')]")).Displayed;

                                                        var courierligitme = driver1.FindElementByLinkText("Courrier légitime").Displayed;
                                                        if (courierligitme)
                                                        {
                                                            visibleselctbox = true;

                                                            driver1.FindElementByLinkText("Courrier légitime").Click();

                                                        }
                                                        /* if(selctbox)
                                                          {
                                                              visibleselctbox = true;

                                                              driver1.FindElement(By.XPath("//span[contains(@class,'ms-Button-label label-55')][contains(text(),'Not spam')]")).Click();
                                                               
                                                            //  MessageBox.Show("here");
                                                    //  driver1.FindElement(By.CssSelector("i[data-icon-name='ChevronDown']")).Click();
                                                      IJavaScriptExecutor js1 = (IJavaScriptExecutor)driver1;

                                                      System.Threading.Thread.Sleep(1500);
                                                     // IWebElement Element1 = driver1.FindElement(By.XPath("//span[contains(@class,'ms-ContextualMenu-itemText label-170')][contains(text(),'Courrier légitime')]"));
                                                         
                                                      System.Threading.Thread.Sleep(15000);
                                                      IWebElement Element = driver1.FindElement(By.XPath("//span[contains(@class,'ms-ContextualMenu-linkContent linkContent-262')][contains(text(),'Not spam')]"));
                                                      //This will scroll the page till the element is found		

                                                      js1.ExecuteScript("arguments[0].scrollIntoView();", Element);
                                                      driver1.FindElement(By.XPath("//span[contains(@class,'ms-ContextualMenu-linkContent linkContent-262')][contains(text(),'Not spam')]")).Click();
                                                        
                                                          }*/

                                                        // MessageBox.Show("over else");
                                                    }
                                                    catch { }
                                                    // new try version anglais 
                                                    try
                                                    {
                                                        var selctbox = driver1.FindElement(By.CssSelector("i[data-icon-name='ChevronDown']")).Displayed;


                                                        if (selctbox)
                                                        {
                                                            visibleselctbox = true;

                                                            //   driver1.FindElement(By.XPath("//span[contains(@class,'ms-Button-label label-55')][contains(text(),'Not spam')]")).Click();


                                                            /*   driver1.FindElement(By.CssSelector("i[data-icon-name='ChevronDown']")).Click();
                                                               IJavaScriptExecutor js1 = (IJavaScriptExecutor)driver1;
                                                               // var rez = driver1.FindElement(By.CssSelector("i[data-icon-name='ChevronDown']")).Location;
                                                               // MessageBox.Show(rez.ToString());
                                                                System.Threading.Thread.Sleep(1500);
                                                               // IWebElement Element1 = driver1.FindElement(By.XPath("//span[contains(@class,'ms-ContextualMenu-itemText label-170')][contains(text(),'Courrier légitime')]"));
                                                         
                                                              //  System.Threading.Thread.Sleep(15000);

                                                                IWebElement Element = driver1.FindElement(By.XPath("//span[contains(@class,'ms-ContextualMenu-itemText label-215')][contains(text(),'Not spam')]"));
                                                                //This will scroll the page till the element is found		
                                                                      // ms-ContextualMenu-item item-160
                                                                string notspm = driver1.FindElement(By.XPath("//span[contains(@class,'ms-ContextualMenu-itemText label-215')][contains(text(),'Not spam')]")).Text.ToString();
                                                                    
                                                                       js1.ExecuteScript("arguments[0].scrollIntoView();",Element);*/
                                                            // MessageBox.Show(notspm);
                                                            // move spam to inbox
                                                            /*    Actions builder = new Actions(driver1);
                                                                IWebElement skype = driver1.FindElement(By.CssSelector("i[data-icon-name='ChevronDown']"));
                                                                builder.MoveToElement(skype, 12, 16).Click().Perform();
                                                                //builder.MoveByOffset(20,20).Click().Perform();
                                                                System.Threading.Thread.Sleep(5000);
                                                                Actions builder1 = new Actions(driver1);
                                                                IJavaScriptExecutor js1 = (IJavaScriptExecutor)driver1;
                                                                // driver1.FindElement(By.XPath("//button[contains(@name,'Not spam')][contains(text(),'Not spam')]"));
                                                                IWebElement skype1 = driver1.FindElement(By.CssSelector("[tabindex='0']"));
                                                                js1.ExecuteScript("arguments[0].scrollIntoView();", skype1);
                                                                    builder.MoveToElement(skype1,160,32).Click().Perform();*/
                                                            //sent with body
                                                            // driver1.FindElementByLinkText("Show blocked content").Click();
                                                            //sent with no body

                                                            driver1.FindElementByLinkText("It's not spam").Click();
                                                            // Element.Click();
                                                            // driver1.FindElement(By.XPath("//span[contains(@class,'ms-ContextualMenu-linkContent linkContent-262')][contains(text(),'Not spam')]")).Click();
                                                            //  driver1.FindElement(By.Name("Not spam")).Click();
                                                        }

                                                        //  MessageBox.Show("anglais");
                                                    }
                                                    catch { }

                                                    // end try version anglais

                                                    System.Threading.Thread.Sleep(2500);
                                                    visible6 = false;

                                                }
                                            } while (markasread);

                                        }

                                        catch
                                        {

                                        }
                                    
                                    }

                                    //  MessageBox.Show("here");


                                }

                                if (markasread == false)
                                {
   continue;
                                }
                            }
                            catch (Exception g)
                            {
                            }

                            // last try
                        }

                        // close this else 
                        /* else
                         {
                             driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']")).Click();
                         }*/

                    }
                    catch
                    {
                    }

                    //  driver1.FindElement(By.XPath("//span[contains(@class,'_3lQ0EN5N3oGHxkBKF6_Ane')][contains(text(),'Courrier indésirable')]")).Click();
                    System.Threading.Thread.Sleep(2000);
                    try
                    {
                      driver1.Close();
                        driver1.Dispose();
                    }
                    catch { }
                }

                else
                {
                    // MessageBox.Show("tt False");
                }

            }
            int reportinglistnumber = 0;
            /*  StringBuilder listViewContent = new StringBuilder();
           
           
              foreach (ListViewItem item in this.listView1.Items)
              {
                  listViewContent.Append(item);
                  listViewContent.Append(Environment.NewLine);
                
                  TextWriter tw = new StreamWriter("C:/Users/Public/ReportingResult" + reportinglistnumber+".txt");

                  tw.WriteLine(listViewContent.ToString());
             
                  tw.Close();
              }*/

            string sPath = "C:/Users/Public/ReportingResult.txt";
            string sPath1 = "C:/Users/Public/openitagain.txt";
            string sPath2 = "C:/Users/Public/BoiteBlocked.txt";

            // string sPath2 = "C:/Users/Public/BoiteBlocked" + reportinglistnumber +DateTime.Now.ToString("yyyyMMdd_hhmmss")+ ".txt";
            System.IO.StreamWriter SaveFile = new System.IO.StreamWriter(sPath);

            System.IO.StreamWriter SaveFile1 = new System.IO.StreamWriter(sPath1);
            System.IO.StreamWriter SaveFile2 = new System.IO.StreamWriter(sPath2);
          

            SaveFile1.Close();


            // done 
         
            MessageBox.Show("Reporting Done!");


        }

        public static bool PingHost(string strIP, int intPort)
        {
            bool blProxy = false;
            try
            {
                TcpClient client = new TcpClient(strIP, intPort);

                blProxy = true;
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Error pinging host:'" + strIP + ":" + intPort .ToString() + "'");
                return false;
            }
            return blProxy;
        }


        private void button25_Click(object sender, EventArgs e)
        {

        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
         //   int SerialsCounter = listBox2.Items.Count;
            //label40.Text = SerialsCounter.ToString();
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            richTextBox7.Clear();



            /* string source = @"First name: Elvis
Last name: Presley
Address: 1 Heaven Street
City: Memphis
State: TN
Zip: 12345
";
           MessageBox.Show(ExtractJist(source));*/
            // string sourceString = "sid [1544764] srv [CFT256] remip [10.0.128.31] fwf []...remip [10.0.128.41] fwf []... ";
            string content = richTextBox5.Text;
            string content2 = Regex.Replace(content, "_", " ");
            string content1 = richTextBox5.Text;
            /*string[] content3 = Regex.Split(content1, @"\W");
               MessageBox.Show(content3[0]);*/

            var matches = Regex.Matches(content2, @"(\d{0,9}\.\d{0,9}\.\d{0,9}\.\d{0,9})");
            Regex regex = new Regex(@"[A-Za-z]+Story[0-9]*");
            var mat = Regex.Matches(richTextBox5.Text, @"[A-Za-z]+[0-9]*");
            //var matches1 = Regex.Match(content3[0],@"^\d");
            foreach (var match in mat)
            {
                // MessageBox.Show(match.ToString());
                //  richTextBox7.Text += richTextBox5.Lines[match].Split('_').First() + "\n";
                /* richTextBox5.Text += match.ToString();
                 richTextBox5.AppendText(Environment.NewLine);*/
                richTextBox7.Text += match.ToString();
                richTextBox7.AppendText(Environment.NewLine);

            }
            /*  for (int n = 0; n < richTextBox5.Lines.Length; ++n)
              {
                  richTextBox7.Text += richTextBox5.Lines[n].Split('_').First() + "\n";
                  //richTextBox7.Text += richTextBox5.Lines[n].Split('_').Last() + "\n";
              }*/

            // var matches = Regex.Matches(content2, @"\b(\d{0,9}\.\d{0,9}\.\d{0,9}\.\d{0,9})\b");
            richTextBox5.Clear();
            foreach (var match in matches)
            {

                richTextBox5.Text += match.ToString();
                richTextBox5.AppendText(Environment.NewLine);

            }

            richTextBox7.Text = string.Join(Environment.NewLine, richTextBox7.Lines.Distinct());
            richTextBox5.Text = string.Join(Environment.NewLine, richTextBox5.Lines.Distinct());
            label22.Text = matches.Count.ToString();
            //label28.Text = matches1.Count.ToString();
            /* string str = "Title = Car Promotion, Model = BMW 323";
             Debug.Write(str.Split('=').Last());*/

            MessageBox.Show("Done!");

            /*  SautinSoft.HtmlToRtf h = new SautinSoft.HtmlToRtf();           
              // string htmlString = @"<b>Hello World!</b>";
              string htmlString = @"\b(\d{0,9}\.\d{0,9}\.\d{0,9}\.\d{0,9})\b";
                 h.OutputFormat = HtmlToRtf.eOutputFormat.TextUnicode;
               string textString = h.ConvertString(htmlString);
               MessageBox.Show(textString);*/
        }

        private void pictureBox7_Click(object sender, EventArgs e)
        {
            richTextBox5.Text = string.Join(Environment.NewLine, richTextBox5.Lines.Distinct());
            label22.Text = richTextBox5.Lines.Count().ToString();
            MessageBox.Show("Done!");
        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {
            richTextBox5.Text = string.Join(Environment.NewLine, richTextBox5.Lines.Distinct());
            label22.Text = richTextBox5.Lines.Count().ToString();
            MessageBox.Show("Done!");
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            richTextBox7.Clear();



            /* string source = @"First name: Elvis
Last name: Presley
Address: 1 Heaven Street
City: Memphis
State: TN
Zip: 12345
";
           MessageBox.Show(ExtractJist(source));*/
            // string sourceString = "sid [1544764] srv [CFT256] remip [10.0.128.31] fwf []...remip [10.0.128.41] fwf []... ";
            string content = richTextBox5.Text;
            string content2 = Regex.Replace(content, "_", " ");
            string content1 = richTextBox5.Text;
            /*string[] content3 = Regex.Split(content1, @"\W");
               MessageBox.Show(content3[0]);*/

            var matches = Regex.Matches(content2, @"(\d{0,9}\.\d{0,9}\.\d{0,9}\.\d{0,9})");
            Regex regex = new Regex(@"[A-Za-z]+Story[0-9]*");
            var mat = Regex.Matches(richTextBox5.Text, @"[A-Za-z]+[0-9]*");
            //var matches1 = Regex.Match(content3[0],@"^\d");
            foreach (var match in mat)
            {
                // MessageBox.Show(match.ToString());
                //  richTextBox7.Text += richTextBox5.Lines[match].Split('_').First() + "\n";
                /* richTextBox5.Text += match.ToString();
                 richTextBox5.AppendText(Environment.NewLine);*/
                richTextBox7.Text += match.ToString();
                richTextBox7.AppendText(Environment.NewLine);

            }
            /*  for (int n = 0; n < richTextBox5.Lines.Length; ++n)
              {
                  richTextBox7.Text += richTextBox5.Lines[n].Split('_').First() + "\n";
                  //richTextBox7.Text += richTextBox5.Lines[n].Split('_').Last() + "\n";
              }*/

            // var matches = Regex.Matches(content2, @"\b(\d{0,9}\.\d{0,9}\.\d{0,9}\.\d{0,9})\b");
            richTextBox5.Clear();
            foreach (var match in matches)
            {

                richTextBox5.Text += match.ToString();
                richTextBox5.AppendText(Environment.NewLine);

            }

            richTextBox7.Text = string.Join(Environment.NewLine, richTextBox7.Lines.Distinct());
            richTextBox5.Text = string.Join(Environment.NewLine, richTextBox5.Lines.Distinct());
            label22.Text = matches.Count.ToString();
            //label28.Text = matches1.Count.ToString();
            /* string str = "Title = Car Promotion, Model = BMW 323";
             Debug.Write(str.Split('=').Last());*/

            MessageBox.Show("Done!");

            /*  SautinSoft.HtmlToRtf h = new SautinSoft.HtmlToRtf();           
              // string htmlString = @"<b>Hello World!</b>";
              string htmlString = @"\b(\d{0,9}\.\d{0,9}\.\d{0,9}\.\d{0,9})\b";
                 h.OutputFormat = HtmlToRtf.eOutputFormat.TextUnicode;
               string textString = h.ConvertString(htmlString);
               MessageBox.Show(textString);*/
        }

        private void label37_Click(object sender, EventArgs e)
        {

        }

        private void button26_Click(object sender, EventArgs e)
        {
            ImageList imgs = new ImageList();
            imgs.ImageSize = new Size(50, 50);
            String[] pathss = { };
            pathss = Directory.GetFiles("C:/imgs");


            // buttun test sans imap
            string line;
            int number = 0;
            System.IO.StreamReader file =
    new System.IO.StreamReader(path);

            while ((line = file.ReadLine()) != null)
            {
                string[] lines = line.Split(',');

                string ips = lines[0];
                // MessageBox.Show(lines[0]);
                string boite = lines[1];
                string pass = lines[2];

                number++;
                string[] ipall = ips.Split(':');
                int port = Convert.ToInt32(ipall[1]);
                bool tt = PingHost(ipall[0], port);
                if (tt == true)
                {

                    // MessageBox.Show("tt True");


                    //**************************************

                    var chromeOptions = new ChromeOptions();
                    //Create a new proxy object
                    var proxy = new Proxy();
                    //Set the http proxy value, host and port.
                    // proxy.HttpProxy = ips;
                    //Set the proxy to the Chrome options
                    //  chromeOptions.Proxy = proxy;
                    //  chromeOptions.AddArguments("--proxy-server=" + ips);
                    chromeOptions.AddArgument("--lang=aus");
                    // disable pop ups 
                    /*DesiredCapabilities capabilities = DesiredCapabilities.Chrome();
                    capabilities.SetCapability(ChromeOptions.Capability, chromeOptions);*/
                    //chromeOptions.AddAdditionalCapability("excludeSwitches", "disable-popup-blocking");
                    //chromeOptions.AddArgument("--disable-popup-blocking");

                    var driver1 = new ChromeDriver(@"C:\webdrivers", chromeOptions);
                    //***********work without proxy******************
                    //var driver1 = new ChromeDriver(@"C:\webdrivers");
                    driver1.Navigate().GoToUrl(BaseUrl);
                    driver1.FindElementByLinkText("Se connecter").Click();

                    var loginBox = driver1.FindElement(By.XPath("//input[contains(text(),'')]"));
                    System.Threading.Thread.Sleep(2500);
                    loginBox.SendKeys(boite);
                    driver1.FindElement(By.XPath("//input[@type='submit'][@value='Suivant']")).Click();
                    System.Threading.Thread.Sleep(2500);
                    try
                    {
                        var suiv = driver1.FindElement(By.XPath("//input[@type='submit'][@value='Suivant']")).Displayed;
                        if (suiv)
                        {
                            /*  listView1.SmallImageList = imgs;
                              listView1.Items.Add("Michael Carrick", 1);*/
                          // listView1.Items.Add(boite + "User Inccorect ! Please Try Again");
                            driver1.Close();
                            driver1.Dispose();
                            continue;
                        }

                    }
                    catch { }

                    var passwordBox = driver1.FindElement(By.XPath("//input[@id='i0118' and contains(text(),'')]"));
                    passwordBox.SendKeys(pass);
                    System.Threading.Thread.Sleep(2500);
                    driver1.FindElement(By.XPath("//input[@type='submit'][@value='Se connecter']")).Click();
                    System.Threading.Thread.Sleep(4000);
                    driver1.Manage().Window.Maximize();
                    System.Threading.Thread.Sleep(4000);
                    // driver1.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);
                    try
                    {

                        //
                        //   MessageBox.Show("he 1");
                        System.Threading.Thread.Sleep(2000);
                        IJavaScriptExecutor js1 = (IJavaScriptExecutor)driver1;

                        IWebElement Element14 = driver1.FindElement(By.CssSelector("i[data-icon-name='Cancel']"));
                        var tst = driver1.FindElement(By.CssSelector("i[data-icon-name='Cancel']")).Displayed;
                        System.Threading.Thread.Sleep(2000);
                        //   MessageBox.Show(tst.ToString());
                        js1.ExecuteScript("arguments[0].scrollIntoView();", Element14);
                        Element14.Click();
                        System.Threading.Thread.Sleep(1000);
                        // driver1.FindElement(By.ClassName("_3QWMcW9VDbQi11Tl_MkRml")).Click();



                    }

                    catch { }



                    System.Threading.Thread.Sleep(2000);
                    try
                    {

                        System.Threading.Thread.Sleep(1500);
                        var suiv1 = driver1.FindElement(By.XPath("//input[@id='i0118' and contains(text(),'')]")).Displayed;
                        // string fde = driver1.FindElement(By.ClassName("alert alert-error")).Text.ToString();
                        System.Threading.Thread.Sleep(1000);
                        //   MessageBox.Show(fde);
                        if (suiv1)
                        {
                            // MessageBox.Show("");
                            driver1.Close();
                            driver1.Dispose();
                            continue;
                        }

                    }
                    catch { }
                    //Nous avons bien enregistré votre demande
                    Boolean visibledemande = false;

                    try
                    {
                        // try to find the element Nous avons bien enregistré votre demande

                        var boiteblocked4 = driver1.FindElement(By.ClassName("text-title")).Displayed;
                        if (boiteblocked4)
                        {
                            visibledemande = true;
                            string txtblock3 = driver1.FindElement(By.ClassName("text-title")).Text.ToString();
                            if (txtblock3 == "Nous avons bien enregistré votre demande")
                            {

                                //  MessageBox.Show("suivant");
                                System.Threading.Thread.Sleep(1000);

                                driver1.FindElement(By.XPath("//input[@type='button'][@value='Suivant']")).Click();

                                // continue;
                            }
                        }

                    }
                    catch (Exception g)
                    {
                    }


                    try
                    {
                        // open fisrt boite step
                        //ms-Icon ms-Icon--ChevronRight
                        var classfirstopen = driver1.FindElement(By.CssSelector("i[class='ms-Icon ms-Icon--ChevronRight']")).Displayed;
                        if (classfirstopen)
                        {
                            System.Threading.Thread.Sleep(3000);
                            //  MessageBox.Show("done fisrt");
                            //1
                            //iconButton nextButton lowerButton
                            IWebElement mv1 = driver1.FindElement(By.CssSelector("i[class='ms-Icon ms-Icon--ChevronRight']"));
                            Actions st = new Actions(driver1);
                            st.MoveToElement(mv1, 30, 30).Click().Build().Perform();

                            System.Threading.Thread.Sleep(3000);
                            //2
                            //iconButton nextButton lowerButton
                            IWebElement mv2 = driver1.FindElement(By.CssSelector("i[class='ms-Icon ms-Icon--ChevronRight']"));
                            Actions st1 = new Actions(driver1);
                            st1.MoveToElement(mv2, 30, 30).Click().Build().Perform();
                            System.Threading.Thread.Sleep(3000);
                            //3
                            //ms-Icon ms-Icon--ChevronRight
                            IWebElement mv3 = driver1.FindElement(By.CssSelector("i[class='ms-Icon ms-Icon--ChevronRight']"));
                            Actions st2 = new Actions(driver1);
                            st2.MoveToElement(mv3, 30, 30).Click().Build().Perform();
                            System.Threading.Thread.Sleep(3000);
                            //4
                            //iconButton nextButton
                            IWebElement mv4 = driver1.FindElement(By.CssSelector("i[class='ms-Icon ms-Icon--ChevronRight']"));
                            Actions st3 = new Actions(driver1);
                            st3.MoveToElement(mv4, 30, 30).Click().Build().Perform();
                            System.Threading.Thread.Sleep(3000);


                            // btn ok 
                            //
                            driver1.FindElement(By.ClassName("primaryButton")).Click();
                            System.Threading.Thread.Sleep(3000);


                            System.Threading.Thread.Sleep(2000);
                            //listBox3.Items.Add(ips + "," + boite + "," + pass + " First Use");

                            driver1.Close();
                            driver1.Dispose();
                            continue;

                        }

                    }
                    catch { }

                    //btn enregistrer
                    try
                    {

                        var enregistrer = driver1.FindElement(By.XPath("//span[contains(@class,'signinTxt')][contains(text(),'Enregistrer')]")).Displayed;

                        if (enregistrer)
                        {

                            System.Threading.Thread.Sleep(3000);
                            var select = driver1.FindElement(By.Id("selTz")).Displayed;
                            if (select)
                            {
                                // MessageBox.Show("here");

                            }

                            driver1.FindElement(By.Id("selTz")).FindElement(By.XPath(".//option[contains(text(),'‎(UTC-12:00)‎ Ligne de changement de date internationale ‎(Ouest)‎')]")).Click();
                            System.Threading.Thread.Sleep(3000);
                            driver1.FindElement(By.XPath("//span[contains(@class,'signinTxt')][contains(text(),'Enregistrer')]")).Click();
                            System.Threading.Thread.Sleep(3000);
                            driver1.Close();
                            driver1.Dispose();
                            continue;


                        }

                    }
                    catch { }


                    Boolean visible = false;

                    try
                    {
                        // try to find the element 
                        var boiteblocked = driver1.FindElement(By.ClassName("text-title")).Displayed;
                        if (boiteblocked)
                            visible = true;
                        string txtblock3 = driver1.FindElement(By.ClassName("text-title")).Text.ToString();
                        if (txtblock3 == "Aidez-nous à protéger votre compte")
                        {
                            System.Threading.Thread.Sleep(1000);
                            driver1.Close();
                            driver1.Dispose();
                            continue;
                        }
                        if (txtblock3 == "Votre compte a été temporairement bloqué")
                        {

                           driver1.FindElement(By.XPath("//input[@type='submit'][@value='Continuer']")).Click();
                            driver1.Close();
                            driver1.Dispose();
                            continue;
                        }
                        else continue;
                    }
                    catch (Exception g)
                    {
                    }
                    // try for older boite
                    Boolean visible1 = false;
                    try
                    {
                        // try to find the element 
                        var boiteblocked = driver1.FindElement(By.XPath("//span[contains(@class,'o365cs-nav-brandingText')][contains(text(),'Courrier Outlook')]")).Displayed;
                        if (boiteblocked)
                            visible1 = true;
                        string txtblock4 = driver1.FindElement(By.XPath("//span[contains(@class,'o365cs-nav-brandingText')][contains(text(),'Courrier Outlook')]")).Text.ToString();
                        MessageBox.Show(txtblock4);
                        if (txtblock4 == "Courrier Outlook")
                        {
                            versionofmail = "olderversion";

                        }

                    }
                    catch (Exception g)
                    {
                    }
                    // try for new boite
                    Boolean visible2 = false;
                    try
                    {
                        // try to find the element 
                        System.Threading.Thread.Sleep(2000);
                        var boiteblocked1 = driver1.FindElement(By.CssSelector("i[data-icon-name='Waffle']")).Displayed;
                        // if I get to here the element exists 
                        // if it is visible 
                        if (boiteblocked1)
                            visible2 = true;
                        string txtblock = driver1.FindElement(By.CssSelector("i[data-icon-name='Waffle']")).Text.ToString();
                        //  MessageBox.Show(txtblock);

                        if (txtblock == "")
                        {
                            versionofmail = "newversion";
                        }
                    }
                    catch (Exception g)
                    {
                    }

                    System.Threading.Thread.Sleep(1000);


                    Boolean visible6 = false;
                    try
                    {
                        /*
                         *   System.Threading.Thread.Sleep(3000);
                           IWebElement firststep = driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']"));
                           System.Threading.Thread.Sleep(2000);
                           var openlistspam = driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']")).Displayed;
                           System.Threading.Thread.Sleep(2000);
                           firststep.Click();
                           System.Threading.Thread.Sleep(2000);
                         * */
                        // try this code if work
                        try
                        {

                            var clickenglich = driver1.FindElement(By.XPath("//span[contains(@class,'_3lQ0EN5N3oGHxkBKF6_Ane')][contains(text(),'Junk Email')]")).Displayed;
                            if (clickenglich)
                            {
                                System.Threading.Thread.Sleep(3000);
                                IWebElement firststep = driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']"));
                                System.Threading.Thread.Sleep(2000);
                                var openlistspam = driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']")).Displayed;
                                System.Threading.Thread.Sleep(2000);
                                firststep.Click();
                                System.Threading.Thread.Sleep(2000);
                            }

                        }
                        catch { }
                        try
                        {

                            var clickfrench = driver1.FindElement(By.XPath("//span[contains(@class,'_3lQ0EN5N3oGHxkBKF6_Ane')][contains(text(),'Courrier indésirable')]")).Displayed;
                            if (clickfrench)
                            {
                                System.Threading.Thread.Sleep(3000);
                                IWebElement firststep = driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']"));
                                System.Threading.Thread.Sleep(2000);
                                var openlistspam = driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']")).Displayed;
                                System.Threading.Thread.Sleep(2000);
                                firststep.Click();
                                System.Threading.Thread.Sleep(2000);
                            }

                        }
                        catch { }


                        System.Threading.Thread.Sleep(3000);
                        IWebElement firststep2 = driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']"));
                        System.Threading.Thread.Sleep(2000);
                        var openlistspam2 = driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']")).Displayed;
                        System.Threading.Thread.Sleep(2000);
                        firststep2.Click();
                        System.Threading.Thread.Sleep(2000);

                        driver1.FindElement(By.CssSelector("i[data-icon-name='Blocked']")).Click();
                        //driver1.FindElement(By.XPath("//span[contains(@class,'_3lQ0EN5N3oGHxkBKF6_Ane')][contains(text(),'Courrier indésirable')]")).Click();
                        var spamshow = driver1.FindElement(By.CssSelector("i[data-icon-name='Blocked']")).Displayed;
                        //  var spamshow = driver1.FindElement(By.XPath("//span[contains(@class,'_3lQ0EN5N3oGHxkBKF6_Ane')][contains(text(),'Courrier indésirable')]")).Displayed;
                        if (spamshow)
                        {
                            System.Threading.Thread.Sleep(2000);
                            // last try
                            try
                            {
                                try
                                {

                                    var dossiervide = driver1.FindElement(By.XPath("//div[contains(@class,'_25o_l9GBJ9r8kRewkYoBfm _2oemjr1JNV5H9BUBUobbw2')][contains(text(),'Ce dossier est vide')]")).Text.ToString();
                                    if (dossiervide == "Ce dossier est vide")
                                    {
                                        // MessageBox.Show("vide");
                                      driver1.Close();
                                        driver1.Dispose();
                                        continue;
                                    }

                                }
                                catch { }


                                try
                                {
                                    var dossiervide1 = driver1.FindElement(By.XPath("//div[contains(@class,'_25o_l9GBJ9r8kRewkYoBfm _2oemjr1JNV5H9BUBUobbw2')][contains(text(),'This folder is empty')]")).Text.ToString();
                                    if (dossiervide1 == "This folder is empty")
                                    {
                                        // MessageBox.Show("vide");

                                        driver1.Close();
                                        driver1.Dispose();
                                        continue;
                                    }

                                }
                                catch { }
                                // try to find the element 
                                System.Threading.Thread.Sleep(2000);
                                var markasread = driver1.FindElement(By.CssSelector("i[data-icon-name='Read']")).Displayed;
                                // if I get to here the element exists 
                                // if it is visible 
                                System.Threading.Thread.Sleep(2000);

                                if (markasread)
                                {
                                    visible2 = true;
                                    string markasread1 = driver1.FindElement(By.CssSelector("i[data-icon-name='Read']")).Text.ToString();
                                    //  MessageBox.Show(txtblock);

                                    if (markasread1 == "")
                                    {
                                        //MessageBox.Show("comme lu");
                                        // try to read spam and move it
                                        try
                                        {
                                            System.Threading.Thread.Sleep(2000);

                                            var boiteblocked6 = driver1.FindElement(By.ClassName("_2EHjCdO2IEh-zlrH2jOD50")).Displayed;
                                            System.Threading.Thread.Sleep(2000);
                                            //   string fe = driver1.FindElement(By.XPath("//span[contains(@class,'ms-Button-label label-55')][contains(text(),'Marquer tout comme lu')]")).Text.ToString();
                                            //  MessageBox.Show(fe.ToString());
                                            // Ce dossier est vide
                                            do
                                            {
                                                if (boiteblocked6)
                                                {

                                                    // Boolean visible1 = false;  // assume it is invisible 
                                                    /* IWebElement firststep = driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']"));
                                                     firststep.Click();*/
                                                    //*********************code for new version 
                                                    System.Threading.Thread.Sleep(2000);
                                                    driver1.FindElement(By.CssSelector("i[data-icon-name='Blocked']")).Click();
                                                    //driver1.FindElement(By.XPath("//span[contains(@class,'_3lQ0EN5N3oGHxkBKF6_Ane')][contains(text(),'Courrier indésirable')]")).Click();
                                                    System.Threading.Thread.Sleep(2000);
                                                    driver1.FindElement(By.ClassName("_2EHjCdO2IEh-zlrH2jOD50")).Click();
                                                    System.Threading.Thread.Sleep(1500);
                                                    /* driver1.FindElement(By.CssSelector("i[data-icon-name='ChevronDown']")).Click();
                                                     IJavaScriptExecutor js1 = (IJavaScriptExecutor)driver1;

                                                     System.Threading.Thread.Sleep(1500);
                                                    // IWebElement Element1 = driver1.FindElement(By.XPath("//span[contains(@class,'ms-ContextualMenu-itemText label-170')][contains(text(),'Courrier légitime')]"));
                                                     IWebElement Elementbtn = driver1.FindElement(By.XPath("//button[contains(@class,'ms-ContextualMenu-link root-146')][contains(@name,'Courrier légitime')]"));
                                                     MessageBox.Show(Elementbtn.ToString());
                                                       
                                                     System.Threading.Thread.Sleep(15000);
                                                    // IWebElement Element = driver1.FindElement(By.XPath("//span[contains(@class,'ms-ContextualMenu-itemText label-170')][contains(text(),'Not spam')]"));
                                                     //This will scroll the page till the element is found		

                                                     js1.ExecuteScript("arguments[0].scrollIntoView();", Elementbtn);
                                                     Elementbtn.Click();*/
                                                    System.Threading.Thread.Sleep(1500);
                                                    Boolean visibleselctbox = false;
                                                    try
                                                    {
                                                        // var selctbox = driver1.FindElement(By.XPath("//span[contains(@class,'ms-Button-label label-55')][contains(text(),'Not spam')]")).Displayed;

                                                        var courierligitme = driver1.FindElementByLinkText("Courrier légitime").Displayed;
                                                        if (courierligitme)
                                                        {
                                                            visibleselctbox = true;

                                                            driver1.FindElementByLinkText("Courrier légitime").Click();

                                                        }
                                                        /* if(selctbox)
                                                          {
                                                              visibleselctbox = true;

                                                              driver1.FindElement(By.XPath("//span[contains(@class,'ms-Button-label label-55')][contains(text(),'Not spam')]")).Click();
                                                               
                                                            //  MessageBox.Show("here");
                                                    //  driver1.FindElement(By.CssSelector("i[data-icon-name='ChevronDown']")).Click();
                                                      IJavaScriptExecutor js1 = (IJavaScriptExecutor)driver1;

                                                      System.Threading.Thread.Sleep(1500);
                                                     // IWebElement Element1 = driver1.FindElement(By.XPath("//span[contains(@class,'ms-ContextualMenu-itemText label-170')][contains(text(),'Courrier légitime')]"));
                                                         
                                                      System.Threading.Thread.Sleep(15000);
                                                      IWebElement Element = driver1.FindElement(By.XPath("//span[contains(@class,'ms-ContextualMenu-linkContent linkContent-262')][contains(text(),'Not spam')]"));
                                                      //This will scroll the page till the element is found		

                                                      js1.ExecuteScript("arguments[0].scrollIntoView();", Element);
                                                      driver1.FindElement(By.XPath("//span[contains(@class,'ms-ContextualMenu-linkContent linkContent-262')][contains(text(),'Not spam')]")).Click();
                                                        
                                                          }*/

                                                        // MessageBox.Show("over else");
                                                    }
                                                    catch { }
                                                    // new try version anglais 
                                                    try
                                                    {
                                                        var selctbox = driver1.FindElement(By.CssSelector("i[data-icon-name='ChevronDown']")).Displayed;


                                                        if (selctbox)
                                                        {
                                                            visibleselctbox = true;

                                                            //   driver1.FindElement(By.XPath("//span[contains(@class,'ms-Button-label label-55')][contains(text(),'Not spam')]")).Click();


                                                            /*   driver1.FindElement(By.CssSelector("i[data-icon-name='ChevronDown']")).Click();
                                                               IJavaScriptExecutor js1 = (IJavaScriptExecutor)driver1;
                                                               // var rez = driver1.FindElement(By.CssSelector("i[data-icon-name='ChevronDown']")).Location;
                                                               // MessageBox.Show(rez.ToString());
                                                                System.Threading.Thread.Sleep(1500);
                                                               // IWebElement Element1 = driver1.FindElement(By.XPath("//span[contains(@class,'ms-ContextualMenu-itemText label-170')][contains(text(),'Courrier légitime')]"));
                                                         
                                                              //  System.Threading.Thread.Sleep(15000);

                                                                IWebElement Element = driver1.FindElement(By.XPath("//span[contains(@class,'ms-ContextualMenu-itemText label-215')][contains(text(),'Not spam')]"));
                                                                //This will scroll the page till the element is found		
                                                                      // ms-ContextualMenu-item item-160
                                                                string notspm = driver1.FindElement(By.XPath("//span[contains(@class,'ms-ContextualMenu-itemText label-215')][contains(text(),'Not spam')]")).Text.ToString();
                                                                    
                                                                       js1.ExecuteScript("arguments[0].scrollIntoView();",Element);*/
                                                            // MessageBox.Show(notspm);
                                                            // move spam to inbox
                                                            /*    Actions builder = new Actions(driver1);
                                                                IWebElement skype = driver1.FindElement(By.CssSelector("i[data-icon-name='ChevronDown']"));
                                                                builder.MoveToElement(skype, 12, 16).Click().Perform();
                                                                //builder.MoveByOffset(20,20).Click().Perform();
                                                                System.Threading.Thread.Sleep(5000);
                                                                Actions builder1 = new Actions(driver1);
                                                                IJavaScriptExecutor js1 = (IJavaScriptExecutor)driver1;
                                                                // driver1.FindElement(By.XPath("//button[contains(@name,'Not spam')][contains(text(),'Not spam')]"));
                                                                IWebElement skype1 = driver1.FindElement(By.CssSelector("[tabindex='0']"));
                                                                js1.ExecuteScript("arguments[0].scrollIntoView();", skype1);
                                                                    builder.MoveToElement(skype1,160,32).Click().Perform();*/
                                                            //sent with body
                                                            // driver1.FindElementByLinkText("Show blocked content").Click();
                                                            //sent with no body

                                                            driver1.FindElementByLinkText("This isn't spam").Click();
                                                            // Element.Click();
                                                            // driver1.FindElement(By.XPath("//span[contains(@class,'ms-ContextualMenu-linkContent linkContent-262')][contains(text(),'Not spam')]")).Click();
                                                            //  driver1.FindElement(By.Name("Not spam")).Click();
                                                        }

                                                        //  MessageBox.Show("anglais");
                                                    }
                                                    catch { }

                                                    // end try version anglais

                                                    System.Threading.Thread.Sleep(2500);
                                                    visible6 = false;

                                                }
                                            } while (markasread);

                                        }

                                        catch
                                        {

                                        }
                                        driver1.Close();
                                        driver1.Dispose();
                                        continue;
                                    }

                                    //  MessageBox.Show("here");


                                }

                                if (markasread == false)
                                {

                                    continue;
                                }
                            }
                            catch (Exception g)
                            {
                            }

                            // last try
                        }

                        // close this else 
                        /* else
                         {
                             driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']")).Click();
                         }*/

                    }
                    catch
                    {
                    }

                    //  driver1.FindElement(By.XPath("//span[contains(@class,'_3lQ0EN5N3oGHxkBKF6_Ane')][contains(text(),'Courrier indésirable')]")).Click();
                    System.Threading.Thread.Sleep(2000);

                }

                else
                {
                   
                }
            }
            int reportinglistnumber = 0;
            /*  StringBuilder listViewContent = new StringBuilder();
           
           
              foreach (ListViewItem item in this.listView1.Items)
              {
                  listViewContent.Append(item);
                  listViewContent.Append(Environment.NewLine);
                
                  TextWriter tw = new StreamWriter("C:/Users/Public/ReportingResult" + reportinglistnumber+".txt");

                  tw.WriteLine(listViewContent.ToString());
             
                  tw.Close();
              }*/
            string sPath = "C:/Users/Public/ReportingResult" + reportinglistnumber + ".txt";

            System.IO.StreamWriter SaveFile = new System.IO.StreamWriter(sPath);


         /*   foreach (var item in listBox3.Items)
            {
                SaveFile.WriteLine(item.ToString());
            }
            SaveFile.Close();
            // done 
            progressBar2.Value = 100;
            MessageBox.Show("Reporting Done!");*/

        }

        private void button25_Click_1(object sender, EventArgs e)
        {
            ImageList imgs = new ImageList();
            imgs.ImageSize = new Size(50, 50);
            String[] pathss = { };
            pathss = Directory.GetFiles("C:/imgs");


            // buttun test sans imap
            string line;
            int number = 0;
            System.IO.StreamReader file =
    new System.IO.StreamReader(path);

            while ((line = file.ReadLine()) != null)
            {
                string[] lines = line.Split(',');

                string ips = lines[0];
                // MessageBox.Show(lines[0]);
                string boite = lines[1];
                string pass = lines[2];

                number++;
                string[] ipall = ips.Split(':');
                int port = Convert.ToInt32(ipall[1]);
                bool tt = PingHost(ipall[0], port);
                if (tt == true)
                {

                    // MessageBox.Show("tt True");


                    //**************************************

                    var chromeOptions = new ChromeOptions();
                    //Create a new proxy object
                    var proxy = new Proxy();
                    //Set the http proxy value, host and port.
                    // proxy.HttpProxy = ips;
                    //Set the proxy to the Chrome options
                    //  chromeOptions.Proxy = proxy;
                    //    chromeOptions.AddArgument("--headless");

                 
                    // disable pop ups 
                    /*DesiredCapabilities capabilities = DesiredCapabilities.Chrome();
                    capabilities.SetCapability(ChromeOptions.Capability, chromeOptions);*/
                    //chromeOptions.AddAdditionalCapability("excludeSwitches", "disable-popup-blocking");
                    //chromeOptions.AddArgument("--disable-popup-blocking");

                    var driver1 = new ChromeDriver(@"C:\webdrivers", chromeOptions);
                    //***********work without proxy******************
                    //var driver1 = new ChromeDriver(@"C:\webdrivers");
                    driver1.Navigate().GoToUrl(BaseUrl);
                    driver1.FindElementByLinkText("Se connecter").Click();

                    var loginBox = driver1.FindElement(By.XPath("//input[contains(text(),'')]"));
                    System.Threading.Thread.Sleep(2500);
                    loginBox.SendKeys(boite);
                    driver1.FindElement(By.XPath("//input[@type='submit'][@value='Suivant']")).Click();
                    System.Threading.Thread.Sleep(2500);
                    try
                    {
                        var suiv = driver1.FindElement(By.XPath("//input[@type='submit'][@value='Suivant']")).Displayed;
                        if (suiv)
                        {
                            /*  listView1.SmallImageList = imgs;
                              listView1.Items.Add("Michael Carrick", 1);*/
                          // listView1.Items.Add(boite + "User Inccorect ! Please Try Again");
                            driver1.Close();
                            driver1.Dispose();
                            continue;
                        }

                    }
                    catch { }

                    var passwordBox = driver1.FindElement(By.XPath("//input[@id='i0118' and contains(text(),'')]"));
                    passwordBox.SendKeys(pass);
                    System.Threading.Thread.Sleep(2500);
                    driver1.FindElement(By.XPath("//input[@type='submit'][@value='Se connecter']")).Click();
                    System.Threading.Thread.Sleep(4000);
                    try
                    {

                        driver1.FindElementByClassName("_29AA3OWLpLPsgv7R6xoYtp").Click();
                    }

                    catch { }

                    System.Threading.Thread.Sleep(12000);
                    driver1.Manage().Window.Maximize();
                    System.Threading.Thread.Sleep(12000);
                    try
                    {
                        string df = driver1.FindElement(By.ClassName("_36rwQx61abZ9wJjdE6uJPX")).Text.ToString();

                        //  MessageBox.Show(df);Sans doute pas
                        if (df == "Not likely")
                        {
                            /* driver1.FindElement(By.ClassName("_3GrIR9uD5dhhMZHdiJ2GC9")).Click();
                             //driver1.FindElement(By.ClassName("ms-Button-icon")).Click();
                             System.Threading.Thread.Sleep(2500);
                             var msgtosent = driver1.FindElement(By.XPath("//textarea[contains(text(),'')]"));
                             System.Threading.Thread.Sleep(2500);
                             msgtosent.SendKeys("good");
                             System.Threading.Thread.Sleep(2500);
                             // driver1.FindElement(By.XPath("//div[contains(@class,'ms-Button-label')][contains(text(),'Envoyer')]")).Click();

                             driver1.FindElement(By.XPath("//div[contains(@class,'ms-Button-label')][contains(text(),'Send')]")).Click();
                             System.Threading.Thread.Sleep(3500);*/

                            driver1.FindElement(By.XPath("//button[contains(@class, 'ms-Button') and contains(@aria-label,'Close')]")).Click();
                            /* IWebElement ele = driver1.FindElement(By.CssSelector("i[data-icon-name='Cancel']"));
                    Actions build = new Actions(driver1);
                      build.MoveToElement(ele,16,16).Click().Build().Perform();*/

                        }
                        if (df == "Sans doute pas")
                        {
                            driver1.FindElement(By.XPath("//button[contains(@class, 'ms-Button') and contains(@aria-label,'Fermer')]")).Click();

                        }
                    }
                    catch { }
                    try
                    {
                        System.Threading.Thread.Sleep(4000);
                        var condition = driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']")).Displayed;
                        if (condition)
                        {
                            //  MessageBox.Show("true");  
                        }
                        else
                        {
                            //  MessageBox.Show("false");
                            driver1.FindElement(By.ClassName("_29AA3OWLpLPsgv7R6xoYtp")).Click();
                        }
                    }
                    catch { }


                    // driver1.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);
                    try
                    {

                        //
                        //   MessageBox.Show("he 1");
                        System.Threading.Thread.Sleep(2000);
                        IJavaScriptExecutor js1 = (IJavaScriptExecutor)driver1;

                        IWebElement Element14 = driver1.FindElement(By.CssSelector("i[data-icon-name='Cancel']"));
                        var tst = driver1.FindElement(By.CssSelector("i[data-icon-name='Cancel']")).Displayed;
                        System.Threading.Thread.Sleep(2000);
                        //   MessageBox.Show(tst.ToString());
                        js1.ExecuteScript("arguments[0].scrollIntoView();", Element14);
                        Element14.Click();
                        System.Threading.Thread.Sleep(1000);
                        // driver1.FindElement(By.ClassName("_3QWMcW9VDbQi11Tl_MkRml")).Click();



                    }

                    catch { }



                    System.Threading.Thread.Sleep(2000);
                    try
                    {

                        System.Threading.Thread.Sleep(1500);
                        var suiv1 = driver1.FindElement(By.XPath("//input[@id='i0118' and contains(text(),'')]")).Displayed;
                        // string fde = driver1.FindElement(By.ClassName("alert alert-error")).Text.ToString();
                        System.Threading.Thread.Sleep(1000);
                        //   MessageBox.Show(fde);
                        if (suiv1)
                        {
                            // MessageBox.Show("");
                          
                        }

                    }
                    catch { }
                    //Nous avons bien enregistré votre demande
                    Boolean visibledemande = false;

                    try
                    {
                        // try to find the element Nous avons bien enregistré votre demande

                        var boiteblocked4 = driver1.FindElement(By.ClassName("text-title")).Displayed;
                        if (boiteblocked4)
                        {
                            visibledemande = true;
                            string txtblock3 = driver1.FindElement(By.ClassName("text-title")).Text.ToString();
                            if (txtblock3 == "Nous avons bien enregistré votre demande")
                            {

                                //  MessageBox.Show("suivant");
                                System.Threading.Thread.Sleep(1000);
                                driver1.FindElement(By.XPath("//input[@type='button'][@value='Suivant']")).Click();

                                // continue;
                            }
                        }

                    }
                    catch (Exception g)
                    {
                    }
                    try
                    {
                        // open fisrt boite step
                        //ms-Icon ms-Icon--ChevronRight
                        var classfirstopen = driver1.FindElement(By.CssSelector("i[data-icon-name='ChevronRight']")).Displayed;
                        if (classfirstopen)
                        {
                            System.Threading.Thread.Sleep(3000);
                            //  MessageBox.Show("done fisrt");
                            //1
                            //iconButton nextButton lowerButton
                            IWebElement mv1 = driver1.FindElement(By.CssSelector("i[data-icon-name='ChevronRight']"));
                            Actions st10 = new Actions(driver1);
                            st10.MoveToElement(mv1, 40, 16).Click().Build().Perform();

                            System.Threading.Thread.Sleep(3000);
                            //2
                            //iconButton nextButton lowerButton
                            IWebElement mv2 = driver1.FindElement(By.CssSelector("i[data-icon-name='ChevronRight']"));
                            Actions st1 = new Actions(driver1);
                            st1.MoveToElement(mv2, 40, 16).Click().Build().Perform();
                            System.Threading.Thread.Sleep(3000);
                            //3
                            //ms-Icon ms-Icon--ChevronRight
                            IWebElement mv8 = driver1.FindElement(By.CssSelector("i[data-icon-name='ChevronRight']"));
                            Actions st2 = new Actions(driver1);
                            st2.MoveToElement(mv8, 40, 16).Click().Build().Perform();
                            System.Threading.Thread.Sleep(3000);
                            //4
                            //iconButton nextButton
                            IWebElement mv4 = driver1.FindElement(By.CssSelector("i[data-icon-name='ChevronRight']"));
                            Actions st3 = new Actions(driver1);
                            st3.MoveToElement(mv4, 40, 16).Click().Build().Perform();
                            System.Threading.Thread.Sleep(3000);


                            // btn ok 
                            //
                            driver1.FindElement(By.XPath("//div[contains(@class,'ms-Button-label')][contains(text(),'Got it')]")).Click();
                            System.Threading.Thread.Sleep(3000);


                            System.Threading.Thread.Sleep(2000);
                         
                            driver1.Close();
                            driver1.Dispose();
                            continue;

                        }

                    }
                    catch { }

                    try
                    {
                        // open fisrt boite step
                        //ms-Icon ms-Icon--ChevronRight
                        var classfirstopen = driver1.FindElement(By.CssSelector("i[class='ms-Icon ms-Icon--ChevronRight']")).Displayed;
                        if (classfirstopen)
                        {
                            System.Threading.Thread.Sleep(3000);
                            //  MessageBox.Show("done fisrt");
                            //1
                            //iconButton nextButton lowerButton
                            IWebElement mv1 = driver1.FindElement(By.CssSelector("i[class='ms-Icon ms-Icon--ChevronRight']"));
                            Actions st = new Actions(driver1);
                            st.MoveToElement(mv1, 30, 30).Click().Build().Perform();

                            System.Threading.Thread.Sleep(3000);
                            //2
                            //iconButton nextButton lowerButton
                            IWebElement mv2 = driver1.FindElement(By.CssSelector("i[class='ms-Icon ms-Icon--ChevronRight']"));
                            Actions st1 = new Actions(driver1);
                            st1.MoveToElement(mv2, 30, 30).Click().Build().Perform();
                            System.Threading.Thread.Sleep(3000);
                            //3
                            //ms-Icon ms-Icon--ChevronRight
                            IWebElement mv3 = driver1.FindElement(By.CssSelector("i[class='ms-Icon ms-Icon--ChevronRight']"));
                            Actions st2 = new Actions(driver1);
                            st2.MoveToElement(mv3, 30, 30).Click().Build().Perform();
                            System.Threading.Thread.Sleep(3000);
                            //4
                            //iconButton nextButton
                            IWebElement mv4 = driver1.FindElement(By.CssSelector("i[class='ms-Icon ms-Icon--ChevronRight']"));
                            Actions st3 = new Actions(driver1);
                            st3.MoveToElement(mv4, 30, 30).Click().Build().Perform();
                            System.Threading.Thread.Sleep(3000);


                            // btn ok 
                            //
                            driver1.FindElement(By.ClassName("primaryButton")).Click();
                            System.Threading.Thread.Sleep(3000);


                            System.Threading.Thread.Sleep(2000);
                           
                            driver1.Close();
                            driver1.Dispose();
                            continue;

                        }

                    }
                    catch { }

                    //btn enregistrer
                    try
                    {

                        var enregistrer = driver1.FindElement(By.XPath("//span[contains(@class,'signinTxt')][contains(text(),'Enregistrer')]")).Displayed;

                        if (enregistrer)
                        {

                            System.Threading.Thread.Sleep(3000);
                            var select = driver1.FindElement(By.Id("selTz")).Displayed;
                            if (select)
                            {
                                // MessageBox.Show("here");

                            }

                            driver1.FindElement(By.Id("selTz")).FindElement(By.XPath(".//option[contains(text(),'‎(UTC-12:00)‎ Ligne de changement de date internationale ‎(Ouest)‎')]")).Click();
                            System.Threading.Thread.Sleep(3000);
                            driver1.FindElement(By.XPath("//span[contains(@class,'signinTxt')][contains(text(),'Enregistrer')]")).Click();
                            System.Threading.Thread.Sleep(3000);
                         

                        }

                    }
                    catch { }


                    Boolean visible = false;

                    try
                    {
                        // try to find the element 
                        var boiteblocked = driver1.FindElement(By.ClassName("text-title")).Displayed;
                        if (boiteblocked)
                            visible = true;
                        string txtblock3 = driver1.FindElement(By.ClassName("text-title")).Text.ToString();
                        if (txtblock3 == "Aidez-nous à protéger votre compte")
                        {
                          driver1.Close();
                            driver1.Dispose();
                            continue;
                        }
                        if (txtblock3 == "Votre compte a été temporairement bloqué")
                        {
                           driver1.FindElement(By.XPath("//input[@type='submit'][@value='Continuer']")).Click();
                            driver1.Close();
                            driver1.Dispose();
                            continue;
                        }
                        else continue;
                    }
                    catch (Exception g)
                    {
                    }
                    // try for older boite
                    Boolean visible1 = false;
                    try
                    {
                        // try to find the element 
                        var boiteblocked = driver1.FindElement(By.XPath("//span[contains(@class,'o365cs-nav-brandingText')][contains(text(),'Courrier Outlook')]")).Displayed;
                        if (boiteblocked)
                            visible1 = true;
                        string txtblock4 = driver1.FindElement(By.XPath("//span[contains(@class,'o365cs-nav-brandingText')][contains(text(),'Courrier Outlook')]")).Text.ToString();
                        MessageBox.Show(txtblock4);
                        if (txtblock4 == "Courrier Outlook")
                        {
                            versionofmail = "olderversion";

                        }

                    }
                    catch (Exception g)
                    {
                    }
                    // try for new boite
                    Boolean visible2 = false;
                    try
                    {
                        // try to find the element 
                        System.Threading.Thread.Sleep(3000);
                        var boiteblocked1 = driver1.FindElement(By.CssSelector("i[data-icon-name='Waffle']")).Displayed;
                        // if I get to here the element exists 
                        // if it is visible 
                        if (boiteblocked1)
                            visible2 = true;
                        string txtblock = driver1.FindElement(By.CssSelector("i[data-icon-name='Waffle']")).Text.ToString();
                        //  MessageBox.Show(txtblock);

                        if (txtblock == "")
                        {
                            versionofmail = "newversion";
                        }
                    }
                    catch (Exception g)
                    {
                    }

                    System.Threading.Thread.Sleep(1000);


                    Boolean visible6 = false;
                    try
                    {
                        /*
                         *   System.Threading.Thread.Sleep(3000);
                           IWebElement firststep = driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']"));
                           System.Threading.Thread.Sleep(2000);
                           var openlistspam = driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']")).Displayed;
                           System.Threading.Thread.Sleep(2000);
                           firststep.Click();
                           System.Threading.Thread.Sleep(2000);
                         * */
                        // try this code if work
                        // try to enter to inbox folder
                        try
                        {

                            var clickinbox = driver1.FindElement(By.CssSelector("i[data-icon-name='Inbox']")).Displayed;
                            if (clickinbox)
                            {
                                System.Threading.Thread.Sleep(3000);
                                IWebElement firststep = driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']"));
                                System.Threading.Thread.Sleep(2000);
                                var openlistspam = driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']")).Displayed;
                                System.Threading.Thread.Sleep(4000);
                                firststep.Click();
                                System.Threading.Thread.Sleep(3000);
                            }

                        }
                        catch { }
                        try
                        {

                            var clickfrench = driver1.FindElement(By.XPath("//span[contains(@class,'_3lQ0EN5N3oGHxkBKF6_Ane')][contains(text(),'Courrier indésirable')]")).Displayed;
                            if (clickfrench)
                            {
                                System.Threading.Thread.Sleep(3000);
                                IWebElement firststep = driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']"));
                                System.Threading.Thread.Sleep(2000);
                                var openlistspam = driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']")).Displayed;
                                System.Threading.Thread.Sleep(2000);
                                firststep.Click();
                                System.Threading.Thread.Sleep(2000);
                            }

                        }
                        catch { }


                        System.Threading.Thread.Sleep(3000);
                        IWebElement firststep2 = driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']"));
                        System.Threading.Thread.Sleep(2000);
                        var openlistspam2 = driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']")).Displayed;
                        System.Threading.Thread.Sleep(4000);
                        firststep2.Click();
                        System.Threading.Thread.Sleep(3000);

                        driver1.FindElement(By.CssSelector("i[data-icon-name='Inbox']")).Click();
                        //driver1.FindElement(By.XPath("//span[contains(@class,'_3lQ0EN5N3oGHxkBKF6_Ane')][contains(text(),'Courrier indésirable')]")).Click();
                        var inboxshow = driver1.FindElement(By.CssSelector("i[data-icon-name='Inbox']")).Displayed;
                        //  var spamshow = driver1.FindElement(By.XPath("//span[contains(@class,'_3lQ0EN5N3oGHxkBKF6_Ane')][contains(text(),'Courrier indésirable')]")).Displayed;
                        if (inboxshow)
                        {
                            System.Threading.Thread.Sleep(2000);
                            // last try
                            try
                            {
                                try
                                {

                                    var dossiervide = driver1.FindElement(By.XPath("//div[contains(@class,'_25o_l9GBJ9r8kRewkYoBfm _2oemjr1JNV5H9BUBUobbw2')][contains(text(),'Vous n’avez plus aucun message dans votre boîte de réception Prioritaire')]")).Text.ToString();
                                    if (dossiervide == "Ce dossier Inbox est vide")
                                    {
                                        // MessageBox.Show("vide");
                                       driver1.Close();
                                        driver1.Dispose();
                                        continue;
                                    }

                                }
                                catch { }


                                try
                                {
                                    var dossiervide1 = driver1.FindElement(By.XPath("//div[contains(@class,'_25o_l9GBJ9r8kRewkYoBfm _2oemjr1JNV5H9BUBUobbw2')][contains(text(),'This folder is empty')]")).Text.ToString();
                                    if (dossiervide1 == "This folder is empty")
                                    {
                                        // MessageBox.Show("vide");

                                       driver1.Close();
                                        driver1.Dispose();
                                        continue;
                                    }

                                }
                                catch { }
                                // try to find the element 
                                // MessageBox.Show("inbox clk!");
                                System.Threading.Thread.Sleep(2000);
                                var markasread = driver1.FindElement(By.CssSelector("i[data-icon-name='Inbox']")).Displayed;
                                // if I get to here the element exists 
                                // if it is visible 
                                System.Threading.Thread.Sleep(2000);

                                if (markasread)
                                {
                                    visible2 = true;
                                    string markasread1inbox = driver1.FindElement(By.CssSelector("i[data-icon-name='Inbox']")).Text.ToString();
                                    //  MessageBox.Show(txtblock);

                                    if (markasread1inbox == "")
                                    {
                                        //MessageBox.Show("inbox clk 1!");
                                        //MessageBox.Show("comme lu");
                                        // try to read spam and move it
                                        try
                                        {
                                            System.Threading.Thread.Sleep(2000);

                                            var boiteblocked6 = driver1.FindElement(By.ClassName("_2EHjCdO2IEh-zlrH2jOD50")).Displayed;
                                            System.Threading.Thread.Sleep(2000);
                                            //   string fe = driver1.FindElement(By.XPath("//span[contains(@class,'ms-Button-label label-55')][contains(text(),'Marquer tout comme lu')]")).Text.ToString();
                                            //  MessageBox.Show(fe.ToString());
                                            // Ce dossier est vide
                                            do
                                            {

                                                if (boiteblocked6)
                                                {
                                                    //  MessageBox.Show("inbox dispaled");
                                                    // Boolean visible1 = false;  // assume it is invisible 
                                                    /* IWebElement firststep = driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']"));
                                                     firststep.Click();*/
                                                    //*********************code for new version 
                                                    System.Threading.Thread.Sleep(2000);
                                                    //  driver1.FindElement(By.CssSelector("i[data-icon-name='Blocked']")).Click();
                                                    //driver1.FindElement(By.XPath("//span[contains(@class,'_3lQ0EN5N3oGHxkBKF6_Ane')][contains(text(),'Courrier indésirable')]")).Click();
                                                    System.Threading.Thread.Sleep(2000);
                                                    driver1.FindElement(By.ClassName("_2EHjCdO2IEh-zlrH2jOD50")).Click();
                                                    System.Threading.Thread.Sleep(1500);
                                                    /* driver1.FindElement(By.CssSelector("i[data-icon-name='ChevronDown']")).Click();
                                                     IJavaScriptExecutor js1 = (IJavaScriptExecutor)driver1;
                                                    
                                                     System.Threading.Thread.Sleep(1500);
                                                    // IWebElement Element1 = driver1.FindElement(By.XPath("//span[contains(@class,'ms-ContextualMenu-itemText label-170')][contains(text(),'Courrier légitime')]"));
                                                     IWebElement Elementbtn = driver1.FindElement(By.XPath("//button[contains(@class,'ms-ContextualMenu-link root-146')][contains(@name,'Courrier légitime')]"));
                                                     MessageBox.Show(Elementbtn.ToString());
                                                       
                                                     System.Threading.Thread.Sleep(15000);
                                                    // IWebElement Element = driver1.FindElement(By.XPath("//span[contains(@class,'ms-ContextualMenu-itemText label-170')][contains(text(),'Not spam')]"));
                                                     //This will scroll the page till the element is found		

                                                     js1.ExecuteScript("arguments[0].scrollIntoView();", Elementbtn);
                                                     Elementbtn.Click();*/
                                                    // MessageBox.Show("down now !");
                                                    int numberofinbox = 0;
                                                 
                                                  


                                                    System.Threading.Thread.Sleep(1500);

                                                    // new try version anglais 


                                                    // end try version anglais

                                                    System.Threading.Thread.Sleep(2500);
                                                    visible6 = false;

                                                }

                                            } while (markasread);

                                        }

                                        catch
                                        {

                                        }
                                       driver1.Dispose();
                                        /*  driver1.Close();
                                          driver1.Dispose();*/
                                        continue;
                                    }

                                    //  MessageBox.Show("here");


                                }

                                if (markasread == false)
                                {

                                 
                                }
                            }
                            catch (Exception g)
                            {
                            }

                            // last try
                        }

                        // close this else 
                        /* else
                         {
                             driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']")).Click();
                         }*/

                    }
                    catch
                    {
                    }
                    try
                    {
                        driver1.Close();
                        driver1.Dispose();
                    }
                    catch { }

                    //  driver1.FindElement(By.XPath("//span[contains(@class,'_3lQ0EN5N3oGHxkBKF6_Ane')][contains(text(),'Courrier indésirable')]")).Click();
                    System.Threading.Thread.Sleep(2000);

                }

                else
                {
                   
                }
            }
            int reportinglistnumber = 0;
            /*  StringBuilder listViewContent = new StringBuilder();
           
           
              foreach (ListViewItem item in this.listView1.Items)
              {
                  listViewContent.Append(item);
                  listViewContent.Append(Environment.NewLine);
                
                  TextWriter tw = new StreamWriter("C:/Users/Public/ReportingResult" + reportinglistnumber+".txt");

                  tw.WriteLine(listViewContent.ToString());
             
                  tw.Close();
              }*/
            string sPath = "C:/Users/Public/ReportingResultInbox" + reportinglistnumber + ".txt";

            System.IO.StreamWriter SaveFile = new System.IO.StreamWriter(sPath);


           

            /*listBox2.Text = "";
            listBox3.Text = "";*/

            // done 
         
        }


        private void button27_Click(object sender, EventArgs e)
        {
            ImageList imgs = new ImageList();
            imgs.ImageSize = new Size(50, 50);
            String[] pathss = { };
            pathss = Directory.GetFiles("C:/imgs");


            // buttun test sans imap
            string line;
            int number = 0;
            System.IO.StreamReader file =
    new System.IO.StreamReader(path);

            while ((line = file.ReadLine()) != null)
            {
                string[] lines = line.Split(',');

                string ips = lines[0];
                // MessageBox.Show(lines[0]);
                string boite = lines[1];
                string pass = lines[2];

                number++;
                string[] ipall = ips.Split(':');
                int port = Convert.ToInt32(ipall[1]);
                bool tt = PingHost(ipall[0], port);
                if (tt == true)
                {

                    // MessageBox.Show("tt True");


                    //**************************************

                    var chromeOptions = new ChromeOptions();
                    //Create a new proxy object
                    var proxy = new Proxy();
                  
                    var driver1 = new ChromeDriver(@"C:\webdrivers", chromeOptions);
                    //***********work without proxy******************
                    //var driver1 = new ChromeDriver(@"C:\webdrivers");
                    driver1.Navigate().GoToUrl(BaseUrl);
                    driver1.FindElementByLinkText("Se connecter").Click();

                    var loginBox = driver1.FindElement(By.XPath("//input[contains(text(),'')]"));
                    System.Threading.Thread.Sleep(2500);
                    loginBox.SendKeys(boite);
                    driver1.FindElement(By.XPath("//input[@type='submit'][@value='Suivant']")).Click();
                    System.Threading.Thread.Sleep(2500);
                    try
                    {
                        var suiv = driver1.FindElement(By.XPath("//input[@type='submit'][@value='Suivant']")).Displayed;
                        if (suiv)
                        {
                            /*  listView1.SmallImageList = imgs;
                              listView1.Items.Add("Michael Carrick", 1);*/
                            // listView1.Items.Add(boite + "User Inccorect ! Please Try Again");
                            driver1.Close();
                            driver1.Dispose();
                            continue;
                        }

                    }
                    catch { }

                    var passwordBox = driver1.FindElement(By.XPath("//input[@id='i0118' and contains(text(),'')]"));
                    passwordBox.SendKeys(pass);
                    System.Threading.Thread.Sleep(2500);
                    try
                    {
                        driver1.FindElement(By.XPath("//input[@type='submit'][@value='Se connecter']")).Click();
                        System.Threading.Thread.Sleep(4000);
                        driver1.Manage().Window.Maximize();
                    }
                    catch
                    {
                    }
                    System.Threading.Thread.Sleep(4000);
                    // driver1.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);

                    try
                    {

                        //
                        //   MessageBox.Show("he 1");
                        System.Threading.Thread.Sleep(2000);
                        IJavaScriptExecutor js1 = (IJavaScriptExecutor)driver1;

                        IWebElement Element14 = driver1.FindElement(By.CssSelector("i[data-icon-name='Cancel']"));
                        var tst = driver1.FindElement(By.CssSelector("i[data-icon-name='Cancel']")).Displayed;
                        System.Threading.Thread.Sleep(2000);
                        //   MessageBox.Show(tst.ToString());
                        js1.ExecuteScript("arguments[0].scrollIntoView();", Element14);
                        Element14.Click();
                        System.Threading.Thread.Sleep(1000);
                        // driver1.FindElement(By.ClassName("_3QWMcW9VDbQi11Tl_MkRml")).Click();



                    }

                    catch { }



                    System.Threading.Thread.Sleep(2000);
                    try
                    {

                        System.Threading.Thread.Sleep(1500);
                        var suiv1 = driver1.FindElement(By.XPath("//input[@id='i0118' and contains(text(),'')]")).Displayed;
                        // string fde = driver1.FindElement(By.ClassName("alert alert-error")).Text.ToString();
                        System.Threading.Thread.Sleep(1000);
                        //   MessageBox.Show(fde);
                        if (suiv1)
                        {
                          
                        }

                    }
                    catch { }
                    //Nous avons bien enregistré votre demande
                    Boolean visibledemande = false;

                    try
                    {
                        // try to find the element Nous avons bien enregistré votre demande

                        var boiteblocked4 = driver1.FindElement(By.ClassName("text-title")).Displayed;
                        if (boiteblocked4)
                        {
                            visibledemande = true;
                            string txtblock3 = driver1.FindElement(By.ClassName("text-title")).Text.ToString();
                            if (txtblock3 == "Nous avons bien enregistré votre demande")
                            {

                                //  MessageBox.Show("suivant");
                                System.Threading.Thread.Sleep(1000);
                                driver1.FindElement(By.XPath("//input[@type='button'][@value='Suivant']")).Click();

                                // continue;
                            }
                        }

                    }
                    catch (Exception g)
                    {
                    }


                    try
                    {
                        // open fisrt boite step
                        //ms-Icon ms-Icon--ChevronRight
                        var classfirstopen = driver1.FindElement(By.CssSelector("i[class='ms-Icon ms-Icon--ChevronRight']")).Displayed;
                        if (classfirstopen)
                        {
                            System.Threading.Thread.Sleep(3000);
                            //  MessageBox.Show("done fisrt");
                            //1
                            //iconButton nextButton lowerButton
                            IWebElement mv1 = driver1.FindElement(By.CssSelector("i[class='ms-Icon ms-Icon--ChevronRight']"));
                            Actions st = new Actions(driver1);
                            st.MoveToElement(mv1, 30, 30).Click().Build().Perform();

                            System.Threading.Thread.Sleep(3000);
                            //2
                            //iconButton nextButton lowerButton
                            IWebElement mv2 = driver1.FindElement(By.CssSelector("i[class='ms-Icon ms-Icon--ChevronRight']"));
                            Actions st1 = new Actions(driver1);
                            st1.MoveToElement(mv2, 30, 30).Click().Build().Perform();
                            System.Threading.Thread.Sleep(3000);
                            //3
                            //ms-Icon ms-Icon--ChevronRight
                            IWebElement mv3 = driver1.FindElement(By.CssSelector("i[class='ms-Icon ms-Icon--ChevronRight']"));
                            Actions st2 = new Actions(driver1);
                            st2.MoveToElement(mv3, 30, 30).Click().Build().Perform();
                            System.Threading.Thread.Sleep(3000);
                            //4
                            //iconButton nextButton
                            IWebElement mv4 = driver1.FindElement(By.CssSelector("i[class='ms-Icon ms-Icon--ChevronRight']"));
                            Actions st3 = new Actions(driver1);
                            st3.MoveToElement(mv4, 30, 30).Click().Build().Perform();
                            System.Threading.Thread.Sleep(3000);


                            // btn ok 
                            //
                            driver1.FindElement(By.ClassName("primaryButton")).Click();
                            System.Threading.Thread.Sleep(3000);


                          

                        }

                    }
                    catch { }

                    //btn enregistrer
                    try
                    {

                        var enregistrer = driver1.FindElement(By.XPath("//span[contains(@class,'signinTxt')][contains(text(),'Enregistrer')]")).Displayed;

                        if (enregistrer)
                        {

                            System.Threading.Thread.Sleep(3000);
                            var select = driver1.FindElement(By.Id("selTz")).Displayed;
                            if (select)
                            {
                                // MessageBox.Show("here");

                            }

                            driver1.FindElement(By.Id("selTz")).FindElement(By.XPath(".//option[contains(text(),'‎(UTC-12:00)‎ Ligne de changement de date internationale ‎(Ouest)‎')]")).Click();
                            System.Threading.Thread.Sleep(3000);
                            driver1.FindElement(By.XPath("//span[contains(@class,'signinTxt')][contains(text(),'Enregistrer')]")).Click();
                            System.Threading.Thread.Sleep(3000);
                             driver1.Close();
                            driver1.Dispose();
                            continue;


                        }

                    }
                    catch { }


                    Boolean visible = false;


                    try
                    {
                        // try to find the element 
                        var boiteblocked = driver1.FindElement(By.ClassName("text-title")).Displayed;
                        if (boiteblocked)
                            visible = true;
                        string txtblock3 = driver1.FindElement(By.ClassName("text-title")).Text.ToString();
                        if (txtblock3 == "Aidez-nous à protéger votre compte")
                        {
                         
                        }
                        if (txtblock3 == "Votre compte a été temporairement bloqué")
                        {

                          driver1.FindElement(By.XPath("//input[@type='submit'][@value='Continuer']")).Click();
                            driver1.Close();
                            driver1.Dispose();
                            continue;
                        }
                        else continue;
                    }
                    catch (Exception g)
                    {
                    }
                    //try for desactiver 2 type 
                    try
                    {
                        string typeofboite = driver1.FindElement(By.XPath("//div[contains(@class,'ms-Button-label')][contains(text(),'New message')]")).Text.ToString();
                        if (typeofboite == "New message")
                        {

                            //desiactiver part anglais 
                            try
                            {
                                System.Threading.Thread.Sleep(4000);
                                // for clk in option to enable prioritaire for new version outlook
                                IWebElement ell = driver1.FindElement(By.CssSelector("i[data-icon-name='Settings']"));
                                //MessageBox.Show(ell.Location.ToString());
                                ell.Click();
                                System.Threading.Thread.Sleep(2000);
                                IJavaScriptExecutor js2 = (IJavaScriptExecutor)driver1;
                                IWebElement elmasquer = driver1.FindElement(By.XPath("//span[contains(@class,'ms-Label')][contains(text(),'Hide')]"));
                                //   IWebElement Element = driver1.FindElement(By.XPath("//button[@id='Toggle111']"));
                                //This will scroll the page till the element is found		

                                js2.ExecuteScript("arguments[0].scrollIntoView();", elmasquer);

                                // driver1.FindElement(By.XPath("//button[@id='Toggle111']")).Click();
                                driver1.FindElement(By.XPath("//span[contains(@class,'ms-Label')][contains(text(),'Masquer')]")).Click();

                                System.Threading.Thread.Sleep(2000);
                                IJavaScriptExecutor js1 = (IJavaScriptExecutor)driver1;
                                //
                                IWebElement elpr = driver1.FindElement(By.CssSelector("i[data-icon-name='ChevronRight']"));
                                // IWebElement elpr = driver1.FindElement(By.ClassName("ms-Toggle-thumb thumb-155"));
                                //   IWebElement Element = driver1.FindElement(By.XPath("//button[@id='Toggle111']"));
                                //This will scroll the page till the element is found		


                                js1.ExecuteScript("arguments[0].scrollIntoView();", elpr);
                                /* Point point = elpr.Location;
                                 int x = point.X;
                                 int y = point.Y;
                                 MessageBox.Show(x.ToString() + "" + y.ToString());*/
                                elpr.Click();
                                //driver1.FindElement(By.XPath("//span[contains(@class,'ms-Label')][contains(text(),'Masquer')]")).Click();
                                System.Threading.Thread.Sleep(2000);
                                IJavaScriptExecutor js3 = (IJavaScriptExecutor)driver1;
                                IWebElement triemsg = driver1.FindElement(By.XPath("//span[contains(@class,'ms-Label')][contains(text(),'Don't sort my messages')]"));
                                js3.ExecuteScript("arguments[0].scrollIntoView();", triemsg);
                                System.Threading.Thread.Sleep(3000);
                                triemsg.Click();
                                //
                                driver1.FindElement(By.XPath("//div[contains(@class,'ms-Button-label')][contains(text(),'Save')]")).Click();
                                System.Threading.Thread.Sleep(3000);

                                // close page of option
                                IWebElement ellclose = driver1.FindElement(By.CssSelector("i[data-icon-name='Cancel']"));
                                //MessageBox.Show(ell.Location.ToString());
                                ellclose.Click();
                            }
                            catch
                            {

                            }
                            // desactiver prioritaire
                        }
                        else
                        {
                            // desactiver priortaire part 2
                            try
                            {
                                System.Threading.Thread.Sleep(4000);
                                // for clk in option to enable prioritaire for new version outlook
                                IWebElement ell = driver1.FindElement(By.CssSelector("i[data-icon-name='Settings']"));
                                //MessageBox.Show(ell.Location.ToString());
                                ell.Click();
                                System.Threading.Thread.Sleep(2000);
                                IJavaScriptExecutor js2 = (IJavaScriptExecutor)driver1;
                                IWebElement elmasquer = driver1.FindElement(By.XPath("//span[contains(@class,'ms-Label')][contains(text(),'Masquer')]"));
                                //   IWebElement Element = driver1.FindElement(By.XPath("//button[@id='Toggle111']"));
                                //This will scroll the page till the element is found		

                                js2.ExecuteScript("arguments[0].scrollIntoView();", elmasquer);

                                // driver1.FindElement(By.XPath("//button[@id='Toggle111']")).Click();
                                driver1.FindElement(By.XPath("//span[contains(@class,'ms-Label')][contains(text(),'Masquer')]")).Click();

                                System.Threading.Thread.Sleep(2000);
                                IJavaScriptExecutor js1 = (IJavaScriptExecutor)driver1;
                                //
                                IWebElement elpr = driver1.FindElement(By.CssSelector("i[data-icon-name='ChevronRight']"));
                                // IWebElement elpr = driver1.FindElement(By.ClassName("ms-Toggle-thumb thumb-155"));
                                //   IWebElement Element = driver1.FindElement(By.XPath("//button[@id='Toggle111']"));
                                //This will scroll the page till the element is found		


                                js1.ExecuteScript("arguments[0].scrollIntoView();", elpr);
                                /* Point point = elpr.Location;
                                 int x = point.X;
                                 int y = point.Y;
                                 MessageBox.Show(x.ToString() + "" + y.ToString());*/
                                elpr.Click();
                                //driver1.FindElement(By.XPath("//span[contains(@class,'ms-Label')][contains(text(),'Masquer')]")).Click();
                                System.Threading.Thread.Sleep(2000);
                                IJavaScriptExecutor js3 = (IJavaScriptExecutor)driver1;
                                IWebElement triemsg = driver1.FindElement(By.XPath("//span[contains(@class,'ms-Label')][contains(text(),'Ne pas trier mes messages')]"));
                                js3.ExecuteScript("arguments[0].scrollIntoView();", triemsg);
                                System.Threading.Thread.Sleep(3000);
                                triemsg.Click();
                                //
                                driver1.FindElement(By.XPath("//div[contains(@class,'ms-Button-label')][contains(text(),'Enregistrer')]")).Click();
                                System.Threading.Thread.Sleep(3000);

                                // close page of option
                                IWebElement ellclose = driver1.FindElement(By.CssSelector("i[data-icon-name='Cancel']"));
                                //MessageBox.Show(ell.Location.ToString());
                                ellclose.Click();
                            }
                            catch
                            {

                            }
                            // desactiver prioritaire
                        }


                    }
                    catch { }


                    //

                    /*  try
                      {
                          System.Threading.Thread.Sleep(4000);
                          // for clk in option to enable prioritaire for new version outlook
                          IWebElement ell = driver1.FindElement(By.CssSelector("i[data-icon-name='Settings']"));
                          //MessageBox.Show(ell.Location.ToString());
                          ell.Click();
                          System.Threading.Thread.Sleep(2000);
                          IJavaScriptExecutor js = (IJavaScriptExecutor)driver1;
                          //
                          IWebElement elmasquer = driver1.FindElement(By.XPath("//span[contains(@class,'ms-Label')][contains(text(),'Masquer')]"));
                       //   IWebElement Element = driver1.FindElement(By.XPath("//button[@id='Toggle111']"));
                          //This will scroll the page till the element is found		

                          js.ExecuteScript("arguments[0].scrollIntoView();", elmasquer);

                         // driver1.FindElement(By.XPath("//button[@id='Toggle111']")).Click();
                          driver1.FindElement(By.XPath("//span[contains(@class,'ms-Label')][contains(text(),'Masquer')]")).Click();
                          System.Threading.Thread.Sleep(1000);
                          // close page of option
                          /*IWebElement ellclose = driver1.FindElement(By.CssSelector("i[data-icon-name='Cancel']"));
                          //MessageBox.Show(ell.Location.ToString());
                          ellclose.Click();*/

                    //
                    /*  System.Threading.Thread.Sleep(4000);
                      // for clk in option to enable prioritaire for new version outlook
                      IWebElement ellpr = driver1.FindElement(By.CssSelector("i[data-icon-name='Settings']"));
                      //MessageBox.Show(ell.Location.ToString());
                      ellpr.Click();


                      System.Threading.Thread.Sleep(2000);
                        
                       

                  }
                  catch
                  { }*/
                    // try for older boite
                    Boolean visible1 = false;

                    // try for new boite
                    Boolean visible2 = false;

                    System.Threading.Thread.Sleep(1000);


                    Boolean visible6 = false;



                    //  driver1.FindElement(By.XPath("//span[contains(@class,'_3lQ0EN5N3oGHxkBKF6_Ane')][contains(text(),'Courrier indésirable')]")).Click();
                    System.Threading.Thread.Sleep(2000);
System.Threading.Thread.Sleep(2000);
                    driver1.Close();
                    driver1.Dispose();

                }


                else
                {
                  
                }

            }
            int reportinglistnumber = 0;
            /*  StringBuilder listViewContent = new StringBuilder();
           
           
              foreach (ListViewItem item in this.listView1.Items)
              {
                  listViewContent.Append(item);
                  listViewContent.Append(Environment.NewLine);
                
                  TextWriter tw = new StreamWriter("C:/Users/Public/ReportingResult" + reportinglistnumber+".txt");

                  tw.WriteLine(listViewContent.ToString());
             
                  tw.Close();
              }*/
            string sPath = "C:/Users/Public/ReportingResultPrioritaire" + reportinglistnumber + ".txt";

            System.IO.StreamWriter SaveFile = new System.IO.StreamWriter(sPath);


           
            SaveFile.Close();
            // done 
           
            
            MessageBox.Show("Reporting Done!");

        }

        private void button28_Click(object sender, EventArgs e)
        {
            string line;
            int number = 0;
            System.IO.StreamReader file =
    new System.IO.StreamReader(path);

            while ((line = file.ReadLine()) != null)
            {
                string[] lines = line.Split(',');

                string ips = lines[0];
                // MessageBox.Show(lines[0]);
                string boite = lines[1];
                string pass = lines[2];

                number++;
                string[] ipall = ips.Split(':');
                int port = Convert.ToInt32(ipall[1]);
                bool tt = PingHost(ipall[0], port);
                if (tt == true)
                {

                    // MessageBox.Show("tt True");


                    //**************************************

                    var chromeOptions = new ChromeOptions();
                    //Create a new proxy object
                    var proxy = new Proxy();
                    //Set the http proxy value, host and port.
                    // proxy.HttpProxy = ips;
                    //Set the proxy to the Chrome options
                    //  chromeOptions.Proxy = proxy;
                  
                    var driver1 = new ChromeDriver(@"C:\webdrivers", chromeOptions);
                    //***********work without proxy******************
                    //var driver1 = new ChromeDriver(@"C:\webdrivers");
                    driver1.Navigate().GoToUrl(BaseUrl);
                    driver1.FindElementByLinkText("Se connecter").Click();

                    var loginBox = driver1.FindElement(By.XPath("//input[contains(text(),'')]"));
                    System.Threading.Thread.Sleep(2500);
                    loginBox.SendKeys(boite);
                    driver1.FindElement(By.XPath("//input[@type='submit'][@value='Suivant']")).Click();
                    System.Threading.Thread.Sleep(2500);
                    try
                    {
                        var suiv = driver1.FindElement(By.XPath("//input[@type='submit'][@value='Suivant']")).Displayed;
                        if (suiv)
                        {
                            /*  listView1.SmallImageList = imgs;
                              listView1.Items.Add("Michael Carrick", 1);*/
                          // listView1.Items.Add(boite + "User Inccorect ! Please Try Again");
                            driver1.Close();
                            driver1.Dispose();
                            continue;
                        }

                    }
                    catch { }

                    var passwordBox = driver1.FindElement(By.XPath("//input[@id='i0118' and contains(text(),'')]"));
                    passwordBox.SendKeys(pass);
                    System.Threading.Thread.Sleep(2500);
                    driver1.FindElement(By.XPath("//input[@type='submit'][@value='Se connecter']")).Click();
                    //System.Threading.Thread.Sleep(10000);
                    driver1.FindElement(By.ClassName("_29AA3OWLpLPsgv7R6xoYtp")).Click();
                    System.Threading.Thread.Sleep(3000);
                    /* try
                     {

                         driver1.FindElementByClassName("_29AA3OWLpLPsgv7R6xoYtp").Click();
                     }

                     catch { }*/
                    driver1.Manage().Window.Maximize();
                    // System.Threading.Thread.Sleep(10000);
                    System.Threading.Thread.Sleep(1000);

                    //driver1.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);
                    try
                    {

                        //
                        //   MessageBox.Show("he 1");
                        System.Threading.Thread.Sleep(2000);
                        IJavaScriptExecutor js1 = (IJavaScriptExecutor)driver1;

                        IWebElement Element14 = driver1.FindElement(By.CssSelector("i[data-icon-name='Cancel']"));
                        var tst = driver1.FindElement(By.CssSelector("i[data-icon-name='Cancel']")).Displayed;
                        System.Threading.Thread.Sleep(2000);
                        //   MessageBox.Show(tst.ToString());
                        js1.ExecuteScript("arguments[0].scrollIntoView();", Element14);
                        Element14.Click();
                        System.Threading.Thread.Sleep(1000);
                        // driver1.FindElement(By.ClassName("_3QWMcW9VDbQi11Tl_MkRml")).Click();



                    }

                    catch { }
                    //driver1.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(6);




                    System.Threading.Thread.Sleep(2000);
                    try
                    {

                        System.Threading.Thread.Sleep(1500);
                        var suiv1 = driver1.FindElement(By.XPath("//input[@id='i0118' and contains(text(),'')]")).Displayed;
                        // string fde = driver1.FindElement(By.ClassName("alert alert-error")).Text.ToString();
                        System.Threading.Thread.Sleep(1000);
                        //   MessageBox.Show(fde);
                        if (suiv1)
                        {
                            // MessageBox.Show("");
                            driver1.Close();
                            driver1.Dispose();
                            continue;
                        }

                    }
                    catch { }
                    //Nous avons bien enregistré votre demande
                    Boolean visibledemande = false;

                    try
                    {
                        // try to find the element Nous avons bien enregistré votre demande

                        var boiteblocked4 = driver1.FindElement(By.ClassName("text-title")).Displayed;
                        if (boiteblocked4)
                        {
                            visibledemande = true;
                            string txtblock3 = driver1.FindElement(By.ClassName("text-title")).Text.ToString();
                            if (txtblock3 == "Nous avons bien enregistré votre demande")
                            {

                                //  MessageBox.Show("suivant");
                                System.Threading.Thread.Sleep(1000);
                                driver1.FindElement(By.XPath("//input[@type='button'][@value='Suivant']")).Click();
                                System.Threading.Thread.Sleep(1000);
                                driver1.Close();
                                driver1.Dispose();
                                // continue;
                            }
                        }

                    }
                    catch (Exception g)
                    {
                    }


                    try
                    {
                        // open fisrt boite step
                        //ms-Icon ms-Icon--ChevronRight
                        var classfirstopen = driver1.FindElement(By.CssSelector("i[class='ms-Icon ms-Icon--ChevronRight']")).Displayed;
                        if (classfirstopen)
                        {
                            System.Threading.Thread.Sleep(3000);
                            //  MessageBox.Show("done fisrt");
                            //1
                            //iconButton nextButton lowerButton
                            IWebElement mv1 = driver1.FindElement(By.CssSelector("i[class='ms-Icon ms-Icon--ChevronRight']"));
                            Actions st = new Actions(driver1);
                            st.MoveToElement(mv1, 30, 30).Click().Build().Perform();

                            System.Threading.Thread.Sleep(3000);
                            //2
                            //iconButton nextButton lowerButton
                            IWebElement mv2 = driver1.FindElement(By.CssSelector("i[class='ms-Icon ms-Icon--ChevronRight']"));
                            Actions st1 = new Actions(driver1);
                            st1.MoveToElement(mv2, 30, 30).Click().Build().Perform();
                            System.Threading.Thread.Sleep(3000);
                            //3
                            //ms-Icon ms-Icon--ChevronRight
                            IWebElement mv3 = driver1.FindElement(By.CssSelector("i[class='ms-Icon ms-Icon--ChevronRight']"));
                            Actions st2 = new Actions(driver1);
                            st2.MoveToElement(mv3, 30, 30).Click().Build().Perform();
                            System.Threading.Thread.Sleep(3000);
                            //4
                            //iconButton nextButton
                            IWebElement mv4 = driver1.FindElement(By.CssSelector("i[class='ms-Icon ms-Icon--ChevronRight']"));
                            Actions st3 = new Actions(driver1);
                            st3.MoveToElement(mv4, 30, 30).Click().Build().Perform();
                            System.Threading.Thread.Sleep(3000);


                            // btn ok 
                            //
                            driver1.FindElement(By.ClassName("primaryButton")).Click();
                            System.Threading.Thread.Sleep(3000);


                            System.Threading.Thread.Sleep(2000);
                           
                            continue;

                        }

                    }
                    catch { }

                    //btn enregistrer
                    try
                    {

                        var enregistrer = driver1.FindElement(By.XPath("//span[contains(@class,'signinTxt')][contains(text(),'Enregistrer')]")).Displayed;

                        if (enregistrer)
                        {

                            System.Threading.Thread.Sleep(3000);
                            var select = driver1.FindElement(By.Id("selTz")).Displayed;
                            if (select)
                            {
                                // MessageBox.Show("here");

                            }

                            driver1.FindElement(By.Id("selTz")).FindElement(By.XPath(".//option[contains(text(),'‎(UTC-12:00)‎ Ligne de changement de date internationale ‎(Ouest)‎')]")).Click();
                            System.Threading.Thread.Sleep(3000);
                            driver1.FindElement(By.XPath("//span[contains(@class,'signinTxt')][contains(text(),'Enregistrer')]")).Click();
                            System.Threading.Thread.Sleep(3000);
                          


                        }

                    }
                    catch { }


                    Boolean visible = false;

                    try
                    {
                        // try to find the element 
                        var boiteblocked = driver1.FindElement(By.ClassName("text-title")).Displayed;
                        if (boiteblocked)
                            visible = true;
                        string txtblock3 = driver1.FindElement(By.ClassName("text-title")).Text.ToString();
                        if (txtblock3 == "Aidez-nous à protéger votre compte")
                        {
                         
                        }
                        if (txtblock3 == "Votre compte a été temporairement bloqué")
                        {
                        }
                        else continue;
                    }
                    catch (Exception g)
                    {
                    }
                    // try for older boite
                    Boolean visible1 = false;
                    try
                    {
                        // try to find the element 
                        var boiteblocked = driver1.FindElement(By.XPath("//span[contains(@class,'o365cs-nav-brandingText')][contains(text(),'Courrier Outlook')]")).Displayed;
                        if (boiteblocked)
                            visible1 = true;
                        string txtblock4 = driver1.FindElement(By.XPath("//span[contains(@class,'o365cs-nav-brandingText')][contains(text(),'Courrier Outlook')]")).Text.ToString();
                        MessageBox.Show(txtblock4);
                        if (txtblock4 == "Courrier Outlook")
                        {
                            versionofmail = "olderversion";

                        }

                    }
                    catch (Exception g)
                    {
                    }
                    // try for new boite
                    Boolean visible2 = false;
                    try
                    {
                        // try to find the element 
                        System.Threading.Thread.Sleep(2000);
                        var boiteblocked1 = driver1.FindElement(By.CssSelector("i[data-icon-name='Waffle']")).Displayed;
                        // if I get to here the element exists 
                        // if it is visible 
                        if (boiteblocked1)
                            visible2 = true;
                        string txtblock = driver1.FindElement(By.CssSelector("i[data-icon-name='Waffle']")).Text.ToString();
                        //  MessageBox.Show(txtblock);

                        if (txtblock == "")
                        {
                            versionofmail = "newversion";
                        }
                    }
                    catch (Exception g)
                    {
                    }

                    System.Threading.Thread.Sleep(1000);


                    Boolean visible6 = false;
                    try
                    {
                        /*
                         *   System.Threading.Thread.Sleep(3000);
                           IWebElement firststep = driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']"));
                           System.Threading.Thread.Sleep(2000);
                           var openlistspam = driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']")).Displayed;
                           System.Threading.Thread.Sleep(2000);
                           firststep.Click();
                           System.Threading.Thread.Sleep(2000);
                         * */
                        // try this code if work
                        try
                        {

                            var clickenglich = driver1.FindElement(By.XPath("//span[contains(@class,'_3lQ0EN5N3oGHxkBKF6_Ane')][contains(text(),'Junk Email')]")).Displayed;
                            if (clickenglich)
                            {
                                System.Threading.Thread.Sleep(3000);
                                IWebElement firststep = driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']"));
                                System.Threading.Thread.Sleep(2000);
                                var openlistspam = driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']")).Displayed;
                                System.Threading.Thread.Sleep(2000);
                                firststep.Click();
                                System.Threading.Thread.Sleep(2000);
                            }

                        }
                        catch { }
                        try
                        {

                            var clickfrench = driver1.FindElement(By.XPath("//span[contains(@class,'_3lQ0EN5N3oGHxkBKF6_Ane')][contains(text(),'Courrier indésirable')]")).Displayed;
                            if (clickfrench)
                            {
                                System.Threading.Thread.Sleep(3000);
                                IWebElement firststep = driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']"));
                                System.Threading.Thread.Sleep(2000);
                                var openlistspam = driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']")).Displayed;
                                System.Threading.Thread.Sleep(2000);
                                firststep.Click();
                                System.Threading.Thread.Sleep(2000);
                            }

                        }
                        catch { }


                        System.Threading.Thread.Sleep(3000);
                        IWebElement firststep2 = driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']"));
                        System.Threading.Thread.Sleep(2000);
                        var openlistspam2 = driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']")).Displayed;
                        System.Threading.Thread.Sleep(2000);
                        firststep2.Click();
                        System.Threading.Thread.Sleep(2000);

                        driver1.FindElement(By.CssSelector("i[data-icon-name='Blocked']")).Click();
                        //driver1.FindElement(By.XPath("//span[contains(@class,'_3lQ0EN5N3oGHxkBKF6_Ane')][contains(text(),'Courrier indésirable')]")).Click();
                        var spamshow = driver1.FindElement(By.CssSelector("i[data-icon-name='Blocked']")).Displayed;
                        //  var spamshow = driver1.FindElement(By.XPath("//span[contains(@class,'_3lQ0EN5N3oGHxkBKF6_Ane')][contains(text(),'Courrier indésirable')]")).Displayed;
                        if (spamshow)
                        {
                            System.Threading.Thread.Sleep(2000);
                            // last try
                            try
                            {
                                try
                                {

                                    var dossiervide = driver1.FindElement(By.XPath("//div[contains(@class,'_25o_l9GBJ9r8kRewkYoBfm _2oemjr1JNV5H9BUBUobbw2')][contains(text(),'Ce dossier est vide')]")).Text.ToString();
                                    if (dossiervide == "Ce dossier est vide")
                                    {
                                        // MessageBox.Show("vide");
                                       
                                    }

                                }
                                catch { }


                                try
                                {
                                    var dossiervide1 = driver1.FindElement(By.XPath("//div[contains(@class,'_25o_l9GBJ9r8kRewkYoBfm _2oemjr1JNV5H9BUBUobbw2')][contains(text(),'This folder is empty')]")).Text.ToString();
                                    if (dossiervide1 == "This folder is empty")
                                    {
                                        // MessageBox.Show("vide");

                                     
                                    }

                                }
                                catch { }
                                // try to find the element 
                                System.Threading.Thread.Sleep(2000);
                                var markasread = driver1.FindElement(By.CssSelector("i[data-icon-name='Read']")).Displayed;
                                // if I get to here the element exists 
                                // if it is visible 
                                System.Threading.Thread.Sleep(2000);

                                if (markasread)
                                {
                                    visible2 = true;
                                    string markasread1 = driver1.FindElement(By.CssSelector("i[data-icon-name='Read']")).Text.ToString();
                                    //  MessageBox.Show(txtblock);

                                    if (markasread1 == "")
                                    {
                                        //MessageBox.Show("comme lu");
                                        // try to read spam and move it
                                        try
                                        {
                                            System.Threading.Thread.Sleep(2000);

                                            var boiteblocked6 = driver1.FindElement(By.ClassName("_2EHjCdO2IEh-zlrH2jOD50")).Displayed;
                                            System.Threading.Thread.Sleep(2000);
                                            //   string fe = driver1.FindElement(By.XPath("//span[contains(@class,'ms-Button-label label-55')][contains(text(),'Marquer tout comme lu')]")).Text.ToString();
                                            //  MessageBox.Show(fe.ToString());
                                            // Ce dossier est vide
                                            do
                                            {
                                                if (boiteblocked6)
                                                {

                                                    // Boolean visible1 = false;  // assume it is invisible 
                                                    /* IWebElement firststep = driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']"));
                                                     firststep.Click();*/
                                                    //*********************code for new version 
                                                    System.Threading.Thread.Sleep(2000);
                                                    driver1.FindElement(By.CssSelector("i[data-icon-name='Blocked']")).Click();
                                                    //driver1.FindElement(By.XPath("//span[contains(@class,'_3lQ0EN5N3oGHxkBKF6_Ane')][contains(text(),'Courrier indésirable')]")).Click();
                                                    System.Threading.Thread.Sleep(2000);
                                                    driver1.FindElement(By.ClassName("_2EHjCdO2IEh-zlrH2jOD50")).Click();
                                                    System.Threading.Thread.Sleep(1500);
                                                    /* driver1.FindElement(By.CssSelector("i[data-icon-name='ChevronDown']")).Click();
                                                     IJavaScriptExecutor js1 = (IJavaScriptExecutor)driver1;

                                                     System.Threading.Thread.Sleep(1500);
                                                    // IWebElement Element1 = driver1.FindElement(By.XPath("//span[contains(@class,'ms-ContextualMenu-itemText label-170')][contains(text(),'Courrier légitime')]"));
                                                     IWebElement Elementbtn = driver1.FindElement(By.XPath("//button[contains(@class,'ms-ContextualMenu-link root-146')][contains(@name,'Courrier légitime')]"));
                                                     MessageBox.Show(Elementbtn.ToString());
                                                       
                                                     System.Threading.Thread.Sleep(15000);
                                                    // IWebElement Element = driver1.FindElement(By.XPath("//span[contains(@class,'ms-ContextualMenu-itemText label-170')][contains(text(),'Not spam')]"));
                                                     //This will scroll the page till the element is found		

                                                     js1.ExecuteScript("arguments[0].scrollIntoView();", Elementbtn);
                                                     Elementbtn.Click();*/
                                                    System.Threading.Thread.Sleep(1500);
                                                    Boolean visibleselctbox = false;
                                                    try
                                                    {
                                                        // var selctbox = driver1.FindElement(By.XPath("//span[contains(@class,'ms-Button-label label-55')][contains(text(),'Not spam')]")).Displayed;

                                                        var courierligitme = driver1.FindElementByLinkText("Courrier légitime").Displayed;
                                                        if (courierligitme)
                                                        {
                                                            visibleselctbox = true;

                                                            driver1.FindElementByLinkText("Courrier légitime").Click();

                                                        }
                                                        /* if(selctbox)
                                                          {
                                                              visibleselctbox = true;

                                                              driver1.FindElement(By.XPath("//span[contains(@class,'ms-Button-label label-55')][contains(text(),'Not spam')]")).Click();
                                                               
                                                            //  MessageBox.Show("here");
                                                    //  driver1.FindElement(By.CssSelector("i[data-icon-name='ChevronDown']")).Click();
                                                      IJavaScriptExecutor js1 = (IJavaScriptExecutor)driver1;

                                                      System.Threading.Thread.Sleep(1500);
                                                     // IWebElement Element1 = driver1.FindElement(By.XPath("//span[contains(@class,'ms-ContextualMenu-itemText label-170')][contains(text(),'Courrier légitime')]"));
                                                         
                                                      System.Threading.Thread.Sleep(15000);
                                                      IWebElement Element = driver1.FindElement(By.XPath("//span[contains(@class,'ms-ContextualMenu-linkContent linkContent-262')][contains(text(),'Not spam')]"));
                                                      //This will scroll the page till the element is found		

                                                      js1.ExecuteScript("arguments[0].scrollIntoView();", Element);
                                                      driver1.FindElement(By.XPath("//span[contains(@class,'ms-ContextualMenu-linkContent linkContent-262')][contains(text(),'Not spam')]")).Click();
                                                        
                                                          }*/

                                                        // MessageBox.Show("over else");
                                                    }
                                                    catch { }
                                                    // new try version anglais 
                                                    try
                                                    {
                                                        var selctbox = driver1.FindElement(By.CssSelector("i[data-icon-name='ChevronDown']")).Displayed;


                                                        if (selctbox)
                                                        {
                                                            visibleselctbox = true;

                                                            //   driver1.FindElement(By.XPath("//span[contains(@class,'ms-Button-label label-55')][contains(text(),'Not spam')]")).Click();


                                                            /*   driver1.FindElement(By.CssSelector("i[data-icon-name='ChevronDown']")).Click();
                                                               IJavaScriptExecutor js1 = (IJavaScriptExecutor)driver1;
                                                               // var rez = driver1.FindElement(By.CssSelector("i[data-icon-name='ChevronDown']")).Location;
                                                               // MessageBox.Show(rez.ToString());
                                                                System.Threading.Thread.Sleep(1500);
                                                               // IWebElement Element1 = driver1.FindElement(By.XPath("//span[contains(@class,'ms-ContextualMenu-itemText label-170')][contains(text(),'Courrier légitime')]"));
                                                         
                                                              //  System.Threading.Thread.Sleep(15000);

                                                                IWebElement Element = driver1.FindElement(By.XPath("//span[contains(@class,'ms-ContextualMenu-itemText label-215')][contains(text(),'Not spam')]"));
                                                                //This will scroll the page till the element is found		
                                                                      // ms-ContextualMenu-item item-160
                                                                string notspm = driver1.FindElement(By.XPath("//span[contains(@class,'ms-ContextualMenu-itemText label-215')][contains(text(),'Not spam')]")).Text.ToString();
                                                                    
                                                                       js1.ExecuteScript("arguments[0].scrollIntoView();",Element);*/
                                                            // MessageBox.Show(notspm);
                                                            // move spam to inbox
                                                            /*    Actions builder = new Actions(driver1);
                                                                IWebElement skype = driver1.FindElement(By.CssSelector("i[data-icon-name='ChevronDown']"));
                                                                builder.MoveToElement(skype, 12, 16).Click().Perform();
                                                                //builder.MoveByOffset(20,20).Click().Perform();
                                                                System.Threading.Thread.Sleep(5000);
                                                                Actions builder1 = new Actions(driver1);
                                                                IJavaScriptExecutor js1 = (IJavaScriptExecutor)driver1;
                                                                // driver1.FindElement(By.XPath("//button[contains(@name,'Not spam')][contains(text(),'Not spam')]"));
                                                                IWebElement skype1 = driver1.FindElement(By.CssSelector("[tabindex='0']"));
                                                                js1.ExecuteScript("arguments[0].scrollIntoView();", skype1);
                                                                    builder.MoveToElement(skype1,160,32).Click().Perform();*/
                                                            //sent with body
                                                            // driver1.FindElementByLinkText("Show blocked content").Click();
                                                            //sent with no body

                                                            driver1.FindElementByLinkText("This isn't spam").Click();
                                                            // Element.Click();
                                                            // driver1.FindElement(By.XPath("//span[contains(@class,'ms-ContextualMenu-linkContent linkContent-262')][contains(text(),'Not spam')]")).Click();
                                                            //  driver1.FindElement(By.Name("Not spam")).Click();
                                                        }

                                                        //  MessageBox.Show("anglais");
                                                    }
                                                    catch { }

                                                    // end try version anglais

                                                    System.Threading.Thread.Sleep(2500);
                                                    visible6 = false;

                                                }
                                            } while (markasread);

                                        }

                                        catch
                                        {

                                        }
                                       
                                    }

                                    //  MessageBox.Show("here");


                                }

                                if (markasread == false)
                                {

                                 
                                }
                            }
                            catch (Exception g)
                            {
                            }

                            // last try
                        }

                        // close this else 
                        /* else
                         {
                             driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']")).Click();
                         }*/

                    }
                    catch
                    {
                    }

                    //  driver1.FindElement(By.XPath("//span[contains(@class,'_3lQ0EN5N3oGHxkBKF6_Ane')][contains(text(),'Courrier indésirable')]")).Click();
                    System.Threading.Thread.Sleep(2000);

                }

                else
                {
                    
                }
            }
            int reportinglistnumber = 0;
            /*  StringBuilder listViewContent = new StringBuilder();
           
           
              foreach (ListViewItem item in this.listView1.Items)
              {
                  listViewContent.Append(item);
                  listViewContent.Append(Environment.NewLine);
                
                  TextWriter tw = new StreamWriter("C:/Users/Public/ReportingResult" + reportinglistnumber+".txt");

                  tw.WriteLine(listViewContent.ToString());
             
                  tw.Close();
              }*/
            string sPath = "C:/Users/Public/Loginsucces" + reportinglistnumber + ".txt";

            System.IO.StreamWriter SaveFile = new System.IO.StreamWriter(sPath);


          
            SaveFile.Close();
          

        }

        private void button29_Click(object sender, EventArgs e)
        {
            ImageList imgs = new ImageList();
            imgs.ImageSize = new Size(50, 50);
            String[] pathss = { };
            pathss = Directory.GetFiles("C:/imgs");


            // buttun test sans imap
            string line;
            int number = 0;
            System.IO.StreamReader file =
    new System.IO.StreamReader(path);

            while ((line = file.ReadLine()) != null)
            {
                string[] lines = line.Split(',');

                string ips = lines[0];
                // MessageBox.Show(lines[0]);
                string boite = lines[1];
                string pass = lines[2];

                number++;
                string[] ipall = ips.Split(':');
                int port = Convert.ToInt32(ipall[1]);
                bool tt = PingHost(ipall[0], port);
                if (tt == true)
                {

                    // MessageBox.Show("tt True");


                    //**************************************

                    var chromeOptions = new ChromeOptions();
                    //Create a new proxy object
                    var proxy = new Proxy();
                    //Set the http proxy value, host and port.
                    // proxy.HttpProxy = ips;
                    //Set the proxy to the Chrome options
                    //  chromeOptions.Proxy = proxy;
                   
                    // disable pop ups 
                    /*DesiredCapabilities capabilities = DesiredCapabilities.Chrome();
                    capabilities.SetCapability(ChromeOptions.Capability, chromeOptions);*/
                    //chromeOptions.AddAdditionalCapability("excludeSwitches", "disable-popup-blocking");
                    //chromeOptions.AddArgument("--disable-popup-blocking");

                    var driver1 = new ChromeDriver(@"C:\webdrivers", chromeOptions);
                    //***********work without proxy******************
                    //var driver1 = new ChromeDriver(@"C:\webdrivers");
                    driver1.Navigate().GoToUrl(BaseUrl);
                    driver1.FindElementByLinkText("Se connecter").Click();

                    var loginBox = driver1.FindElement(By.XPath("//input[contains(text(),'')]"));
                    System.Threading.Thread.Sleep(2500);
                    loginBox.SendKeys(boite);
                    driver1.FindElement(By.XPath("//input[@type='submit'][@value='Suivant']")).Click();
                    System.Threading.Thread.Sleep(2500);
                    try
                    {
                        var suiv = driver1.FindElement(By.XPath("//input[@type='submit'][@value='Suivant']")).Displayed;
                        if (suiv)
                        {
                            /*  listView1.SmallImageList = imgs;
                              listView1.Items.Add("Michael Carrick", 1);*/
                            // listView1.Items.Add(boite + "User Inccorect ! Please Try Again");
                            driver1.Close();
                            driver1.Dispose();
                            continue;
                        }

                    }
                    catch { }

                    var passwordBox = driver1.FindElement(By.XPath("//input[@id='i0118' and contains(text(),'')]"));
                    passwordBox.SendKeys(pass);
                    System.Threading.Thread.Sleep(2500);
                    try
                    {
                        driver1.FindElement(By.XPath("//input[@type='submit'][@value='Se connecter']")).Click();
                        System.Threading.Thread.Sleep(12000);
                        driver1.Manage().Window.Maximize();


                        System.Threading.Thread.Sleep(12000);
                    }
                    catch
                    {
                    }
                    try
                    {
                        // open fisrt boite step
                        //ms-Icon ms-Icon--ChevronRight
                        var classfirstopen = driver1.FindElement(By.CssSelector("i[data-icon-name='ChevronRight']")).Displayed;
                        if (classfirstopen)
                        {
                            System.Threading.Thread.Sleep(3000);
                            //  MessageBox.Show("done fisrt");
                            //1
                            //iconButton nextButton lowerButton
                            IWebElement mv1 = driver1.FindElement(By.CssSelector("i[data-icon-name='ChevronRight']"));
                            Actions st10 = new Actions(driver1);
                            st10.MoveToElement(mv1, 40, 16).Click().Build().Perform();

                            System.Threading.Thread.Sleep(3000);
                            //2
                            //iconButton nextButton lowerButton
                            IWebElement mv2 = driver1.FindElement(By.CssSelector("i[data-icon-name='ChevronRight']"));
                            Actions st1 = new Actions(driver1);
                            st1.MoveToElement(mv2, 40, 16).Click().Build().Perform();
                            System.Threading.Thread.Sleep(3000);
                            //3
                            //ms-Icon ms-Icon--ChevronRight
                            IWebElement mv8 = driver1.FindElement(By.CssSelector("i[data-icon-name='ChevronRight']"));
                            Actions st2 = new Actions(driver1);
                            st2.MoveToElement(mv8, 40, 16).Click().Build().Perform();
                            System.Threading.Thread.Sleep(3000);
                            //4
                            //iconButton nextButton
                            IWebElement mv4 = driver1.FindElement(By.CssSelector("i[data-icon-name='ChevronRight']"));
                            Actions st3 = new Actions(driver1);
                            st3.MoveToElement(mv4, 40, 16).Click().Build().Perform();
                            System.Threading.Thread.Sleep(3000);


                            // btn ok 
                            //
                            driver1.FindElement(By.XPath("//div[contains(@class,'ms-Button-label')][contains(text(),'Got it')]")).Click();
                            System.Threading.Thread.Sleep(3000);


                            System.Threading.Thread.Sleep(2000);
                          
                            driver1.Close();
                            driver1.Dispose();
                            continue;

                        }

                    }
                    catch { }
                    System.Threading.Thread.Sleep(4000);

                    try
                    {
                        string df = driver1.FindElement(By.ClassName("_36rwQx61abZ9wJjdE6uJPX")).Text.ToString();

                        //  MessageBox.Show(df);Sans doute pas
                        if (df == "Not likely")
                        {
                            /* driver1.FindElement(By.ClassName("_3GrIR9uD5dhhMZHdiJ2GC9")).Click();
                             //driver1.FindElement(By.ClassName("ms-Button-icon")).Click();
                             System.Threading.Thread.Sleep(2500);
                             var msgtosent = driver1.FindElement(By.XPath("//textarea[contains(text(),'')]"));
                             System.Threading.Thread.Sleep(2500);
                             msgtosent.SendKeys("good");
                             System.Threading.Thread.Sleep(2500);
                             // driver1.FindElement(By.XPath("//div[contains(@class,'ms-Button-label')][contains(text(),'Envoyer')]")).Click();

                             driver1.FindElement(By.XPath("//div[contains(@class,'ms-Button-label')][contains(text(),'Send')]")).Click();
                             System.Threading.Thread.Sleep(3500);*/

                            driver1.FindElement(By.XPath("//button[contains(@class, 'ms-Button') and contains(@aria-label,'Close')]")).Click();
                            /* IWebElement ele = driver1.FindElement(By.CssSelector("i[data-icon-name='Cancel']"));
                    Actions build = new Actions(driver1);
                      build.MoveToElement(ele,16,16).Click().Build().Perform();*/

                        }
                        if (df == "Sans doute pas")
                        {
                            driver1.FindElement(By.XPath("//button[contains(@class, 'ms-Button') and contains(@aria-label,'Fermer')]")).Click();

                        }

                    }
                    catch { }
                    // driver1.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);
                    try
                    {

                        //
                        //   MessageBox.Show("he 1");
                        System.Threading.Thread.Sleep(2000);
                        IJavaScriptExecutor js1 = (IJavaScriptExecutor)driver1;

                        IWebElement Element14 = driver1.FindElement(By.CssSelector("i[data-icon-name='Cancel']"));
                        var tst = driver1.FindElement(By.CssSelector("i[data-icon-name='Cancel']")).Displayed;
                        System.Threading.Thread.Sleep(2000);
                        //   MessageBox.Show(tst.ToString());
                        js1.ExecuteScript("arguments[0].scrollIntoView();", Element14);
                        Element14.Click();
                        System.Threading.Thread.Sleep(1000);
                        // driver1.FindElement(By.ClassName("_3QWMcW9VDbQi11Tl_MkRml")).Click();



                    }

                    catch { }




                    System.Threading.Thread.Sleep(2000);
                    try
                    {

                        System.Threading.Thread.Sleep(1500);
                        var suiv1 = driver1.FindElement(By.XPath("//input[@id='i0118' and contains(text(),'')]")).Displayed;
                        // string fde = driver1.FindElement(By.ClassName("alert alert-error")).Text.ToString();
                        System.Threading.Thread.Sleep(1000);
                        //   MessageBox.Show(fde);
                        if (suiv1)
                        {
                            // MessageBox.Show("");
                        
                        }

                    }
                    catch { }

                    //Nous avons bien enregistré votre demande
                    Boolean visibledemande = false;

                    try
                    {
                        // try to find the element Nous avons bien enregistré votre demande

                        var boiteblocked4 = driver1.FindElement(By.ClassName("text-title")).Displayed;
                        if (boiteblocked4)
                        {
                            visibledemande = true;
                            string txtblock3 = driver1.FindElement(By.ClassName("text-title")).Text.ToString();
                            if (txtblock3 == "Nous avons bien enregistré votre demande")
                            {

                                //  MessageBox.Show("suivant");
                                System.Threading.Thread.Sleep(1000);
                                driver1.FindElement(By.XPath("//input[@type='button'][@value='Suivant']")).Click();

                                // continue;
                            }
                        }

                    }
                    catch (Exception g)
                    {
                    }
                    // MessageBox.Show("zoom");


                    //try for desactiver 2 type 
                    try
                    {
                        // desactiver prioritaire
                        string typeofboite1 = driver1.FindElement(By.XPath("//div[contains(@class,'ms-Button-label')][contains(text(),'Nouveau courrier')]")).Text.ToString();
                        if (typeofboite1 == "Nouveau courrier")
                        {
                            //  MessageBox.Show("francais");
                            // desactiver priortaire part 2
                            try
                            {
                                // MessageBox.Show(driver1.FindElement(By.XPath("//div[contains(@class,'ms-Button-label')][contains(text(),'Nouveau courrier')]")).Text.ToString());
                                System.Threading.Thread.Sleep(2000);
                                // driver1.FindElement(By.XPath("//div[contains(@class,'ms-Button-label')][contains(text(),'Nouveau courrier')]")).Click();

                                // for clk in option to enable prioritaire for new version outlook
                                IWebElement ell = driver1.FindElement(By.CssSelector("i[data-icon-name='Settings']"));
                                //MessageBox.Show(ell.Location.ToString());
                                ell.Click();

                                /*   IJavaScriptExecutor zoom = (IJavaScriptExecutor)driver1; 
                                    zoom.ExecuteScript("document.body.style.zoom='50%'");*/

                                System.Threading.Thread.Sleep(4000);
                                IWebElement elmasquer = driver1.FindElement(By.XPath("//span[contains(@class,'ms-ChoiceFieldLabel')][contains(text(),'Masquer')]"));
                                IJavaScriptExecutor js2 = (IJavaScriptExecutor)driver1;
                                //This will scroll the page till the element is found	
                                js2.ExecuteScript("arguments[0].scrollIntoView(true);", elmasquer);

                                // js2.ExecuteScript("window.scrollTo(0, document.body.scrollHeight)");
                                //  MessageBox.Show("ok end");
                                // driver1.FindElement(By.XPath("//button[@id='Toggle111']")).Click();
                                System.Threading.Thread.Sleep(2000);
                                driver1.FindElement(By.XPath("//span[contains(@class,'ms-ChoiceFieldLabel')][contains(text(),'Masquer')]")).Click();

                                System.Threading.Thread.Sleep(2000);
                                IJavaScriptExecutor js1 = (IJavaScriptExecutor)driver1;
                                //
                                IWebElement elpr = driver1.FindElement(By.CssSelector("i[data-icon-name='MiniExpandMirrored']"));
                                // IWebElement elpr = driver1.FindElement(By.ClassName("ms-Toggle-thumb thumb-155"));
                                //   IWebElement Element = driver1.FindElement(By.XPath("//button[@id='Toggle111']"));
                                //This will scroll the page till the element is found		


                                js1.ExecuteScript("arguments[0].scrollIntoView(true);", elpr);

                                elpr.Click();
                                //driver1.FindElement(By.XPath("//span[contains(@class,'ms-Label')][contains(text(),'Masquer')]")).Click();
                                System.Threading.Thread.Sleep(2000);
                                IJavaScriptExecutor js3 = (IJavaScriptExecutor)driver1;
                                IWebElement triemsg = driver1.FindElement(By.XPath("//span[contains(@class,'ms-ChoiceFieldLabel')][contains(text(),'Ne pas trier mes messages')]"));
                                js3.ExecuteScript("arguments[0].scrollIntoView(true);", triemsg);
                                System.Threading.Thread.Sleep(3000);
                                triemsg.Click();
                                //
                                driver1.FindElement(By.XPath("//div[contains(@class,'ms-Button-label')][contains(text(),'Enregistrer')]")).Click();
                                System.Threading.Thread.Sleep(3000);

                                // close page of option
                                IWebElement ellclose = driver1.FindElement(By.CssSelector("i[data-icon-name='Cancel']"));
                                //MessageBox.Show(ell.Location.ToString());
                                ellclose.Click();
                            }
                            catch
                            {

                            }
                        }
                    }
                    catch { }

                    // desactiver prioritaire



                    try
                    {
                        string typeofboite = driver1.FindElement(By.XPath("//div[contains(@class,'ms-Button-label')][contains(text(),'New message')]")).Text.ToString();

                        System.Threading.Thread.Sleep(2000);
                        // MessageBox.Show(typeofboite);
                        if (typeofboite == "New message")
                        {

                            //desiactiver part anglais 
                            try
                            {
                                System.Threading.Thread.Sleep(2000);
                                // for clk in option to enable prioritaire for new version outlook
                                IWebElement ell = driver1.FindElement(By.CssSelector("i[data-icon-name='Settings']"));
                                //MessageBox.Show(ell.Location.ToString());
                                ell.Click();
                                System.Threading.Thread.Sleep(2000);

                                IJavaScriptExecutor js2 = (IJavaScriptExecutor)driver1;
                                IWebElement elmasquer = driver1.FindElement(By.XPath("//span[contains(@class,'ms-ChoiceFieldLabel')][contains(text(),'Hide')]"));
                                //   IWebElement Element = driver1.FindElement(By.XPath("//button[@id='Toggle111']"));
                                //This will scroll the page till the element is found		

                                js2.ExecuteScript("arguments[0].scrollIntoView(true);", elmasquer);

                                // driver1.FindElement(By.XPath("//button[@id='Toggle111']")).Click();
                                driver1.FindElement(By.XPath("//span[contains(@class,'ms-ChoiceFieldLabel')][contains(text(),'Hide')]")).Click();

                                System.Threading.Thread.Sleep(2000);

                                // driver1.FindElement(By.ClassName("ms-Toggle-thumb")).Click();

                                IJavaScriptExecutor active = (IJavaScriptExecutor)driver1;
                                IWebElement active1 = driver1.FindElement(By.XPath("//button[contains(@class, 'ms-Toggle-background') and contains(@aria-label,'Focused Inbox')]"));
                                active.ExecuteScript("arguments[0].scrollIntoView(true);", active1);
                                active1.Click();

                                //driver1.FindElement(By.XPath("//div[@class=ms-Toggle-thumb'][0]")).Click();
                                // driver1.FindElement(By.XPath("//div[@class='_28DdWT5RjMfTS_-sWLhDFP']//div[@class='ms-Toggle-thumb']")).Click();
                                System.Threading.Thread.Sleep(2000);


                            }
                            catch
                            {

                            }

                        }
                    }
                    catch { }



                    //



                    // try for older boite
                    Boolean visible1 = false;

                    // try for new boite
                    Boolean visible2 = false;

                    System.Threading.Thread.Sleep(1000);


                    Boolean visible6 = false;
                    System.Threading.Thread.Sleep(2000);
                    System.Threading.Thread.Sleep(3000);
                    driver1.Close();
                    driver1.Dispose();
                }

                else
                {
                   
                }
            }
            int reportinglistnumber = 0;



            string sPath = "C:/Users/Public/ReportingResultPrioritaire" + reportinglistnumber + ".txt";

            System.IO.StreamWriter SaveFile = new System.IO.StreamWriter(sPath);


           
            // done 
        
            MessageBox.Show("Reporting Done!");

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void button30_Click(object sender, EventArgs e)
        {
            ImageList imgs = new ImageList();
            imgs.ImageSize = new Size(50, 50);
            String[] pathss = { };
            pathss = Directory.GetFiles("C:/imgs");


            // buttun test sans imap
            string line;
            int number = 0;
            System.IO.StreamReader file =
    new System.IO.StreamReader(path);

            while ((line = file.ReadLine()) != null)
            {
                string[] lines = line.Split(',');

                string ips = lines[0];
                // MessageBox.Show(lines[0]);
                string boite = lines[1];
                string pass = lines[2];

                number++;
                string[] ipall = ips.Split(':');
                int port = Convert.ToInt32(ipall[1]);
                bool tt = PingHost(ipall[0], port);
                if (tt == true)
                {

                    // MessageBox.Show("tt True");


                    //**************************************

                    var chromeOptions = new ChromeOptions();
                    //Create a new proxy object
                    var proxy = new Proxy();
                    //Set the http proxy value, host and port.
                    // proxy.HttpProxy = ips;
                    //Set the proxy to the Chrome options
                    //  chromeOptions.Proxy = proxy;
                  
                    // disable pop ups 
                    /*DesiredCapabilities capabilities = DesiredCapabilities.Chrome();
                    capabilities.SetCapability(ChromeOptions.Capability, chromeOptions);*/
                    //chromeOptions.AddAdditionalCapability("excludeSwitches", "disable-popup-blocking");
                    //chromeOptions.AddArgument("--disable-popup-blocking");

                    var driver1 = new ChromeDriver(@"C:\webdrivers", chromeOptions);
                    //***********work without proxy******************
                    //var driver1 = new ChromeDriver(@"C:\webdrivers");
                    driver1.Navigate().GoToUrl(BaseUrl);
                    driver1.FindElementByLinkText("Se connecter").Click();

                    var loginBox = driver1.FindElement(By.XPath("//input[contains(text(),'')]"));
                    System.Threading.Thread.Sleep(2500);
                    loginBox.SendKeys(boite);
                    driver1.FindElement(By.XPath("//input[@type='submit'][@value='Suivant']")).Click();
                    System.Threading.Thread.Sleep(2500);
                    try
                    {
                        var suiv = driver1.FindElement(By.XPath("//input[@type='submit'][@value='Suivant']")).Displayed;
                        if (suiv)
                        {
                            /*  listView1.SmallImageList = imgs;
                              listView1.Items.Add("Michael Carrick", 1);*/
                          // listView1.Items.Add(boite + "User Inccorect ! Please Try Again");
                            driver1.Close();
                            driver1.Dispose();
                            continue;
                        }

                    }
                    catch { }

                    var passwordBox = driver1.FindElement(By.XPath("//input[@id='i0118' and contains(text(),'')]"));
                    passwordBox.SendKeys(pass);
                    System.Threading.Thread.Sleep(2500);
                    try
                    {
                        driver1.FindElement(By.XPath("//input[@type='submit'][@value='Se connecter']")).Click();
                        System.Threading.Thread.Sleep(4000);
                        driver1.Manage().Window.Maximize();
                    }
                    catch
                    {
                    }
                    System.Threading.Thread.Sleep(4000);
                    // driver1.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);
                    try
                    {

                        //
                        //   MessageBox.Show("he 1");
                        System.Threading.Thread.Sleep(2000);
                        IJavaScriptExecutor js1 = (IJavaScriptExecutor)driver1;

                        IWebElement Element14 = driver1.FindElement(By.CssSelector("i[data-icon-name='Cancel']"));
                        var tst = driver1.FindElement(By.CssSelector("i[data-icon-name='Cancel']")).Displayed;
                        System.Threading.Thread.Sleep(2000);
                        //   MessageBox.Show(tst.ToString());
                        js1.ExecuteScript("arguments[0].scrollIntoView();", Element14);
                        Element14.Click();
                        System.Threading.Thread.Sleep(1000);
                        // driver1.FindElement(By.ClassName("_3QWMcW9VDbQi11Tl_MkRml")).Click();



                    }

                    catch { }




                    System.Threading.Thread.Sleep(2000);
                    try
                    {

                        System.Threading.Thread.Sleep(1500);
                        var suiv1 = driver1.FindElement(By.XPath("//input[@id='i0118' and contains(text(),'')]")).Displayed;
                        // string fde = driver1.FindElement(By.ClassName("alert alert-error")).Text.ToString();
                        System.Threading.Thread.Sleep(1000);
                        //   MessageBox.Show(fde);
                        if (suiv1)
                        {
                            // MessageBox.Show("");
                          
                        }

                    }
                    catch { }
                    //Nous avons bien enregistré votre demande
                    Boolean visibledemande = false;

                    try
                    {
                        // try to find the element Nous avons bien enregistré votre demande

                        var boiteblocked4 = driver1.FindElement(By.ClassName("text-title")).Displayed;
                        if (boiteblocked4)
                        {
                            visibledemande = true;
                            string txtblock3 = driver1.FindElement(By.ClassName("text-title")).Text.ToString();
                            if (txtblock3 == "Nous avons bien enregistré votre demande")
                            {

                                //  MessageBox.Show("suivant");
                                System.Threading.Thread.Sleep(1000);
                                driver1.FindElement(By.XPath("//input[@type='button'][@value='Suivant']")).Click();

                                // continue;
                            }
                        }

                    }
                    catch (Exception g)
                    {
                    }

                    //try for desactiver 2 type 
                    try
                    {
                        // desactiver prioritaire
                        string typeofboite1 = driver1.FindElement(By.XPath("//div[contains(@class,'ms-Button-label')][contains(text(),'Nouveau courrier')]")).Text.ToString();
                        if (typeofboite1 == "Nouveau courrier")
                        {
                            // desactiver priortaire part 2
                            try
                            {
                                System.Threading.Thread.Sleep(2000);
                                // for clk in option to enable prioritaire for new version outlook
                                IWebElement ell = driver1.FindElement(By.CssSelector("i[data-icon-name='Settings']"));
                                //MessageBox.Show(ell.Location.ToString());
                                ell.Click();
                                System.Threading.Thread.Sleep(2000);
                                IJavaScriptExecutor js2 = (IJavaScriptExecutor)driver1;
                                IWebElement elmasquer = driver1.FindElement(By.XPath("//span[contains(@class,'ms-Label')][contains(text(),'Masquer')]"));
                                //   IWebElement Element = driver1.FindElement(By.XPath("//button[@id='Toggle111']"));
                                //This will scroll the page till the element is found		

                                js2.ExecuteScript("arguments[0].scrollIntoView();", elmasquer);

                                // driver1.FindElement(By.XPath("//button[@id='Toggle111']")).Click();
                                driver1.FindElement(By.XPath("//span[contains(@class,'ms-Label')][contains(text(),'Masquer')]")).Click();

                                System.Threading.Thread.Sleep(2000);
                                IJavaScriptExecutor js1 = (IJavaScriptExecutor)driver1;
                                //
                                IWebElement elpr = driver1.FindElement(By.CssSelector("i[data-icon-name='ChevronRight']"));
                                // IWebElement elpr = driver1.FindElement(By.ClassName("ms-Toggle-thumb thumb-155"));
                                //   IWebElement Element = driver1.FindElement(By.XPath("//button[@id='Toggle111']"));
                                //This will scroll the page till the element is found		


                                js1.ExecuteScript("arguments[0].scrollIntoView();", elpr);

                                elpr.Click();
                                //driver1.FindElement(By.XPath("//span[contains(@class,'ms-Label')][contains(text(),'Masquer')]")).Click();
                                System.Threading.Thread.Sleep(2000);
                                IJavaScriptExecutor js3 = (IJavaScriptExecutor)driver1;
                                IWebElement triemsg = driver1.FindElement(By.XPath("//span[contains(@class,'ms-Label')][contains(text(),'Ne pas trier mes messages')]"));
                                js3.ExecuteScript("arguments[0].scrollIntoView();", triemsg);
                                System.Threading.Thread.Sleep(3000);
                                triemsg.Click();
                                //
                                driver1.FindElement(By.XPath("//div[contains(@class,'ms-Button-label')][contains(text(),'Enregistrer')]")).Click();
                                System.Threading.Thread.Sleep(3000);

                                // close page of option
                                IWebElement ellclose = driver1.FindElement(By.CssSelector("i[data-icon-name='Cancel']"));
                                //MessageBox.Show(ell.Location.ToString());
                                ellclose.Click();
                            }
                            catch
                            {

                            }
                        }
                    }
                    catch { }

                    // desactiver prioritaire



                    try
                    {
                        string typeofboite = driver1.FindElement(By.XPath("//div[contains(@class,'ms-Button-label')][contains(text(),'New message')]")).Text.ToString();

                        System.Threading.Thread.Sleep(2000);
                        // MessageBox.Show(typeofboite);
                        if (typeofboite == "New message")
                        {

                            //desiactiver part anglais 
                            try
                            {
                                System.Threading.Thread.Sleep(2000);
                                // for clk in option to enable prioritaire for new version outlook
                                IWebElement ell = driver1.FindElement(By.CssSelector("i[data-icon-name='Settings']"));
                                //MessageBox.Show(ell.Location.ToString());
                                ell.Click();
                                System.Threading.Thread.Sleep(2000);
                                IJavaScriptExecutor js2 = (IJavaScriptExecutor)driver1;
                                IWebElement elmasquer = driver1.FindElement(By.XPath("//span[contains(@class,'ms-Label')][contains(text(),'Hide')]"));
                                //   IWebElement Element = driver1.FindElement(By.XPath("//button[@id='Toggle111']"));
                                //This will scroll the page till the element is found		

                                js2.ExecuteScript("arguments[0].scrollIntoView();", elmasquer);

                                // driver1.FindElement(By.XPath("//button[@id='Toggle111']")).Click();
                                driver1.FindElement(By.XPath("//span[contains(@class,'ms-Label')][contains(text(),'Hide')]")).Click();

                                System.Threading.Thread.Sleep(2000);
                                IJavaScriptExecutor js1 = (IJavaScriptExecutor)driver1;
                                //
                                IWebElement elpr = driver1.FindElement(By.CssSelector("i[data-icon-name='ChevronRight']"));

                                //This will scroll the page till the element is found	
                                js1.ExecuteScript("arguments[0].scrollIntoView();", elpr);

                                elpr.Click();
                                //driver1.FindElement(By.XPath("//span[contains(@class,'ms-Label')][contains(text(),'Masquer')]")).Click();
                                System.Threading.Thread.Sleep(2000);
                                string sf1 = driver1.FindElement(By.XPath("//span[contains(@class,'_18S85w-gVyogHEMcBxB7gy')][contains(text(),'Settings')]")).Text.ToString();
                                System.Threading.Thread.Sleep(2000);
                                // MessageBox.Show(sf1);
                                driver1.FindElement(By.XPath("//div[contains(@class,'ms-Button-label')][contains(text(),'General')]")).Click();
                                System.Threading.Thread.Sleep(2000);
                                driver1.FindElement(By.XPath("//div[contains(@class,'ms-Button-label')][contains(text(),'Language and time')]")).Click();
                                System.Threading.Thread.Sleep(2000);
                                //string sf2 = driver1.FindElement(By.XPath("//span[contains(@class,'_18S85w-gVyogHEMcBxB7gy')][contains(text(),'Focused Inbox')]")).Text.ToString();
                                System.Threading.Thread.Sleep(2000);
                                // MessageBox.Show(sf2);
                                driver1.FindElement(By.ClassName("ms-Dropdown-container")).FindElement(By.XPath(".//option[contains(text(),'‎français (France)')]")).Click();


                                System.Threading.Thread.Sleep(2000);
                                // IJavaScriptExecutor js3 = (IJavaScriptExecutor)driver1;
                                /*   IWebElement elementsort = driver1.FindElement(By.XPath("//span[contains(@class,'ms-Label')][contains(text(),'Don't sort my messages')]"));
                                   js1.ExecuteScript("arguments[0].scrollIntoView();", elementsort);
                                   System.Threading.Thread.Sleep(3000);
                                   elementsort.Click();*/

                                //
                                driver1.FindElement(By.XPath("//div[contains(@class,'ms-Button-label')][contains(text(),'Save')]")).Click();
                                System.Threading.Thread.Sleep(3000);

                                // close page of option
                                IWebElement ellclose = driver1.FindElement(By.CssSelector("i[data-icon-name='Cancel']"));
                                //MessageBox.Show(ell.Location.ToString());
                                ellclose.Click();
                            }
                            catch
                            {

                            }

                        }
                    }
                    catch { }



                    //



                    // try for older boite
                    Boolean visible1 = false;

                    // try for new boite
                    Boolean visible2 = false;

                    System.Threading.Thread.Sleep(1000);


                    Boolean visible6 = false;
                    System.Threading.Thread.Sleep(2000);
                  System.Threading.Thread.Sleep(2000);
                    driver1.Close();
                    driver1.Dispose();
                }

              
            }
            int reportinglistnumber = 0;



            string sPath = "C:/Users/Public/ReportingResult" + reportinglistnumber + ".txt";
            string sPath1 = "C:/Users/Public/openitagain.txt";

            System.IO.StreamWriter SaveFile = new System.IO.StreamWriter(sPath);

            System.IO.StreamWriter SaveFile1 = new System.IO.StreamWriter(sPath1);
          
            
            // done 
         
            MessageBox.Show("Reporting Done!");
        }

        private void button31_Click(object sender, EventArgs e)
        {
            string sPath = "C:/Users/Public/ReportingResult.txt";
            string sPath1 = "C:/Users/Public/openitagain.txt";
            string sPath2 = "C:/Users/Public/BoiteBlocked.txt";

            string lines = null;
            // string line_to_delete = "Boite Blocked";



            string search_text = "Boite Blocked";
            string[] wordsToDelete = { "Boite Blocked" };

            //  string search_text2 = "Votre compte a été temporairement bloqué";
            string old;
            string n = "";
            StreamReader sr = File.OpenText("C:/Users/Public/ReportingResult.txt");
            int ie = 0;
            while ((old = sr.ReadLine()) != null)
            {

                if (!old.Contains(search_text))
                {
                    n += old + Environment.NewLine;
                }



            }
            sr.Close();
            File.WriteAllText("C:/Users/Public/ReportingResult.txt", n);


            System.Threading.Thread.Sleep(8000);


            string search_text2 = "temporairement";
            string old2;
            string n2 = "";
            StreamReader sr2 = File.OpenText("C:/Users/Public/ReportingResult.txt");

            while ((old2 = sr2.ReadLine()) != null)
            {

                if (!old2.Contains(search_text2))
                {
                    n2 += old2 + Environment.NewLine;
                }

                MessageBox.Show(old2.Contains(search_text2).ToString());

            }
            sr2.Close();
            File.WriteAllText("C:/Users/Public/ReportingResult.txt", n2);


        }

        // start api google test 

        // end api
        static string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
        static string ApplicationName = "Google Sheets API .NET Quickstart";

        private void button32_Click(object sender, EventArgs e)
        {

        }
        public static UserCredential GetCredential()
        {
            // TODO: Change placeholder below to generate authentication credentials. See:
            // https://developers.google.com/sheets/quickstart/dotnet#step_3_set_up_the_sample
            //
            // Authorize using one of the following scopes:
            //     "https://www.googleapis.com/auth/drive"
            //     "https://www.googleapis.com/auth/drive.file"
            //     "https://www.googleapis.com/auth/spreadsheets"
            return null;
        }

        private void richTextBox9_TextChanged(object sender, EventArgs e)
        {

        }

        private void button33_Click(object sender, EventArgs e)
        {
          
        }
        public string[] shoppingList = new string[] { "Item 1", "Item 2", "Item 3", "Item 3", "Item 1", "Item 1", "Item 3", "Item 2" };

        private void button34_Click(object sender, EventArgs e)
        {
           
        }
       public static string path="";
        private void button35_Click(object sender, EventArgs e)
        {
           
            openFileDialog1.FileName = "";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                StreamReader reader = new StreamReader(openFileDialog1.FileName);
           
            }
            MessageBox.Show("Your List Has Been Reloaded");
          
            path = openFileDialog1.FileName;
           // MessageBox.Show(path);
        }

        private void button36_Click(object sender, EventArgs e)
        {
            //save spam
          
                SaveFileDialog save = new SaveFileDialog();

                save.FileName = "";

                save.Filter = "Text File | *.txt";

                if (save.ShowDialog() == DialogResult.OK)

                {

                    StreamWriter writer = new StreamWriter(save.OpenFile());

                   

                    writer.Dispose();

                    writer.Close();

                }
                MessageBox.Show("Save Opened List");
           






            string sPath = "C:/Users/Public/ReportingResult.txt";
            string sPath1 = "C:/Users/Public/openitagain.txt";
            string sPath2 = "C:/Users/Public/BoiteBlocked.txt";

            // string sPath2 = "C:/Users/Public/BoiteBlocked" + reportinglistnumber +DateTime.Now.ToString("yyyyMMdd_hhmmss")+ ".txt";

            System.IO.StreamWriter SaveFile = new System.IO.StreamWriter(sPath);
            System.IO.StreamWriter SaveFile1 = new System.IO.StreamWriter(sPath1);
            System.IO.StreamWriter SaveFile2 = new System.IO.StreamWriter(sPath2);
          
            SaveFile.Close();
          

            SaveFile1.Close();
          

            SaveFile2.Close();

        }

        private void button37_Click(object sender, EventArgs e)
        {
            //save open it again 
            
        
            SaveFileDialog save1 = new SaveFileDialog();

            save1.FileName = "";

            save1.Filter = "Text File | *.txt";

            if (save1.ShowDialog() == DialogResult.OK)

            {

                StreamWriter writer1 = new StreamWriter(save1.OpenFile());

               

                writer1.Dispose();

                writer1.Close();

            }
            MessageBox.Show("List to Open Again Saved");


        }

        private void button38_Click(object sender, EventArgs e)
        {

            //save blocked list :
           
                SaveFileDialog save2 = new SaveFileDialog();

                save2.FileName = "";

                save2.Filter = "Text File | *.txt";

                if (save2.ShowDialog() == DialogResult.OK)

                {

                    StreamWriter writer2 = new StreamWriter(save2.OpenFile());

                    

                    writer2.Dispose();

                    writer2.Close();

                }
                MessageBox.Show("Save blocked List");
          
        }

        private void listBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button39_Click(object sender, EventArgs e)
        {
            label8.Text = "";
            label5.Text = "";
            //proc = 0;
            progressBar1.Value = 0;
            //progressBar1.ResetText();
            if (comboBox1.Text == "" && comboBox2.Text == "")
            {
                MessageBox.Show("Please Enter Your Login and Password");
            }
            else
            {

                using (var client = new MailKit.Net.Imap.ImapClient())
                {
                    // For demo-purposes, accept all SSL certificates
                    client.ServerCertificateValidationCallback = (s, c, h, v) => true;
                    comboBox1.Items.Add(comboBox1.Text);
                    comboBox2.Items.Add(comboBox2.Text);
                   if(comboBox3.Text.Contains("to"))
                    {
                        if (comboBox1.Text.Contains("gmail.com"))
                        {
                            client.Connect("imap.gmail.com", 993, true);

                            client.Authenticate(comboBox1.Text, comboBox2.Text);
                            // MessageBox.Show("here");
                            var test = client.IsAuthenticated;
                            //  MessageBox.Show(test.ToString());
                            var spam = client.GetFolder("[Gmail]/Spam");
                            spam.Open(FolderAccess.ReadOnly);
                            label8.Text = spam.Count.ToString();
                            for (int i = 0; i < spam.Count; i++)
                            {
                                // MessageBox.Show("Start");
                                this.timer1.Start();
                                var messagespam = spam.GetMessage(i);
                                // MessageBox.Show("Subject: {0}", messagespam.Subject);
                                listBox1.Items.Add(messagespam.Subject);
                                richTextBox6.Text += messagespam.To + "\n";
                                //  progressBar1.Minimum = 0;
                                // progressBar1.Maximum = 100;
                                // progressBar1.Step = 1;
                                // ThreadPool.QueueUserWorkItem(new WaitCallback(AnimateScroll));

                                // here's the change:

                            }
                        }
                        if (comboBox1.Text.Contains("hotmail.com"))
                        {
                            client.Connect("imap-mail.outlook.com", 993, true);
                            client.Authenticate(comboBox1.Text, comboBox2.Text);
                            var spam = client.GetFolder("junk");
                            spam.Open(FolderAccess.ReadOnly);
                            label8.Text = spam.Count.ToString();
                            for (int i = 0; i < spam.Count; i++)
                            {
                                // MessageBox.Show("Start");
                                this.timer1.Start();
                                var messagespam = spam.GetMessage(i);
                                // MessageBox.Show("Subject: {0}", messagespam.Subject);
                                listBox1.Items.Add(messagespam.Subject);
                                richTextBox6.Text += messagespam.To + "\n";
                                //  progressBar1.Minimum = 0;
                                // progressBar1.Maximum = 100;
                                // progressBar1.Step = 1;
                                // ThreadPool.QueueUserWorkItem(new WaitCallback(AnimateScroll));

                                // here's the change:

                            }
                        }
                        if (comboBox1.Text.Contains("outlook.com"))
                        {
                            client.Connect("imap-mail.outlook.com", 993, true);
                            client.Authenticate(comboBox1.Text, comboBox2.Text);
                            var spam = client.GetFolder("junk");
                            spam.Open(FolderAccess.ReadOnly);
                            label8.Text = spam.Count.ToString();
                            for (int i = 0; i < spam.Count; i++)
                            {
                                // MessageBox.Show("Start");
                                this.timer1.Start();
                                var messagespam = spam.GetMessage(i);
                                // MessageBox.Show("Subject: {0}", messagespam.Subject);
                                listBox1.Items.Add(messagespam.Subject);
                                richTextBox6.Text += messagespam.To + "\n";
                                //  progressBar1.Minimum = 0;
                                // progressBar1.Maximum = 100;
                                // progressBar1.Step = 1;
                                // ThreadPool.QueueUserWorkItem(new WaitCallback(AnimateScroll));

                                // here's the change:

                            }
                        }
                        if (comboBox1.Text.Contains("yahoo.com"))
                        {
                            client.Connect("imap.mail.yahoo.com", 993, true);

                            client.Authenticate(comboBox1.Text, comboBox2.Text);
                            var spam = client.GetFolder("Bulk Mail");
                            spam.Open(FolderAccess.ReadOnly);
                            label8.Text = spam.Count.ToString();
                            for (int i = 0; i < spam.Count; i++)
                            {
                                // MessageBox.Show("Start");
                                this.timer1.Start();
                                var messagespam = spam.GetMessage(i);
                                // MessageBox.Show("Subject: {0}", messagespam.Subject);
                                listBox1.Items.Add(messagespam.Subject);
                                richTextBox6.Text += messagespam.To + "\n";
                                //  progressBar1.Minimum = 0;
                                // progressBar1.Maximum = 100;
                                // progressBar1.Step = 1;
                                // ThreadPool.QueueUserWorkItem(new WaitCallback(AnimateScroll));

                                // here's the change:

                            }
                        }
                        if (comboBox1.Text.Contains("aol.com"))
                        {
                            client.Connect("imap.aol.com", 993, true);

                            client.Authenticate(comboBox1.Text, comboBox2.Text);
                            var spam = client.GetFolder("Bulk Mail");
                            spam.Open(FolderAccess.ReadOnly);
                            label8.Text = spam.Count.ToString();
                            for (int i = 0; i < spam.Count; i++)
                            {
                                // MessageBox.Show("Start");
                                this.timer1.Start();
                                var messagespam = spam.GetMessage(i);
                                // MessageBox.Show("Subject: {0}", messagespam.Subject);
                                listBox1.Items.Add(messagespam.Subject);
                                richTextBox6.Text += messagespam.To + "\n";
                                //  progressBar1.Minimum = 0;
                                // progressBar1.Maximum = 100;
                                // progressBar1.Step = 1;
                                // ThreadPool.QueueUserWorkItem(new WaitCallback(AnimateScroll));

                                // here's the change:

                            }
                        }
                    }
                   //end geting to
                    if (comboBox3.Text.Contains("subject"))
                    {
                        if (comboBox1.Text.Contains("gmail.com"))
                        {
                            client.Connect("imap.gmail.com", 993, true);

                            client.Authenticate(comboBox1.Text, comboBox2.Text);
                            // MessageBox.Show("here");
                            var test = client.IsAuthenticated;
                            //  MessageBox.Show(test.ToString());
                            var spam = client.GetFolder("[Gmail]/Spam");
                            spam.Open(FolderAccess.ReadOnly);
                            label8.Text = spam.Count.ToString();
                            for (int i = 0; i < spam.Count; i++)
                            {
                                // MessageBox.Show("Start");
                                this.timer1.Start();
                                var messagespam = spam.GetMessage(i);
                                // MessageBox.Show("Subject: {0}", messagespam.Subject);
                                listBox1.Items.Add(messagespam.Subject);
                                richTextBox6.Text += messagespam.To + "\n";
                                //  progressBar1.Minimum = 0;
                                // progressBar1.Maximum = 100;
                                // progressBar1.Step = 1;
                                // ThreadPool.QueueUserWorkItem(new WaitCallback(AnimateScroll));

                                // here's the change:

                            }
                        }
                        if (comboBox1.Text.Contains("hotmail.com"))
                        {
                            client.Connect("imap-mail.outlook.com", 993, true);
                            client.Authenticate(comboBox1.Text, comboBox2.Text);
                            var spam = client.GetFolder("junk");
                            spam.Open(FolderAccess.ReadOnly);
                            label8.Text = spam.Count.ToString();
                            for (int i = 0; i < spam.Count; i++)
                            {
                                // MessageBox.Show("Start");
                                this.timer1.Start();
                                var messagespam = spam.GetMessage(i);
                                // MessageBox.Show("Subject: {0}", messagespam.Subject);
                                listBox1.Items.Add(messagespam.Subject);
                                richTextBox6.Text += messagespam.Subject + "\n";
                                //  progressBar1.Minimum = 0;
                                // progressBar1.Maximum = 100;
                                // progressBar1.Step = 1;
                                // ThreadPool.QueueUserWorkItem(new WaitCallback(AnimateScroll));

                                // here's the change:

                            }
                        }
                        if (comboBox1.Text.Contains("outlook.com"))
                        {
                            client.Connect("imap-mail.outlook.com", 993, true);
                            client.Authenticate(comboBox1.Text, comboBox2.Text);
                            var spam = client.GetFolder("junk");
                            spam.Open(FolderAccess.ReadOnly);
                            label8.Text = spam.Count.ToString();
                            for (int i = 0; i < spam.Count; i++)
                            {
                                // MessageBox.Show("Start");
                                this.timer1.Start();
                                var messagespam = spam.GetMessage(i);
                                // MessageBox.Show("Subject: {0}", messagespam.Subject);
                                listBox1.Items.Add(messagespam.Subject);
                                richTextBox6.Text += messagespam.Subject + "\n";
                                //  progressBar1.Minimum = 0;
                                // progressBar1.Maximum = 100;
                                // progressBar1.Step = 1;
                                // ThreadPool.QueueUserWorkItem(new WaitCallback(AnimateScroll));

                                // here's the change:

                            }
                        }
                        if (comboBox1.Text.Contains("yahoo.com"))
                        {
                            client.Connect("imap.mail.yahoo.com", 993, true);

                            client.Authenticate(comboBox1.Text, comboBox2.Text);
                            var spam = client.GetFolder("Bulk Mail");
                            spam.Open(FolderAccess.ReadOnly);
                            label8.Text = spam.Count.ToString();
                            for (int i = 0; i < spam.Count; i++)
                            {
                                // MessageBox.Show("Start");
                                this.timer1.Start();
                                var messagespam = spam.GetMessage(i);
                                // MessageBox.Show("Subject: {0}", messagespam.Subject);
                                listBox1.Items.Add(messagespam.Subject);
                                richTextBox6.Text += messagespam.Subject + "\n";
                                //  progressBar1.Minimum = 0;
                                // progressBar1.Maximum = 100;
                                // progressBar1.Step = 1;
                                // ThreadPool.QueueUserWorkItem(new WaitCallback(AnimateScroll));

                                // here's the change:

                            }
                        }
                        if (comboBox1.Text.Contains("aol.com"))
                        {
                            client.Connect("imap.aol.com", 993, true);

                            client.Authenticate(comboBox1.Text, comboBox2.Text);
                            var spam = client.GetFolder("Bulk Mail");
                            spam.Open(FolderAccess.ReadOnly);
                            label8.Text = spam.Count.ToString();
                            for (int i = 0; i < spam.Count; i++)
                            {
                                // MessageBox.Show("Start");
                                this.timer1.Start();
                                var messagespam = spam.GetMessage(i);
                                // MessageBox.Show("Subject: {0}", messagespam.Subject);
                                listBox1.Items.Add(messagespam.Subject);
                                richTextBox6.Text += messagespam.Subject + "\n";
                                //  progressBar1.Minimum = 0;
                                // progressBar1.Maximum = 100;
                                // progressBar1.Step = 1;
                                // ThreadPool.QueueUserWorkItem(new WaitCallback(AnimateScroll));

                                // here's the change:

                            }
                        }
                    }
                    //end geting  subject
                    if (comboBox3.Text.Contains("body"))
                    {
                        if (comboBox1.Text.Contains("gmail.com"))
                        {
                            client.Connect("imap.gmail.com", 993, true);

                            client.Authenticate(comboBox1.Text, comboBox2.Text);
                            // MessageBox.Show("here");
                            var test = client.IsAuthenticated;
                            //  MessageBox.Show(test.ToString());
                            var spam = client.GetFolder("[Gmail]/Spam");
                            spam.Open(FolderAccess.ReadOnly);
                            label8.Text = spam.Count.ToString();
                            for (int i = 0; i < spam.Count; i++)
                            {
                                // MessageBox.Show("Start");
                                this.timer1.Start();
                                var messagespam = spam.GetMessage(i);
                                // MessageBox.Show("Subject: {0}", messagespam.Subject);
                                listBox1.Items.Add(messagespam.Subject);
                                richTextBox6.Text += messagespam.HtmlBody + "\n";
                                //  progressBar1.Minimum = 0;
                                // progressBar1.Maximum = 100;
                                // progressBar1.Step = 1;
                                // ThreadPool.QueueUserWorkItem(new WaitCallback(AnimateScroll));

                                // here's the change:

                            }
                        }
                        if (comboBox1.Text.Contains("hotmail.com"))
                        {
                            client.Connect("imap-mail.outlook.com", 993, true);
                            client.Authenticate(comboBox1.Text, comboBox2.Text);
                            var spam = client.GetFolder("junk");
                            spam.Open(FolderAccess.ReadOnly);
                            label8.Text = spam.Count.ToString();
                            for (int i = 0; i < spam.Count; i++)
                            {
                                // MessageBox.Show("Start");
                                this.timer1.Start();
                                var messagespam = spam.GetMessage(i);
                                // MessageBox.Show("Subject: {0}", messagespam.Subject);
                                listBox1.Items.Add(messagespam.Subject);
                                richTextBox6.Text += messagespam.HtmlBody+ "\n";
                                //  progressBar1.Minimum = 0;
                                // progressBar1.Maximum = 100;
                                // progressBar1.Step = 1;
                                // ThreadPool.QueueUserWorkItem(new WaitCallback(AnimateScroll));

                                // here's the change:

                            }
                        }
                        if (comboBox1.Text.Contains("outlook.com"))
                        {
                            client.Connect("imap-mail.outlook.com", 993, true);
                            client.Authenticate(comboBox1.Text, comboBox2.Text);
                            var spam = client.GetFolder("junk");
                            spam.Open(FolderAccess.ReadOnly);
                            label8.Text = spam.Count.ToString();
                            for (int i = 0; i < spam.Count; i++)
                            {
                                // MessageBox.Show("Start");
                                this.timer1.Start();
                                var messagespam = spam.GetMessage(i);
                                // MessageBox.Show("Subject: {0}", messagespam.Subject);
                                listBox1.Items.Add(messagespam.Subject);
                                richTextBox6.Text += messagespam.HtmlBody + "\n";
                                //  progressBar1.Minimum = 0;
                                // progressBar1.Maximum = 100;
                                // progressBar1.Step = 1;
                                // ThreadPool.QueueUserWorkItem(new WaitCallback(AnimateScroll));

                                // here's the change:

                            }
                        }
                        if (comboBox1.Text.Contains("yahoo.com"))
                        {
                            client.Connect("imap.mail.yahoo.com", 993, true);

                            client.Authenticate(comboBox1.Text, comboBox2.Text);
                            var spam = client.GetFolder("Bulk Mail");
                            spam.Open(FolderAccess.ReadOnly);
                            label8.Text = spam.Count.ToString();
                            for (int i = 0; i < spam.Count; i++)
                            {
                                // MessageBox.Show("Start");
                                this.timer1.Start();
                                var messagespam = spam.GetMessage(i);
                                // MessageBox.Show("Subject: {0}", messagespam.Subject);
                                listBox1.Items.Add(messagespam.Subject);
                                richTextBox6.Text += messagespam.HtmlBody + "\n";
                                //  progressBar1.Minimum = 0;
                                // progressBar1.Maximum = 100;
                                // progressBar1.Step = 1;
                                // ThreadPool.QueueUserWorkItem(new WaitCallback(AnimateScroll));

                                // here's the change:

                            }
                        }
                        if (comboBox1.Text.Contains("aol.com"))
                        {
                            client.Connect("imap.aol.com", 993, true);

                            client.Authenticate(comboBox1.Text, comboBox2.Text);
                            var spam = client.GetFolder("Bulk Mail");
                            spam.Open(FolderAccess.ReadOnly);
                            label8.Text = spam.Count.ToString();
                            for (int i = 0; i < spam.Count; i++)
                            {
                                // MessageBox.Show("Start");
                                this.timer1.Start();
                                var messagespam = spam.GetMessage(i);
                                // MessageBox.Show("Subject: {0}", messagespam.Subject);
                                listBox1.Items.Add(messagespam.Subject);
                                richTextBox6.Text += messagespam.HtmlBody + "\n";
                                //  progressBar1.Minimum = 0;
                                // progressBar1.Maximum = 100;
                                // progressBar1.Step = 1;
                                // ThreadPool.QueueUserWorkItem(new WaitCallback(AnimateScroll));

                                // here's the change:

                            }
                        }

                        richTextBox6.Text = string.Join(Environment.NewLine, richTextBox6.Lines.Distinct());

                        //richTextBox6.Lines = linesToAdd.ToArray();
                    }
                    //end geting body 

                    

                    // Get the first personal namespace and list the toplevel folders under it.
                    /* var personal = client.GetFolder(client.PersonalNamespaces[0]);
                     foreach (var folder in personal.GetSubfolders(false))
                         MessageBox.Show("[folder] {0}", folder.Name);*/
                    // MessageBox.Show("start");
                    // The Inbox folder is always available on all IMAP servers...


                    client.Disconnect(true);
                    //MessageBox.Show("Done !");
                    progressBar1.Value = 100;

                }

                string sPath = "C:/Users/Public/passandlog.txt";
                if (File.Exists("C:/Users/Public/passandlog.txt"))
                {
                    System.IO.StreamWriter SaveFile = new System.IO.StreamWriter(sPath);
                    foreach (var item in comboBox1.Items)
                    {
                        SaveFile.WriteLine(item);
                    }
                    foreach (var item in comboBox2.Items)
                    {
                        SaveFile.WriteLine(item);
                    }

                    SaveFile.Close();
                }
                else
                {
                    System.IO.StreamWriter SaveFile = new System.IO.StreamWriter(sPath);
                    foreach (var item in comboBox1.Items)
                    {
                        SaveFile.WriteLine(item);
                    }
                    foreach (var item in comboBox2.Items)
                    {
                        SaveFile.WriteLine(item);
                    }

                    SaveFile.Close();
                }
            }
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void pictureBox7_Click_1(object sender, EventArgs e)
        {
            string dirty = richTextBox6.Text;

            // only space, capital A-Z, lowercase a-z, and digits 0-9 are allowed in the string
            string clean = Regex.Replace(dirty, "[^A-Za-z0-9 ]", " "+ Environment.NewLine);
            richTextBox6.Clear();
            richTextBox6.Text = string.Join(Environment.NewLine, richTextBox6.Lines.Distinct());
            richTextBox6.Text = clean;
            richTextBox6.Text = string.Join(Environment.NewLine, richTextBox6.Lines.Distinct());
            // MessageBox.Show(clean);
        }

        private void tabPage4_Click(object sender, EventArgs e)
        {

        }

        private void richTextBox11_TextChanged(object sender, EventArgs e)
        {

        }

        private void richTextBox10_TextChanged(object sender, EventArgs e)
        {

        }

        private void richTextBox9_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void groupBox6_Enter(object sender, EventArgs e)
        {

        }

        private void button19_Click_1(object sender, EventArgs e)
        {
            // butn mark with color
            string[] Scopes = { SheetsService.Scope.Spreadsheets };
            string ApplicationName = "Google Sheets API .NET Quickstart"; //SheetUpdate update this!

            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal);

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }


            SheetsService sheetsService = new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // The ID of the spreadsheet to update.
            string spreadsheetId = "1-HfInY46gmtxE59-2kSBN_jSBOoaImCytA7LIX1aasM";


            String range = "Your Servers!B2:E";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    sheetsService.Spreadsheets.Values.Get(spreadsheetId, range);

            // Prints the names and majors of students in a sample spreadsheet:
            // https://docs.google.com/spreadsheets/d/1-HfInY46gmtxE59-2kSBN_jSBOoaImCytA7LIX1aasM/edit
            ValueRange response = request.Execute();

            IList<IList<Object>> values = response.Values;


            foreach (string current in richTextBox8.Lines)
            {
                //MessageBox.Show(current);


                int i = 0;
                for (i = 0; i < values.Count; ++i)
                {
                    string tocpmare = values[i][0].ToString();

                    if (tocpmare == current)
                    {
                        richTextBox5.Text += tocpmare + Environment.NewLine;
                        int xe = i;
                        //  MessageBox.Show("in");
                        //coloring spicifc cells
                        Spreadsheet spr = sheetsService.Spreadsheets.Get(spreadsheetId).Execute();
                        Sheet sh = spr.Sheets.Where(s => s.Properties.Title == "Your Servers").FirstOrDefault();
                        int sheetId = (int)sh.Properties.SheetId;

                        //define cell color
                        var userEnteredFormat = new CellFormat()
                        {
                            BackgroundColor = new Data.Color()
                            {
                                Red = 1,
                                Green = 0,
                                Blue = 1,
                                Alpha = 0
                            },
                            TextFormat = new TextFormat()
                            {
                                Bold = true
                            }
                        };
                        BatchUpdateSpreadsheetRequest bussr = new BatchUpdateSpreadsheetRequest();

                        //create the update request for cells from the first row
                        /*
                         *  StartColumnIndex = 0,
                                    StartRowIndex = 0,
                                    EndColumnIndex = 1,
                                    EndRowIndex = 2
                         * */

                        //  int numberspamacolor = richTextBox3.Lines.Count();

                        var updateCellsRequest = new Request()
                        {
                            RepeatCell = new RepeatCellRequest()
                            {
                                Range = new GridRange()
                                {/*
                          SheetId = sheetId,
                          // remarque dima StartColumnIndex ykon sgher men  EndColumnIndex
                       StartColumnIndex = 0, hada chmen colmun
                        StartRowIndex = 0, hada l bdya d color min ay3mer
                        EndColumnIndex = 3, hada ina column aysali
                        EndRowIndex = numberinboxacolor / hada l 3adad d stora li aylwen
                        */
                                    SheetId = sheetId,
                                    StartColumnIndex = 0,
                                    StartRowIndex = i + 1,
                                    EndColumnIndex = 1,
                                    EndRowIndex = i + 2
                                },
                                Cell = new CellData()
                                {
                                    UserEnteredFormat = userEnteredFormat
                                },
                                Fields = "UserEnteredFormat(BackgroundColor,TextFormat)"
                            }
                        };
                        bussr.Requests = new List<Request>();
                        bussr.Requests.Add(updateCellsRequest);
                        var bur = sheetsService.Spreadsheets.BatchUpdate(bussr, spreadsheetId);

                        bur.Execute();
                        // MessageBox.Show("--->" + values[i][0]);
                    }
                    if (tocpmare != current)
                    {
                        //  MessageBox.Show("not exist "+current);
                        // richTextBox5.Text = current;
                    }
                    else
                    {

                    }

                }

            }
            // end of foreach search name of server
            //  richTextBox4.Text = "";
            //richTextBox5.Text = "";
            foreach (var objLI1 in richTextBox8.Lines)
            {

                if (richTextBox5.Lines.Contains(objLI1))
                    richTextBox5.Text += objLI1 + "\n";
                // listBox3.Items.Add(objLI); //Adding Shared List
                else
                    // listBox4.Items.Add(objLI); // Different List

                    richTextBox9.Text += objLI1 + "\n";
            }

            // end test
            if (values != null && values.Count > 0)
            {
                MessageBox.Show("Name, Major");
                foreach (var row in values)
                {


                    // Print columns A and E, which correspond to indices 0 and 4.
                    //   MessageBox.Show("{0}, {1}",row[0].ToString()+" "+row[6].ToString());
                    //  Console.WriteLine("{0}, {1}", row[0], row[4]);
                    //  MessageBox.Show(row[0].ToString());

                    if (row[0].ToString() == "chabba")
                    {


                        //  MessageBox.Show("exist" +  row[2].ToString() );

                        /*   Spreadsheet spr = sheetsService.Spreadsheets.Get(spreadsheetId).Execute();
                           Sheet sh = spr.Sheets.Where(s => s.Properties.Title == "serversuivi").FirstOrDefault();
                           int sheetId = (int)sh.Properties.SheetId;

                           //define cell color
                           var userEnteredFormat = new CellFormat()
                           {
                               BackgroundColor = new Data.Color()
                               {
                                   Red = 1,
                                   Green = 1,
                                   Blue = 0,
                                   Alpha = 1
                               },
                               TextFormat = new TextFormat()
                               {
                                   Bold = true
                               }
                           };
                           BatchUpdateSpreadsheetRequest bussr = new BatchUpdateSpreadsheetRequest();

                           //create the update request for cells from the first row
                           /*
                            *  StartColumnIndex = 0,
                                       StartRowIndex = 0,
                                       EndColumnIndex = 1,
                                       EndRowIndex = 2*/

                        /*    var updateCellsRequest = new Request()
                            {
                                RepeatCell = new RepeatCellRequest()
                                {
                                    Range = new GridRange()
                                    {
                                        SheetId = sheetId,
                                        StartColumnIndex = 4,
                                        StartRowIndex = 0,
                                        EndColumnIndex = 5,
                                        EndRowIndex = 2
                                    },
                                    Cell = new CellData()
                                    {
                                        UserEnteredFormat = userEnteredFormat
                                    },
                                    Fields = "UserEnteredFormat(BackgroundColor,TextFormat)"
                                }
                            };
                            bussr.Requests = new List<Request>();
                            bussr.Requests.Add(updateCellsRequest);
                            var bur = sheetsService.Spreadsheets.BatchUpdate(bussr, spreadsheetId);

                            bur.Execute();*/

                    }

                }
            }
            else
            {
                MessageBox.Show("No data found.");
            }
            Console.Read();


        }
        // method not work just for still remember


    }

    internal class Account
    {
    }
}







