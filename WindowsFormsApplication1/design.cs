using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
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
using System.Net.Http;
using System.Configuration;
using System.Web;
/*using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;*/

//using BespokeFusion;

namespace WindowsFormsApplication1
{
    public partial class design : Form
    {
        public design()
        {
            InitializeComponent();
        }

        private void bunifuTextbox1_OnTextChange(object sender, EventArgs e)
        {

        }

        private void bunifuCustomLabel2_Click(object sender, EventArgs e)
        {

        }

        private void bunifuImageButton5_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void bunifuFlatButton3_Click(object sender, EventArgs e)
        {
        
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
                if (bunifuCheckbox2.Checked)
                {
                    try
                    {

                        if (lines[3] == null)
                        {

                        }
                        else
                        {
                            useragent = lines[3];
                        }
                    }
                    catch
                    {
                        MessageBox.Show("Please Add UserAgent to Your List");
                        break;
                    }
                }
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
                    if (bunifuCheckbox2.Checked)
                    {
                        chromeOptions.AddArgument("--user-agent=" + useragent);
                    }
                    if (bunifuCheckbox1.Checked)
                    {

                        chromeOptions.AddArgument("--lang=aus");
                    }
                    else
                    {
                        chromeOptions.AddArguments("--proxy-server=" + ips);
                        chromeOptions.AddArgument("--lang=aus");
                    }
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
                            listBox4.Items.Add(ips + "," + boite + "," + pass + " : User Inccorect ! Please Try Again");
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
                            listBox4.Items.Add(ips + "," + boite + "," + pass + "  : Password Inccorect ! Please Try Again");
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
                            listBox3.Items.Add(ips + "," + boite + "," + pass + " First Use");

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
                            listBox3.Items.Add(ips + "," + boite + "," + pass + " First Use");

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
                            listBox3.Items.Add(ips + "," + boite + "," + pass + " First Use Time zone Updated");
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
                            listBox3.Items.Add(ips + "," + boite + "," + pass + "   : Boite Blocked");
                            System.Threading.Thread.Sleep(1000);
                            driver1.Close();
                            driver1.Dispose();
                            continue;
                        }
                        if (txtblock3 == "Votre compte a été temporairement bloqué")
                        {
                            listBox3.Items.Add(ips + "," + boite + "," + pass + " : Boite Blocked");
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
                                        listBox3.Items.Add(ips + "," + boite + "," + pass + "  : Inbox Folder Empty");
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

                                        listBox3.Items.Add(ips + "," + boite + "," + pass + "  : Spam Empty");
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
                                                    int forcompare = Convert.ToInt32(textBox2.Text);
                                                   
                                                    do
                                                    {

                                                        driver1.FindElement(By.CssSelector("i[data-icon-name='Down']")).Click();
                                                        forcompare--;
                                                        System.Threading.Thread.Sleep(1500);
                                                        // MessageBox.Show(forcompare.ToString());
                                                        System.Threading.Thread.Sleep(1500);
                                                    } while (forcompare != numberofinbox);


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
                                        listBox3.Items.Add(ips + "," + boite + "," + pass + " Done");
                                        driver1.Close();
                                        driver1.Dispose();
                                        /*  driver1.Close();
                                          driver1.Dispose();*/
                                        continue;
                                    }

                                    //  MessageBox.Show("here");


                                }

                                if (markasread == false)
                                {

                                    listBox3.Items.Add(ips + "," + boite + "," + pass + "  :Spam Empty");
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
                    try
                    {
                        listBox4.Items.Add(ips + "," + boite + "," + pass + "   :Try to Open it Again");
                        driver1.Close();
                        driver1.Dispose();
                    }
                    catch { }

                    //  driver1.FindElement(By.XPath("//span[contains(@class,'_3lQ0EN5N3oGHxkBKF6_Ane')][contains(text(),'Courrier indésirable')]")).Click();
                    System.Threading.Thread.Sleep(2000);

                }

                else
                {
                    listBox4.Items.Add(ips + "," + boite + "," + pass + "   :Proxy dosen't work");
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
            string sPath = "C:/Users/Public/ReportingResultInbox" + reportinglistnumber + ".txt";

            System.IO.StreamWriter SaveFile = new System.IO.StreamWriter(sPath);


            foreach (var item in listBox3.Items)
            {
                SaveFile.WriteLine(item.ToString());
            }
            SaveFile.Close();

            /*listBox2.Text = "";
            listBox3.Text = "";*/

            // done 
            bunifuProgressBar1.Value = 100;
            int SerialsCounter = listBox3.Items.Count;
            label2.Text = SerialsCounter.ToString();
            MessageBox.Show("Reporting Done!");
        }

        private void bunifuImageButton6_Click(object sender, EventArgs e)
        {
          
            this.WindowState = FormWindowState.Minimized;
          //  this.Hide();
        }

        public static IPAddress GetIPAddress(string hostName)
        {
            Ping ping = new Ping();
            var replay = ping.Send(hostName);

            if (replay.Status == IPStatus.Success)
            {
                return replay.Address;
            }
            return null;
        }
        private static string GetMachineNameFromIPAddress(string ipAdress)
        {
            string machineName = string.Empty;
            try
            {
                System.Net.IPHostEntry hostEntry = System.Net.Dns.GetHostEntry(ipAdress);
                machineName = hostEntry.HostName;
            }
            catch (Exception ex)
            {
                //log here
            }
            return machineName;
        }
        // this return ip if you give it hostname
        private static string GetIPAddressFromMachineName(string machinename)
        {
            string ipAddress = string.Empty;
            try
            {
                System.Net.IPAddress ip = System.Net.Dns.GetHostEntry(machinename).AddressList.Where(o => o.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork).First();
                ipAddress = ip.ToString();
            }
            catch (Exception ex)
            {
                //log here
            }
            return ipAddress;
        }

        public static string BaseUrl2 = "https://www.iplocation.net/find-ip-address";
        private void design_Load(object sender, EventArgs e)
        {

           /* try
            {

            //   MessageBox.Show(HostName);
            var chromeOptions = new ChromeOptions();
            //Create a new proxy object
            var proxy = new Proxy();
            //Set the http proxy value, host and port.
            // proxy.HttpProxy = ips;
            //Set the proxy to the Chrome options
            //  chromeOptions.Proxy = proxy;
          
            chromeOptions.AddArguments("--disable-notifications");
             chromeOptions.AddArgument("--headless");
               chromeOptions.AddArgument("--lang=aus");

             //  chromeOptions.AddArguments("--proxy-server=" + "104.223.54.3:3838");
                var driver1 = new ChromeDriver(@"C:\webdrivers", chromeOptions);
            //***********work without proxy******************
            //var driver1 = new ChromeDriver(@"C:\webdrivers");
            driver1.Navigate().GoToUrl(BaseUrl2);
            string iptocompare = driver1.FindElement(By.XPath("//div[contains(@class, 'col col_6_of_12') and contains(@style,'padding-left: 0px;')]//span")).Text.ToString();
           //MessageBox.Show(iptocompare);
             driver1.Close();
             driver1.Dispose();

                string[] array1 = { "81.192.142.176", "81.192.142.89", "81.192.142.92", "173.212.201.187", "81.192.142.93" };
                foreach (string lang in array1)
                {
                  
                 if(iptocompare == lang)
                    {
                        // MessageBox.Show("egale"+lang);
                        break;
                    }
                  if(iptocompare!= lang)
                    {
                        MessageBox.Show("You Dont Have Access To this App ! Please Contact The Owner");
                        System.Windows.Forms.Application.Exit();
                        driver1.Close();
                        driver1.Dispose();
                    }
                 
                }
                }

            catch
            {
                MessageBox.Show("You Dont Have Access To this App ! Please Contact The Owner");
                System.Windows.Forms.Application.Exit();
               
            }*/

            bunifuCheckbox1.Checked = false;
            bunifuCheckbox2.Checked = false;


        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            int SerialsCounter = listBox2.Items.Count;
            label1.Text = SerialsCounter.ToString();
        }
        public static string path = "";
        private void bunifuImageButton1_Click(object sender, EventArgs e)
        {
            listBox2.Items.Clear();
            openFileDialog1.FileName = "";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                StreamReader reader = new StreamReader(openFileDialog1.FileName);
                string line;
                while ((line = reader.ReadLine()) != null)
                    listBox2.Items.Add(line);
                reader.Close();
            }
            MessageBox.Show("Your List Has Been Reloaded");
            if (this.listBox2.Items.Count > 0)
                this.listBox2.SetSelected(0, true);

            path = openFileDialog1.FileName;
        }

        private void bunifuImageButton2_Click(object sender, EventArgs e)
        {
            //save open it again 


            SaveFileDialog save1 = new SaveFileDialog();

            save1.FileName = "";

            save1.Filter = "Text File | *.txt";

            if (save1.ShowDialog() == DialogResult.OK)

            {

                StreamWriter writer1 = new StreamWriter(save1.OpenFile());

                for (int i = 0; i < listBox4.Items.Count; i++)

                {

                    writer1.WriteLine(listBox4.Items[i].ToString());

                }

                writer1.Dispose();

                writer1.Close();

            }
            MessageBox.Show("List to Open Again Saved");
        }

        private void bunifuImageButton3_Click(object sender, EventArgs e)
        {
            //save blocked list :

            SaveFileDialog save2 = new SaveFileDialog();

            save2.FileName = "";

            save2.Filter = "Text File | *.txt";

            if (save2.ShowDialog() == DialogResult.OK)

            {

                StreamWriter writer2 = new StreamWriter(save2.OpenFile());

                for (int i = 0; i < listBox3.Items.Count; i++)

                {
                    if (listBox3.Items[i].ToString().Contains("BoiteBlocked"))
                    {

                        writer2.WriteLine(listBox3.Items[i].ToString());
                    }
                }

                writer2.Dispose();

                writer2.Close();

            }
            MessageBox.Show("Save blocked List");
        }

        private void bunifuImageButton4_Click(object sender, EventArgs e)
        {
            //save spam

            SaveFileDialog save = new SaveFileDialog();

            save.FileName = "";

            save.Filter = "Text File | *.txt";

            if (save.ShowDialog() == DialogResult.OK)

            {

                StreamWriter writer = new StreamWriter(save.OpenFile());

                for (int i = 0; i < listBox3.Items.Count; i++)

                {
                    if (listBox3.Items[i].ToString().Contains("Spam Empty") || listBox3.Items[i].ToString().Contains("Done"))
                    {

                        writer.WriteLine(listBox2.Items[i].ToString());
                    }
                }

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
            foreach (var item in listBox3.Items)
            {
                if (item.ToString().Contains("Spam Empty"))
                {
                    SaveFile.WriteLine(item.ToString());
                }
                if (item.ToString().Contains("Done"))
                {
                    SaveFile.WriteLine(item.ToString());
                }
            }
            SaveFile.Close();
            foreach (var item in listBox4.Items)
            {
                SaveFile1.WriteLine(item.ToString());
            }

            SaveFile1.Close();
            foreach (var item in listBox3.Items)
            {
                if (item.ToString().Contains("BoiteBlocked"))
                {
                    SaveFile2.WriteLine(item.ToString());
                }
            }

            SaveFile2.Close();
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
        public static string useragent = "";
        public string versionofmail = "";

        public static string BaseUrl = "https://outlook.live.com/owa/";
        private void bunifuFlatButton1_Click(object sender, EventArgs e)
        {


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
                if (bunifuCheckbox2.Checked)
                {
                    try { 

                    if(lines[3] == null)
                    {
                      
                    }
                    else {
                    useragent = lines[3];
                    }
                    }
                    catch
                    {
                        MessageBox.Show("Please Add UserAgent to Your List");
                        break;
                    }
                }
              
            
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
                    if(bunifuCheckbox2.Checked)
                    {
                        chromeOptions.AddArgument("--user-agent=" + useragent);
                    }
                   
                    //
                    if (bunifuCheckbox1.Checked)
                    {

                        chromeOptions.AddArgument("--lang=aus");
                    }
                    else
                    {
                        chromeOptions.AddArguments("--proxy-server=" + ips);
                        chromeOptions.AddArgument("--lang=aus");
                    }
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
                            listBox4.Items.Add(ips + "," + boite + "," + pass + " : User Inccorect ! Please Try Again");
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
                            listBox4.Items.Add(ips + "," + boite + "," + pass + "  : Password Inccorect ! Please Try Again");
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
                            listBox3.Items.Add(ips + "," + boite + "," + pass + " First Use");

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
                            listBox3.Items.Add(ips + "," + boite + "," + pass + " First Use");

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
                            listBox3.Items.Add(ips + "," + boite + "," + pass + " First Use Time zone Updated");
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
                            listBox3.Items.Add(ips + "," + boite + "," + pass + "   : Boite Blocked");
                            System.Threading.Thread.Sleep(1000);
                            driver1.Close();
                            driver1.Dispose();
                            continue;
                        }
                        if (txtblock3 == "Votre compte a été temporairement bloqué")
                        {
                            listBox3.Items.Add(ips + "," + boite + "," + pass + " : Boite Blocked");
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
                                        listBox3.Items.Add(ips + "," + boite + "," + pass + "   : Spam Empty");
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

                                        listBox3.Items.Add(ips + "," + boite + "," + pass + "  : Spam Empty");
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
                                        listBox3.Items.Add(ips + "," + boite + "," + pass + " Done");
                                        driver1.Close();
                                        driver1.Dispose();
                                        continue;
                                    }

                                    //  MessageBox.Show("here");


                                }

                                if (markasread == false)
                                {

                                    listBox3.Items.Add(ips + "," + boite + "," + pass + "  :Spam Empty");
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
                        listBox4.Items.Add(ips + "," + boite + "," + pass + "   :Try to Open it Again");
                        driver1.Close();
                        driver1.Dispose();
                    }
                    catch { }
                }

                else
                {
                    listBox4.Items.Add(ips + "," + boite + "," + pass + "   :Proxy dosen't work");
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
            foreach (var item in listBox3.Items)
            {
                SaveFile.WriteLine(item.ToString());
            }
            SaveFile.Close();
            foreach (var item in listBox4.Items)
            {
                SaveFile1.WriteLine(item.ToString());
            }

            SaveFile1.Close();


            // done 
            bunifuProgressBar1.Value = 100;
            int SerialsCounter = listBox3.Items.Count;
            label2.Text = SerialsCounter.ToString();
            MessageBox.Show("Reporting Done!");
        }

        private void bunifuFlatButton2_Click(object sender, EventArgs e)
        {
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
                if (bunifuCheckbox2.Checked)
                {
                    try
                    {

                        if (lines[3] == null)
                        {

                        }
                        else
                        {
                            useragent = lines[3];
                        }
                    }
                    catch
                    {
                        MessageBox.Show("Please Add UserAgent to Your List");
                        break;
                    }
                }
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
                    if (bunifuCheckbox2.Checked)
                    {
                        chromeOptions.AddArgument("--user-agent=" + useragent);
                    }
                    if (bunifuCheckbox1.Checked)
                    {

                        chromeOptions.AddArgument("--lang=aus");
                    }
                    else
                    {
                        chromeOptions.AddArguments("--proxy-server=" + ips);
                        chromeOptions.AddArgument("--lang=aus");
                    }
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
                            listBox4.Items.Add(ips + "," + boite + "," + pass + " : User Inccorect ! Please Try Again");
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
                        System.Threading.Thread.Sleep(11000);
                        driver1.Manage().Window.Maximize();


                        System.Threading.Thread.Sleep(10000);
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
                            listBox3.Items.Add(ips + "," + boite + "," + pass + " First Use");

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
                            listBox4.Items.Add(ips + "," + boite + "," + pass + "  : Password Inccorect ! Please Try Again");
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
                    listBox3.Items.Add(ips + "," + boite + "," + pass + "   :Volet de lecture Masquer et Boîte de réception Prioritaire Désactiver");
                    System.Threading.Thread.Sleep(3000);
                    driver1.Close();
                    driver1.Dispose();
                }

                else
                {
                    listBox4.Items.Add(ips + "," + boite + "," + pass + "  :Proxy dosen't work");

                }
            }
            int reportinglistnumber = 0;



            string sPath = "C:/Users/Public/ReportingResultPrioritaire" + reportinglistnumber + ".txt";

            System.IO.StreamWriter SaveFile = new System.IO.StreamWriter(sPath);


            foreach (var item in listBox3.Items)
            {
                SaveFile.WriteLine(item.ToString());
            }
            SaveFile.Close();
            // done 
            bunifuProgressBar1.Value = 100;
            int SerialsCounter = listBox3.Items.Count;
            label2.Text = SerialsCounter.ToString();
            MessageBox.Show("Reporting Done!");
        }

        private void bunifuImageButton13_Click(object sender, EventArgs e)
        {
           
        }

        private void bunifuFlatButton4_Click(object sender, EventArgs e)
        {
           // MaterialMessageBox.Show("Your cool message here", "The awesome message title");
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
                if (bunifuCheckbox2.Checked)
                {
                    try
                    {

                        if (lines[3] == null)
                        {

                        }
                        else
                        {
                            useragent = lines[3];
                        }
                    }
                    catch
                    {
                        MessageBox.Show("Please Add UserAgent to Your List");
                        break;
                    }
                }
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
                    if (bunifuCheckbox2.Checked)
                    {
                        chromeOptions.AddArgument("--user-agent=" + useragent);
                    }
                    if (bunifuCheckbox1.Checked)
                    {

                        chromeOptions.AddArgument("--lang=aus");
                    }
                    else
                    {
                        chromeOptions.AddArguments("--proxy-server=" + ips);
                        chromeOptions.AddArgument("--lang=aus");
                    }
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
                            listBox4.Items.Add(ips + "," + boite + "," + pass + " : User Inccorect ! Please Try Again");
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
                            listBox4.Items.Add(ips + "," + boite + "," + pass + "  : Password Inccorect ! Please Try Again");
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
                            listBox3.Items.Add(ips + "," + boite + "," + pass + " First Use");

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
                            listBox3.Items.Add(ips + "," + boite + "," + pass + " First Use");

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
                            listBox3.Items.Add(ips + "," + boite + "," + pass + " First Use Time zone Updated");
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
                            listBox3.Items.Add(ips + "," + boite + "," + pass + "   : Boite Blocked");
                            System.Threading.Thread.Sleep(1000);
                            driver1.Close();
                            driver1.Dispose();
                            continue;
                        }
                        if (txtblock3 == "Votre compte a été temporairement bloqué")
                        {
                            listBox3.Items.Add(ips + "," + boite + "," + pass + " : Boite Blocked");
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

                      //  driver1.FindElement(By.CssSelector("i[data-icon-name='Inbox']")).Click();
                        //driver1.FindElement(By.XPath("//span[contains(@class,'_3lQ0EN5N3oGHxkBKF6_Ane')][contains(text(),'Courrier indésirable')]")).Click();
                        var inboxshow = driver1.FindElement(By.CssSelector("i[data-icon-name='Inbox']")).Displayed;
                        System.Threading.Thread.Sleep(1000);
                        // MessageBox.Show("mark now");
                        //  var spamshow = driver1.FindElement(By.XPath("//span[contains(@class,'_3lQ0EN5N3oGHxkBKF6_Ane')][contains(text(),'Courrier indésirable')]")).Displayed;
                        try
                        {
                            //driver1.FindElement(By.XPath("//button[contains(@class, 'ms-Button') and contains(@name,'Mark all as read')]")).Click();

                            // var partfrench = driver1.FindElement(By.XPath("//div[contains(@class,'ms-Button-label')][contains(text(),'Marquer tout comme lu')]")).Displayed;

                         //   MessageBox.Show("catch");
                            driver1.FindElement(By.XPath("//button[contains(@class, 'ms-Button') and contains(@name,'Mark all as read')]")).Click();
                        }
                        catch { }
                        try
                        {
                            //driver1.FindElement(By.XPath("//button[contains(@class, 'ms-Button') and contains(@name,'Mark all as read')]")).Click();

                            // var partfrench = driver1.FindElement(By.XPath("//div[contains(@class,'ms-Button-label')][contains(text(),'Marquer tout comme lu')]")).Displayed;

                            System.Threading.Thread.Sleep(2000);
                            driver1.FindElement(By.XPath("//div[contains(@class,'ms-Button-label')][contains(text(),'OK')]")).Click();
                        }
                        catch { }
                       

                        try
                        {
                            driver1.FindElement(By.XPath("//div[contains(@class,'ms-Button-label')][contains(text(),'Marquer tout comme lu')]")).Click();

                        }
                        catch { }
                        try
                        {
                            //driver1.FindElement(By.XPath("//button[contains(@class, 'ms-Button') and contains(@name,'Mark all as read')]")).Click();

                            // var partfrench = driver1.FindElement(By.XPath("//div[contains(@class,'ms-Button-label')][contains(text(),'Marquer tout comme lu')]")).Displayed;

                            System.Threading.Thread.Sleep(2000);
                            driver1.FindElement(By.XPath("//div[contains(@class,'ms-Button-label')][contains(text(),'OK')]")).Click();
                        }
                        catch { }

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
                                        listBox3.Items.Add(ips + "," + boite + "," + pass + "  : Inbox Folder Empty");
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

                                        listBox3.Items.Add(ips + "," + boite + "," + pass + "  : Spam Empty");
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
                                     /*   try
                                        {
                                            System.Threading.Thread.Sleep(2000);

                                            var boiteblocked6 = driver1.FindElement(By.ClassName("_2EHjCdO2IEh-zlrH2jOD50")).Displayed;
                                            System.Threading.Thread.Sleep(2000);*/
                                            //   string fe = driver1.FindElement(By.XPath("//span[contains(@class,'ms-Button-label label-55')][contains(text(),'Marquer tout comme lu')]")).Text.ToString();
                                            //  MessageBox.Show(fe.ToString());
                                            // Ce dossier est vide
                                        /*    do
                                            {

                                                if (boiteblocked6)
                                                {
                                                    //  MessageBox.Show("inbox dispaled");
                                                    // Boolean visible1 = false;  // assume it is invisible 
                                                    /* IWebElement firststep = driver1.FindElement(By.CssSelector("i[data-icon-name='GlobalNavButton']"));
                                                     firststep.Click();*/
                                                    //*********************code for new version 
                                                 //   System.Threading.Thread.Sleep(2000);
                                                    //  driver1.FindElement(By.CssSelector("i[data-icon-name='Blocked']")).Click();
                                                    //driver1.FindElement(By.XPath("//span[contains(@class,'_3lQ0EN5N3oGHxkBKF6_Ane')][contains(text(),'Courrier indésirable')]")).Click();
                                                //    System.Threading.Thread.Sleep(2000);
                                                    // for click in email in inbox

                                                   // driver1.FindElement(By.ClassName("_2EHjCdO2IEh-zlrH2jOD50")).Click();

                                                   // System.Threading.Thread.Sleep(1500);
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
                                                   // int numberofinbox = 0;
                                                  //  int forcompare = 0;

                                                  /*  do
                                                    {
                                                        // code pour marquer tout comme lu
                                                       // driver1.FindElement(By.XPath("//div[contains(@class,'ms-Button-label')][contains(text(),'Marquer tout comme lu')]")).Click();
                                                      //  driver1.FindElement(By.CssSelector("i[data-icon-name='Down']")).Click();
                                                        forcompare--;
                                                        // chabba
                                                       // forcompare++;
                                                        System.Threading.Thread.Sleep(1500);
                                                        // MessageBox.Show(forcompare.ToString());
                                                        System.Threading.Thread.Sleep(1500);
                                                    } while (forcompare != numberofinbox);

                                                    */

                                                 //   System.Threading.Thread.Sleep(1500);

                                                    // new try version anglais 


                                                    // end try version anglais

                                                  //  System.Threading.Thread.Sleep(2500);
                                             /*       visible6 = false;

                                                }

                                            } while (markasread);*/

                                     /*   }

                                        catch
                                        {

                                        }*/

                                        listBox3.Items.Add(ips + "," + boite + "," + pass + " Done");
                                        driver1.Close();
                                        driver1.Dispose();
                                        /*  driver1.Close();
                                          driver1.Dispose();*/
                                        continue;
                                    }

                                    //  MessageBox.Show("here");


                                }

                                if (markasread == false)
                                {

                                    listBox3.Items.Add(ips + "," + boite + "," + pass + "  :Spam Empty");
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
                    try
                    {
                        listBox4.Items.Add(ips + "," + boite + "," + pass + "   :Try to Open it Again");
                        driver1.Close();
                        driver1.Dispose();
                    }
                    catch { }

                    //  driver1.FindElement(By.XPath("//span[contains(@class,'_3lQ0EN5N3oGHxkBKF6_Ane')][contains(text(),'Courrier indésirable')]")).Click();
                    System.Threading.Thread.Sleep(2000);

                }

                else
                {
                    listBox4.Items.Add(ips + "," + boite + "," + pass + "   :Proxy dosen't work");
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
            string sPath = "C:/Users/Public/ReportingResultInbox" + reportinglistnumber + ".txt";

            System.IO.StreamWriter SaveFile = new System.IO.StreamWriter(sPath);


            foreach (var item in listBox3.Items)
            {
                SaveFile.WriteLine(item.ToString());
            }
            SaveFile.Close();

            /*listBox2.Text = "";
            listBox3.Text = "";*/

            // done 
            bunifuProgressBar1.Value = 100;
            int SerialsCounter = listBox3.Items.Count;
            label2.Text = SerialsCounter.ToString();
            MessageBox.Show("Reporting Done!");
        }

        private void bunifuCheckbox1_OnChange(object sender, EventArgs e)
        {

        }

        private void bunifuImageButton15_Click(object sender, EventArgs e)
        {
           
            DialogResult dr = MessageBox.Show("Are You sure Do You want to Clear Result ?", "Clear Information !", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

            if (dr == DialogResult.Yes)
            {
                listBox3.Items.Clear();
                listBox4.Items.Clear();
            }
            else if (dr == DialogResult.Cancel)
            {
                //
            }
           
          

        }

        private void bunifuGradientPanel1_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
