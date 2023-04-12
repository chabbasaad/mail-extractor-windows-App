using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WindowsFormsApplication2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Application.Run(new Form1());
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SKGL.Validate ValidateAKey = new SKGL.Validate();// create an object
            ValidateAKey.secretPhase = "My$ecretPa$$W0rd"; // the passsword
            ValidateAKey.Key = "LMFME-OTQAF-JVBUP-OKFGP"; // enter a valid key
            Console.WriteLine(ValidateAKey.IsValid); // check whether the key has been modified or not
            if (ValidateAKey.IsValid)
            {
                // displaying date
                // remember to use .ToShortDateString after each date!
               MessageBox.Show("This key is created {0}", ValidateAKey.CreationDate.ToShortDateString());
               MessageBox.Show("This key will expire {0}", ValidateAKey.ExpireDate.ToShortDateString());

               MessageBox.Show("This key is set to be valid in {0} day(s)"+ValidateAKey.SetTime);
               MessageBox.Show("This key has {0} day(s) left"+ValidateAKey.DaysLeft);

            }
            else
            {
                // if invalid
                Console.WriteLine("Invalid!");
            }
        }
    }
}
