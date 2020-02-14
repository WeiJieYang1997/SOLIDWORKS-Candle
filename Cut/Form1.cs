using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Cut
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }


        public double radius;
        public double depth;
        private void bOK_Click(object sender, EventArgs e)
        {
            double radiusResult;
            double depthResult;

            //Convert strings to doubles and
            //make sure that the text in the two
            //text boxes on the form are positive, 
            //non-zero, numeric values 
            double.TryParse(txtRadius.Text, out radiusResult);
            double.TryParse(txtDepth.Text, out depthResult);
            if (radiusResult > 0 && depthResult > 0)
            {
                radius = radiusResult;
                depth = depthResult;
                //Close the form and continue
                Close();
            }
            else
            {
                //Display a message box telling the user to type
                //positive, non-zero, numeric values in both text boxes
                MessageBox.Show("You must type positive, non-zero, numeric values in both Radius and Depth.");
            }

        }
    }
}
