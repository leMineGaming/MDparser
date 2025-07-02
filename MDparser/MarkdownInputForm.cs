using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace MDparser
{
    public partial class MarkdownInputForm : Form
    {
        
        public MarkdownInputForm()
        {
            InitializeComponent();
        }

        public string MarkdownTextInput
        {
            get { return MarkdownText.Text; } // Replace textBox1 with your TextBox's name
        }


        private void MarkdownInputForm_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK; // Set the dialog result to OK when the button is clicked
            Close(); // Close the form
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel; // Set the dialog result to Cancel when the button is clicked
            Close(); // Close the form
        }
    }
}
