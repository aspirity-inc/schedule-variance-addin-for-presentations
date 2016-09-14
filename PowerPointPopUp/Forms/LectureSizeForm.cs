using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PowerPointPopUp.Properties;

namespace PowerPointPopUp.Forms
{
    public partial class LectureSizeForm : Form
    {
        public int TimeForLectureInSeconds { get; set; }

        public LectureSizeForm()
        {
            InitializeComponent();
        }

        private void send_Click(object sender, EventArgs e)
        {
            int parsedValue;
            if (int.TryParse(textBox.Text, out parsedValue) && parsedValue > 0 && parsedValue < int.Parse(Settings.Default.MaxLectureTimeInMinutes))
            {
                TimeForLectureInSeconds = parsedValue;
                DialogResult = DialogResult.OK;
                Close();
            }
            else
            {
                MessageBox.Show(string.Format(Resources.LectureSizeFormValidationMessage, 0, Settings.Default.MaxLectureTimeInMinutes));
            }
        }
    }
}
