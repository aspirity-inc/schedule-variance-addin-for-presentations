using System.Drawing;
using System.Windows.Forms;

namespace PowerPointPopUp.Forms
{
    public partial class ProgressForm : Form
    {
        public ProgressForm()
        {
            InitializeComponent();
            ResizeSlideShowProgress();
        }

        public void SetSlideShowProgressContent(string content)
        {
            SlideShowProgress.Text = content;
        }
        public void SetSlideShowProgressForeColor(Color color)
        {
            SlideShowProgress.ForeColor = color;
        }

        private void ProgressForm_Resize(object sender, System.EventArgs e)
        {
            ResizeSlideShowProgress();
        }

        private void ResizeSlideShowProgress()
        {
            SlideShowProgress.Font = new Font("Microsoft Sans Serif", (ClientSize.Height + ClientSize.Width) / 8);
        }
    }
}
