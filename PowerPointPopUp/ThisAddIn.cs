using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using PowerPointPopUp.Forms;
using PowerPointPopUp.Models;
using PowerPointPopUp.Properties;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PowerPointPopUp
{
    public partial class ThisAddIn
    {
        private ProgressForm _progressForm;
        private SlideshowSession _slideshowSession;
        private readonly Timer _progressTimer = new Timer();
        private int _currentSlideNumber;

        /// <summary>
        /// Initializes application
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Application.SlideShowBegin += ApplicationOnSlideShowBegin;
            Application.SlideShowEnd += ApplicationOnSlideShowEnd;
            Application.SlideShowNextSlide += ApplicationOnSlideShowNextSlide;
        }

        /// <summary>
        /// Initializes slideshow session
        /// </summary>
        /// <param name="wn"></param>
        private void ApplicationOnSlideShowBegin(PowerPoint.SlideShowWindow wn)
        {
            //get lecture time
            var lectureSizeForm = new LectureSizeForm();
            lectureSizeForm.ShowDialog();

            if (lectureSizeForm.DialogResult == DialogResult.OK)
            {
                _progressForm = new ProgressForm();

                //get position of actual screen
                var width = (int)wn.Width;
                var left = (int)wn.Left + width - _progressForm.Width;
                var top = (int)wn.Top;

                //show progress form on actual screen
                _progressForm.Show();
                _progressForm.SetBounds(left, top, _progressForm.Width, _progressForm.Height);

                //prerequisites
                _slideshowSession = new SlideshowSession(lectureSizeForm.TimeForLectureInSeconds, wn.Presentation.Slides.Count);
                _progressTimer.Interval = _slideshowSession.TimerIntervalInMiliSec;
                _progressTimer.Tick += TimerOnTick;
                _progressTimer.Start();
            }
        }

        /// <summary>
        /// Handles move to next slide
        /// </summary>
        /// <param name="wn"></param>
        private void ApplicationOnSlideShowNextSlide(PowerPoint.SlideShowWindow wn)
        {
            _currentSlideNumber = wn.View.Slide.SlideNumber;
        }

        /// <summary>
        /// Handles the end of slideshow
        /// </summary>
        /// <param name="pres"></param>
        private void ApplicationOnSlideShowEnd(PowerPoint.Presentation pres)
        {
            _progressTimer.Stop();
            _progressForm.Close();
        }

        /// <summary>
        /// Handles one timer tick for rewrite current slideshow progress
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="eventArgs"></param>
        private void TimerOnTick(object sender, EventArgs eventArgs)
        {
            long secondsTillStart = (long)(DateTime.UtcNow - _slideshowSession.SlideShowStartTime).TotalSeconds;
            long neededSlideNumber = secondsTillStart / _slideshowSession.TimeForOneSlideInSeconds + 1;

            int progressInPercents = (int)((_currentSlideNumber / (double)neededSlideNumber - 1) * 100);

            _progressForm.SetSlideShowProgressForeColor(_currentSlideNumber >= neededSlideNumber ? Color.Green : Color.Red);
            _progressForm.SetSlideShowProgressContent($"{progressInPercents}%");
        }

        /// <summary>
        /// Finalizes application
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
