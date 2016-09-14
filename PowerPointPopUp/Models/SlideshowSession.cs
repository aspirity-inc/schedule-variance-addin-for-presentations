using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointPopUp.Models
{
    public class SlideshowSession
    {
        public DateTime SlideShowStartTime { get; }
        public long TimeForLectureInSeconds { get; }
        public long TimeForOneSlideInSeconds { get; }
        public int TimerIntervalInMiliSec { get; }

        public SlideshowSession(int timeForLectureInMinutes, int slidesCount)
        {
            SlideShowStartTime = DateTime.UtcNow;
            TimeForLectureInSeconds = timeForLectureInMinutes*60;
            TimeForOneSlideInSeconds = TimeForLectureInSeconds/slidesCount;
            TimerIntervalInMiliSec = 1000;
        }
    }
}
