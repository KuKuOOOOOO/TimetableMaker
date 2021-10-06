using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TimetableMaker.Models
{
    class TimetableModel
    {
        public string _ClassName { get; set; }
        public string _TeacherName { get; set; }
        public DateTime _StartTime { get; set; }
        public DateTime _EndTime { get; set; }
        public bool _isBusy { get; set; }
        public string _BusyText { get; set; }
    }
}
