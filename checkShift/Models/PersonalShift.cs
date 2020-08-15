using System.Collections.Generic;

namespace checkShift.Models
{
    class PersonalShift
    {
        public string UserId { get; set; }
        public string UserName { get; set; }
        public List<WorkDay> WorkDays { get; set; }
    }
}
