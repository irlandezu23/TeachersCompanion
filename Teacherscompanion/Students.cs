//------------------------------------------------------------------------------
// <auto-generated>
//    This code was generated from a template.
//
//    Manual changes to this file may cause unexpected behavior in your application.
//    Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace Teacherscompanion
{
    using System;
    using System.Collections.Generic;
    
    public partial class Students
    {
        public Students()
        {
            this.Logs = new HashSet<Logs>();
        }
    
        public int RFid { get; set; }
        public string Firstname { get; set; }
        public string Surname { get; set; }
        public int ClassId { get; set; }
    
        public virtual Classes Class { get; set; }
        public virtual ICollection<Logs> Logs { get; set; }
    }
}
