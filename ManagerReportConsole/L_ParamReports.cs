//------------------------------------------------------------------------------
// <auto-generated>
//    Этот код был создан из шаблона.
//
//    Изменения, вносимые в этот файл вручную, могут привести к непредвиденной работе приложения.
//    Изменения, вносимые в этот файл вручную, будут перезаписаны при повторном создании кода.
// </auto-generated>
//------------------------------------------------------------------------------

namespace ManagerReportConsole
{
    using System;
    using System.Collections.Generic;
    
    public partial class L_ParamReports
    {
        public int id { get; set; }
        public string name { get; set; }
        public int id_type { get; set; }
        public string value_fix { get; set; }
        public Nullable<int> id_value_exp { get; set; }
        public Nullable<int> id_value_directory { get; set; }
        public Nullable<int> id_value_format { get; set; }
        public int id_filereport { get; set; }
        public string email { get; set; }
    
        public virtual L_FilesReport L_FilesReport { get; set; }
        public virtual S_ParamDirectory S_ParamDirectoryValueDirectory { get; set; }
        public virtual S_ParamSys S_ParamSysType { get; set; }
        public virtual S_ParamSys S_ParamSysValueExp { get; set; }
        public virtual S_ParamSys S_ParamSysValueFormat { get; set; }
    }
}
