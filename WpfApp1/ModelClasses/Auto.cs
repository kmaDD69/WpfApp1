namespace WpfApp1.ModelClasses
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("Auto")]
    public partial class Auto
    {
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int ID { get; set; }

        [StringLength(17)]
        public string VIN { get; set; }

        [StringLength(4)]
        public string SeriaPasport { get; set; }

        [StringLength(6)]
        public string NumbePasport { get; set; }

        [StringLength(150)]
        public string VidanPasport { get; set; }

        [StringLength(150)]
        public string Model { get; set; }

        [StringLength(150)]
        public string TypeV { get; set; }

        [StringLength(150)]
        public string Category { get; set; }

        [StringLength(150)]
        public string RegistrationMark { get; set; }

        [Column(TypeName = "date")]
        public DateTime? YearOfRelease { get; set; }

        [StringLength(10)]
        public string EngineNumber { get; set; }

        [StringLength(150)]
        public string Chassis { get; set; }

        [StringLength(150)]
        public string Bodywork { get; set; }

        [StringLength(30)]
        public string Color { get; set; }

        public int? UserID { get; set; }

        public virtual Users Users { get; set; }
    }
}
