using System;
using System.Collections.Generic;
using System.ComponentModel;

namespace XMLIndexer
{
    /// <summary>
    /// Represents a material specification from the MaterialTable
    /// </summary>
    public class MaterialSpec
    {
        public int ID { get; set; }
        public string MaterialPartNo { get; set; } = string.Empty;
        public string BysoftMaterialCode { get; set; } = string.Empty;
        public string SolidWorksMaterialCode { get; set; } = string.Empty;
        public double Thickness { get; set; }
        public string Gauge { get; set; } = string.Empty;
        public double ScrapFactor { get; set; }
        public double Pounds { get; set; }
        public string SheetMajor { get; set; } = string.Empty;
        public string SheetMinor { get; set; } = string.Empty;
    }

    /// <summary>
    /// Represents a single line item in a cutlist
    /// </summary>
    public class CutlistItem : INotifyPropertyChanged
    {
        private bool _engCompleted;
        private bool _nstCompleted;
        private bool _lsrCompleted;
        private bool _pchCompleted;
        private bool _frmCompleted;
        private bool _pemCompleted;

        public string PROG { get; set; } = string.Empty;
        public string PartDescription { get; set; } = string.Empty;
        public int QTY { get; set; }
        public int YQTY { get; set; }
        public double XAX { get; set; }
        public double YAX { get; set; }
        public string GA { get; set; } = string.Empty;
        public string Quality { get; set; } = string.Empty;

        // Process tracking checkboxes
        public bool ENGCompleted 
        { 
            get => _engCompleted; 
            set { _engCompleted = value; OnPropertyChanged(nameof(ENGCompleted)); } 
        }
        
        public bool NSTCompleted 
        { 
            get => _nstCompleted; 
            set { _nstCompleted = value; OnPropertyChanged(nameof(NSTCompleted)); } 
        }
        
        public bool LSRCompleted 
        { 
            get => _lsrCompleted; 
            set { _lsrCompleted = value; OnPropertyChanged(nameof(LSRCompleted)); } 
        }
        
        public bool PCHCompleted 
        { 
            get => _pchCompleted; 
            set { _pchCompleted = value; OnPropertyChanged(nameof(PCHCompleted)); } 
        }
        
        public bool FRMCompleted 
        { 
            get => _frmCompleted; 
            set { _frmCompleted = value; OnPropertyChanged(nameof(FRMCompleted)); } 
        }
        
        public bool PEMCompleted 
        { 
            get => _pemCompleted; 
            set { _pemCompleted = value; OnPropertyChanged(nameof(PEMCompleted)); } 
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    /// <summary>
    /// Represents a complete cutlist for a job
    /// </summary>
    public class Cutlist
    {
        public string JobNumber { get; set; } = string.Empty;
        public string GSTNNumber { get; set; } = string.Empty;
        public DateTime GeneratedDate { get; set; } = DateTime.Now;
        public List<CutlistItem> Items { get; set; } = new List<CutlistItem>();
        
        /// <summary>
        /// Formatted title for the cutlist report
        /// </summary>
        public string Title => $"CUT LIST FOR JOB {JobNumber}";
    }
}