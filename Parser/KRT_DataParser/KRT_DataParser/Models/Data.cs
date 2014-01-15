using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace KRT_DataParser.Models
{
    public class Data
    {
        public int Year { get; set; }
        public int Month { get; set; }
        public List<Station> Stations { get; set; }
    }
}