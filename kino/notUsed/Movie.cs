using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportsGenerator
{
    public class Movie
    {
        public int ID { get; set; }
        public string Title { get; set; }
        public string Gatunek { get; set; }
        public string Opis { get; set; }
        public string MoviesInfo
        {
            get
            {
                return $"{ ID } { Title }"; 
            }
        }

    }
}
