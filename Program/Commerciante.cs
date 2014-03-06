using System;

namespace ExtractFromPagineGialle
{
    class Commerciante
    {
        public String Nome { get; set; }
        public String Citta { get; set; }
        public String Cap { get; set; }
        public String Indirizzo { get; set; }
        public String Contatti { get; set; }
        public String Link { get; set; }
        public String LinkPagineGialle { get; set; }
        public String Descrizione { get; set; }

        public Commerciante()
        {
            Contatti = String.Empty;
        }
    }
}
