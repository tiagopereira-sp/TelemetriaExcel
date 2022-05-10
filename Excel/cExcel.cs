using System.Collections.Generic;

namespace Excel
{
    class cExcel
    {
        public static List<cExcel> lExcel = new List<cExcel>();
        public string Seq { get; set; }
        public string Hora { get; set; }
        public string HoraRecepcao { get; set; }
        public string Velocidade { get; set; }
        public string Coordenadas { get; set; }
        public string Altitude { get; set; }
        public string Localizacao { get; set; }
        public string Parametros { get; set; }

        public cExcel() { }

        public cExcel(string hora, string horaRecepcao, string velocidade, string coordenadas, string altitude, string localizacao, string paramentros) 
        { 
            Hora = hora;
            HoraRecepcao = horaRecepcao;
            Velocidade = velocidade;    
            Coordenadas = coordenadas;
            Altitude = altitude;
            Localizacao = localizacao;
            Parametros = paramentros;
        }
    }
}
