using Spire.Xls;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;

namespace Excel
{
    public partial class FormExcel : Form
    {
        private string CaminhoOri = @"C:\";                                                                                 // Caminho de Origem do Arquivo
        private string ArquivoOri = "Telemetria.xlsx";                                                                      // Arquivo de Origem

        private string CaminhoDest = @"C:\";                                                                                // Caminho de Destino do Arquivo
        private string ArquivoDest = "Telemetria_Processada.xlsx";                                                          // Arquivo de Destino
        /// <summary>
        /// 
        /// </summary>
        public FormExcel() { InitializeComponent(); }                                                                       // Inicializa os componentes
        /// <summary>
        /// Executa quando o formulario se abre
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FormExcel_Load(object sender, EventArgs e)
        {
            lblCaminho.Text = string.Concat(CaminhoOri, ArquivoOri);                                                        // Mostra na pela o caminho do arquivo
        }
        /// <summary>
        /// Executa quando pressiona o botão processar
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnProcessar_Click(object sender, EventArgs e)
        {
            DeletaArquivo();                                                                                                // Deleta o arquivo de destino existente
            GetExcel();                                                                                                     // Gera o arquivo excel novo
            if (cExcel.lExcel.Count == 0) return;                                                                           // Não leu a planilha de origem

            string[] Col = GetColunas(cExcel.lExcel);                                                                       // Busca as colunas para criar o novo arquivo
            if (Col == null) return;                                                                                        // Se não leu as colunas não pode continuar
            
            CriarExcelColunas(Col);                                                                                         // Cria o Excel com as colunas
            GerarRelatorio();                                                                                               // Gera o Relatorio
            Mon("Processamento Concluido!");                                                                                // Debug
        }
        /// <summary>
        /// Monitor para debug
        /// </summary>
        /// <param name="texto"></param>
        private void Mon(string texto)
        {
            labelHeader.Text =texto;                                                                                        // Imprime na tela          
            Application.DoEvents();                                                                                         // Atualiza a tela
            Thread.Sleep(1000);                                                                                             // Aguarda 1 segundo
        }
        /// <summary>
        /// Quando click no botão sair
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSair_Click(object sender, EventArgs e)
        {
            Application.Exit();                                                                                             // Cai fora
        }
        private void DeletaArquivo()
        {
            try
            {
                Mon("Vamos apagar o arquivo de destino");                                                                   // Debug
                pgbProcesso.Maximum = 100;                                                                                  // Valor maximo do progress bar
                string path = string.Concat(CaminhoDest, ArquivoDest);                                                      // Caminho do arquivo
                if (File.Exists(path))                                                                                      // Verifica se existe o arquivo no caminho
                {
                    Mon(string.Concat("Arquivo: ", ArquivoDest, " foi encontrado!"));                                       // Debug
                    File.Delete(path);                                                                                      // Apaga o arquivo
                    for (int i = 0; i < 100; i++) pgbProcesso.Value = i;                                                    // Processa o progress bar
                    Mon(string.Concat("Arquivo: ", ArquivoDest, " foi apagado com sucesso!"));                              // Debug
                }
                pgbProcesso.Value = 0;                                                                                      // Zera o valor do progress bar
            }
            catch (Exception) { Mon("Er.Deletar arquivo já processado"); }                                                  // Debug
        }
        /// <summary>
        /// 
        /// </summary>
        private void GetExcel()
        {
            try
            {
                string path = string.Concat(CaminhoOri, ArquivoOri);                                                        // Caminho do arquivo de origem
                Mon(string.Concat("Lendo planilha de Origem :", ArquivoOri));                                               // Debug
                var xls = new XLWorkbook(path);                                                                             // Acessa a planilha
                var planilha = xls.Worksheets.First(w => w.Name == "Planilha1");                                            // Nome da planilha
                var totalLinhas = planilha.Rows().Count();                                                                  // Quantidade de linhas da planilha
                pgbProcesso.Maximum = totalLinhas;                                                                          // Diz ao progress bar a quantidade do valor maximo
                cExcel.lExcel.Clear();                                                                                      // Limpa a lista

                for (int l = 2; l <= totalLinhas; l++)                                                                      // Processa as linhas da planilha
                {
                    if ((planilha.Cell($"H{l}").Value.ToString().Contains("GTFRI")) ||
                        planilha.Cell($"E{l}").Value.ToString() == "---") { pgbProcesso.Value = l; continue; }              // Ignora esses registros   

                    cExcel.lExcel.Add(new cExcel()
                    {
                        Seq = planilha.Cell($"A{l}").Value.ToString(),
                        Hora = planilha.Cell($"B{l}").Value.ToString(),
                        HoraRecepcao = planilha.Cell($"C{l}").Value.ToString(),
                        Velocidade = planilha.Cell($"D{l}").Value.ToString(),
                        Coordenadas = planilha.Cell($"E{l}").Value.ToString(),
                        Altitude = planilha.Cell($"F{l}").Value.ToString(),
                        Localizacao = planilha.Cell($"G{l}").Value.ToString(),
                        Parametros = planilha.Cell($"H{l}").Value.ToString().Replace(" ", ""),
                    });                                                                                                     // Popula a lista com os valores da planilha
                    pgbProcesso.Value = l;                                                                                  // Incrementa o progressbar
                }
                pgbProcesso.Value = 0;                                                                                      // Zera o valor do progress bar
                xls.Dispose();                                                                                              // Libera o recurso do excel
            }
            catch (Exception) { Mon("Er.Lendo planilha de Origem"); }                                                       // Debug
        }
        /// <summary>
        /// Busca os dados da planilha
        /// </summary>
        /// <returns></returns>
        private string[] GetColunas(List<cExcel> lxExcel)
        {
            try
            {
                List<string> lCol = new List<string>();
                Mon("Parse da coluna parametros (Atributos)");                                                              // Debug
                string coluna = string.Empty;                                                                               // Cria a variavel para manipular
                
                pgbProcesso.Maximum = lxExcel.Count();                                                                      // Atribui a quantidade da lista no progress bar
                lCol.Add("Seq"); lCol.Add("Hora"); lCol.Add("Horaderecepcao"); lCol.Add("Velocidade_km_h");
                lCol.Add("Coordenadas"); lCol.Add("Altitude_m"); lCol.Add("Localizacao");                                   // Coloca os campos na lista
                foreach (var l in lxExcel)                                                                                  // Processa os campos da lista
                {
                    var xCol = l.Parametros.Split(',');                                                                     // Separa os campos da coluna parametros
                    foreach (var x in xCol)
                    {
                        string campo = string.Empty;
                        switch (x.Substring(0, x.IndexOf("=")))                                                             // Troca os campos quando leitura por odb k-line é dif Can
                        {
                            case "can_state"            : campo = "obd_connect=";           break;
                            case "pwr_ext"              : campo = "obd_pwr=";               break;
                            case "can_vehicle_speed"    : campo = "obd_speed=";             break;
                            case "can_total_dist_hect"  : campo = "obd_mileage=";           break;
                            case "can_engine_rpm"       : campo = "engine_rpm=";            break;
                            case "can_eng_cool_temp"    : campo = "engine_coolant_temp=";   break;
                            case "I/O"                  : campo = "IO="   ;                 break;
                            default                     : campo = x;                        break;
                        }
                        if (campo.IndexOf("=") == -1) continue;
                        var v = lCol.Where(c => c == campo.Substring(0, campo.IndexOf("="))).Count();
                        if (v == 0) lCol.Add(campo.Substring(0, campo.IndexOf("=")));
                    }
                    pgbProcesso.Value += 1;                                                                                 // Imcrementa o progress bar
                }
                pgbProcesso.Value = 0;                                                                                      // Zera o valor do progress bar
                
                StringBuilder str = new StringBuilder();                                                                    // Variavel para utilizar na montagem do arquivo
                foreach (var x in lCol) { str.Append("["); str.Append(x); str.Append("],"); }                               // Monta as colunas
                lCol.Clear();
                coluna = str.ToString().Remove(str.Length - 1, 1);                                                          // Remove a ultima virgula que sobro quando montou o texto no for
                return coluna.Split(',');                                                                                   // Retorma as colunas utilizando o split
            }
            catch (Exception) { Mon("Er.Coletando as colunas do arquivo"); }                                                // Debug
            return null;                                                                                                    // Retorna null se deu merda
        }
        /// <summary>
        /// Cria o arquivo para ser processado
        /// </summary>
        /// <param name="colunas"></param>
        private void CriarExcelColunas(string[] colunas)
        {
            Mon("Gerando o arquivo para ser processado!");                                                                  // Debug
            int i = 0;
            pgbProcesso.Maximum = colunas.Count();                                                                          // Atribui a quantidade de colunas no progress bar
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(CaminhoOri + ArquivoOri);                                                                 // Abrir o arquivo atraves do caminho
            Worksheet sheet = workbook.Worksheets["Planilha2"];                                                             // Abre a planilha 2
            Worksheet sheet1 = workbook.Worksheets["Planilha1"];                                                            // Abre a planilha 2
            foreach (var t in colunas)
            {
                sheet.Cells[i++].Value = t.ToString().Replace("[", "").Replace("]", "").Replace(" ", "");
                pgbProcesso.Value += 1;
            }
            sheet1.Remove();                                                                                                // Remove a planilha 1
            workbook.SaveToFile(CaminhoDest + ArquivoDest);                                                                 // Salva o novo arquivo já com as colunas
            workbook.Dispose();                                                                                             // Libera o recurso
            pgbProcesso.Value = 0;                                                                                          // Zera o progress bar
        }
        /// <summary>
        /// Gera o relatorio
        /// </summary>
        private void GerarRelatorio()
        {   
            string path = string.Concat(CaminhoDest, ArquivoDest);
            using (OleDbConnection conn = new OleDbConnection(string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0 Xml;HDR=YES;ReadOnly=False';", path)))
            {
                OleDbCommand cmd = new OleDbCommand();
                Mon("Processando o Arquivo de Telemetria.....");                                                            // Debug
                try
                {
                    conn.Open();                                                                                            // Abre a conexão
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;

                    StringBuilder str = new StringBuilder();          
                    pgbProcesso.Maximum = cExcel.lExcel.Count();
                    foreach (cExcel p in cExcel.lExcel)
                    {
                        if (p.Parametros.Length == 0) continue;
                        p.Parametros = string.Concat("Seq,", "Hora,", "Horaderecepcao,", "Velocidade_km_h,", "Coordenadas,", "Altitude_m,",
                            "Localizacao,", p.Parametros);
                        string Colunas = string.Empty;
                        string[] Col = p.Parametros.Split(',');
                        foreach (var c in p.Parametros.Split(','))
                        {
                            string xCol = string.Empty;
                            switch (c.Split('=')[0])
                            {
                                case "can_state": xCol = "obd_connect"; break;
                                case "pwr_ext": xCol = "obd_pwr"; break;
                                case "can_vehicle_speed": xCol = "obd_speed"; break;
                                case "can_total_dist_hect": xCol = "obd_mileage"; break;
                                case "can_engine_rpm": xCol = "engine_rpm"; break;
                                case "can_eng_cool_temp": xCol = "engine_coolant_temp"; break;
                                case "I/O": xCol = "IO"; break;
                                default: xCol = c.Split('=')[0]; break;
                            }
                            Colunas += string.Concat("[", xCol.Replace(" ", ""), "],");
                        }

                        Colunas = Colunas.Remove(Colunas.Length - 1, 1);
                        str.Clear();
                        str.Append("INSERT INTO [Planilha2$] (");
                        str.Append(Colunas); str.Append(")");
                        str.Append("VALUES (");
                        string Query = string.Empty;
                        int i = 0;
                        foreach (var col in p.Parametros.Split(','))
                        {
                            string campo = string.Empty;
                            switch (col.Split('=')[0].Replace(" ", ""))
                            {
                                case "can_state": campo = "obd_connect"; break;
                                case "pwr_ext": campo = "obd_pwr"; break;
                                case "can_vehicle_speed": campo = "obd_speed"; break;
                                case "can_total_dist_hect": campo = "obd_mileage"; break;
                                case "can_engine_rpm": campo = "engine_rpm"; break;
                                case "can_eng_cool_temp": campo = "engine_coolant_temp"; break;
                                case "I/O": campo = "IO"; break;
                                default: campo = col.Split('=')[0].Replace(" ", ""); break;
                            }

                            string val = string.Concat("@", campo);
                            str.Append(val); str.Append(",");

                            switch (i)
                            {
                                case 0: cmd.Parameters.Add(val, OleDbType.VarChar).Value = p.Seq; break;
                                case 1: cmd.Parameters.Add(val, OleDbType.VarChar).Value = p.Hora; break;
                                case 2: cmd.Parameters.Add(val, OleDbType.VarChar).Value = p.HoraRecepcao; break;
                                case 3: cmd.Parameters.Add(val, OleDbType.VarChar).Value = p.Velocidade; break;
                                case 4: cmd.Parameters.Add(val, OleDbType.VarChar).Value = p.Coordenadas; break;
                                case 5: cmd.Parameters.Add(val, OleDbType.VarChar).Value = p.Altitude; break;
                                case 6: cmd.Parameters.Add(val, OleDbType.VarChar).Value = p.Localizacao; break;
                                default:cmd.Parameters.Add(val, OleDbType.VarChar).Value = col.Split('=')[1]; break;
                            }
                            i++;
                        }
                        Query = str.Replace(",", ")", str.Length - 1, 1).ToString();
                        if (!string.IsNullOrEmpty(Query))
                        {
                            cmd.CommandText = Query;
                            cmd.ExecuteNonQuery();
                        }
                        cmd.Parameters.Clear();
                        pgbProcesso.Value += 1;
                    }
                }
                catch (Exception) { Mon("Er. GerarRelatorio"); }                                                            // Debug
            }
        }
        /// <summary>
        /// Busca o caminho para os arquivos
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCaminhoArq_Click(object sender, EventArgs e)
        {
            if (ofdCaminho.ShowDialog() == DialogResult.OK)
            {
                CaminhoOri = ofdCaminho.FileName.Replace(ofdCaminho.SafeFileName ,"");
                CaminhoDest = CaminhoOri;
                ArquivoOri = ofdCaminho.SafeFileName;
                if ((CaminhoOri.Length + ArquivoOri.Length) > 50)
                    lblCaminho.Text = string.Concat(CaminhoOri, ArquivoOri).Substring(0, 50) + "...";
                else
                    lblCaminho.Text = ofdCaminho.FileName;
            }
        }
    }
}