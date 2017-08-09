//
// AutoImportSvc
// Um serviço para Windows para funcionamento do AutoImport
// Por Christian Haagensen Gontijo, jun/2017
//
using System;
using System.IO;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace AutoImportSvc
{
    public partial class AutoImportSvc : ServiceBase
    {

        string sINI = "";
        string sStrCon = "";
        int minutosAteProximaVerificacaoPastas = 30;
        Boolean usarModoDebug = false;
        DisysDAO.CAIHorarios listaDeHorarios = new DisysDAO.CAIHorarios();
        DisysDAO.CAIMaquinas listaDeMaquinas = new DisysDAO.CAIMaquinas();

        public AutoImportSvc()
        {
            InitializeComponent();
			// Texto sobre construção de serviços em C#:
			// https://msdn.microsoft.com/en-us/library/zt39148a(v=vs.110).aspx
            eventLog1 = new System.Diagnostics.EventLog();
            if (!System.Diagnostics.EventLog.SourceExists("AutoImportSvc")) {
                System.Diagnostics.EventLog.CreateEventSource("AutoImportSvc", "AILog");
            }
            eventLog1.Source = "AutoImportSvc";
        }

        //
        // Lê qual INI será usado pelo sistema, lendo "AUTOIMPORT.CONFIG".
        //
        // Se este arquivo estiver vazio, o primeiro INI do Disys que o programa encontrar
        // será colocado no AUTOIMPORT.CONFIG, para que o serviço possa continuar.
        //
        private string ObtemINI()
        {
            const string ARQUIVOCONFIG = @"C:\Disys\CONFIG\AutoImport.config";
            FnINI INI = new FnINI();
            sINI = INI.ReadINI(ARQUIVOCONFIG, "Srv", "INI");
            if (sINI == "")
            {
                DirectoryInfo di = new DirectoryInfo(@"C:\Disys\CONFIG");
                foreach (var file in di.EnumerateFiles("*.INI"))
                {
                    if (file.Name.ToLower() != "disysviewer.ini") {
                        INI.WriteINI(ARQUIVOCONFIG, "Srv", "INI", file.Name);
                        sINI = file.Name;
                        break;
                    }
                }
            }

            // Se mesmo assim não houver nada...
            if (sINI == "")
            {
                eventLog1.WriteEntry("Nenhum INI especificado em AutoImport.config. Saindo.");
                Stop();
            }

            int minutos = 0;
            try {
                minutos = Convert.ToInt16(INI.ReadINI(ARQUIVOCONFIG, "Srv", "MinutosAteProximaVerificacaoPastas"));
            } catch {
                minutos = 30;
            }
            if (minutos<1) minutos=30;
            minutosAteProximaVerificacaoPastas = minutos;

            string sDebug = INI.ReadINI(ARQUIVOCONFIG, "Srv", "ModoDebug");
            usarModoDebug = (sDebug == "1" || sDebug.ToUpper().StartsWith("T") || sDebug.ToUpper().StartsWith("Y") || sDebug.ToUpper().StartsWith("S"));
            if (usarModoDebug) eventLog1.WriteEntry(@"Entradas de Debug para o AutoImportWorker serão gravadas em arquivo próprio.");

            return sINI;
        }

        //
        // Obtenção da string de conexão
        //
        private string ObtemStrCon()
        {
            DisysDAO.DadosConexao dc = new DisysDAO.DadosConexao();
            return dc.ObtemStringConexaoBanco(sINI);
        }

        //
        // Verifica se o processamento será feito nesta máquina.
        //
        private Boolean ProcessamentoFeitoNestaMaquina()
        {
            DisysDAO.CAIPadrao obj = new DisysDAO.CAIPadrao();
            DisysDAO.CAIPadraoDAO dao = new DisysDAO.CAIPadraoDAO();
            dao.StringConexao = sStrCon;
            obj = dao.Obtem();
            if (obj != null)
                if (obj.LocalOp == 0) {
                    if (obj.Maquina_Remota == Environment.MachineName.ToUpper())
                        return true;
                } else {
                    return true;
                }
            return false;
        }

        private void ObtemListaPastasMaquina()
        {
            if (listaDeMaquinas == null || listaDeMaquinas.Count < 1)
            {
                DisysDAO.CAIMaquina maquina = new DisysDAO.CAIMaquina();
                maquina.Maquina = Environment.MachineName;
                DisysDAO.CAIMaquinaDAO daoMaq = new DisysDAO.CAIMaquinaDAO();
                daoMaq.StringConexao = sStrCon;
                listaDeMaquinas = daoMaq.Localiza(maquina);
                eventLog1.WriteEntry("Número de entradas definidas para este computador: " + listaDeMaquinas.Count);
            }
        }

        private void ObtemListaHorarios()
        {
            if (listaDeHorarios == null || listaDeHorarios.Count < 1)
            {
                DisysDAO.CAIHorarioDAO daoHorario = new DisysDAO.CAIHorarioDAO();
                daoHorario.StringConexao = sStrCon;
                listaDeHorarios = daoHorario.Localiza(new DisysDAO.CAIHorario());
                eventLog1.WriteEntry("Número de horários encontrados: " + listaDeHorarios.Count);
            }
        }
        
        protected override void OnStart(string[] args)
        {
            eventLog1.WriteEntry("Serviço iniciando...");

            try {
                sINI = ObtemINI();
            } catch (Exception ex) {
                eventLog1.WriteEntry("Erro geral ao determinar arquivo INI do AutoImport: " + ex.Message + ". Saindo.");
                Stop();
                return;
            }
            if (!sINI.StartsWith(@"C:\Disys")) sINI = @"C:\Disys\CONFIG\" + sINI;
            eventLog1.WriteEntry("Lendo dados do arquivo " + sINI);

            try
            {
                sStrCon = ObtemStrCon();
            } catch (Exception ex) {
                eventLog1.WriteEntry("Erro ao obter string de conexão ao banco de dados: " + ex.Message + ". Saindo.");
                Stop();
                return;
            }

            //
            // Continuamos?
            //
            if (!ProcessamentoFeitoNestaMaquina()) {
                eventLog1.WriteEntry("A pasta de arquivos do AutoImport está em outro computador. Portanto, este serviço nada fará.\n\n" +
                    "(Se, posteriormente, o AutoImport for configurado para usar operação local, basta reiniciar este serviço ou este computador)");
                Stop();
                return;
            }
            
            //
            // Obtém uma lista com todas as máquinas cadastradas
            //
            try
            {
                ObtemListaPastasMaquina();
                if (listaDeMaquinas == null || listaDeMaquinas.Count < 1) {
                    eventLog1.WriteEntry("Não há pastas cadastradas para este computador no momento. O serviço continuará, e fará nova verificação em um minuto.");
                }
            } catch (Exception ex) {
                eventLog1.WriteEntry("Erro geral ao determinar lista de pastas deste computador: " + ex.Message + ". Saindo.");
                Stop();
                return;
            }

            //
            // Obtém uma lista com todos os horários cadastrados
            //
            try
            {
                ObtemListaHorarios();
                if (listaDeHorarios == null || listaDeHorarios.Count < 1) {
                    eventLog1.WriteEntry("Não há horários definidos neste momento. O serviço continuará, e fará nova verificação em um minuto.");
                }
            } catch (Exception ex) {
                eventLog1.WriteEntry("Erro geral ao determinar lista de horários: " + ex.Message + ". Saindo.");
                Stop();
                return;
            }

            eventLog1.WriteEntry("Criando timer para atualização das listas. Próxima atualização em " + minutosAteProximaVerificacaoPastas + " minutos.");
            System.Timers.Timer timerRefresh = new System.Timers.Timer();
            timerRefresh.Interval = minutosAteProximaVerificacaoPastas * 60 * 1000; // padrão: meia hora
            timerRefresh.Elapsed += new System.Timers.ElapsedEventHandler(this.OnTimerRefresh);
            timerRefresh.Start();
            
            eventLog1.WriteEntry("Criando timer para monitoramento das pastas. Próxima atualização em um minuto.");
            System.Timers.Timer timer1 = new System.Timers.Timer();
            timer1.Interval = 60 * 1000; // 60 segundos
            timer1.Elapsed += new System.Timers.ElapsedEventHandler(this.OnTimer);
            timer1.Start();
        }

        //
        // Força uma atualização nas listas de horários e de máquinas. Disparado a cada meia hora.
        //
        private void OnTimerRefresh(object sender, System.Timers.ElapsedEventArgs e)
        {
            eventLog1.WriteEntry("Atualizando listas internas.");

            try {
                listaDeMaquinas = null; 
                ObtemListaPastasMaquina();
            } catch (Exception ex) {
                eventLog1.WriteEntry("Erro geral ao determinar lista de pastas do computador: " + ex.Message + ".");
            }

            try {
                listaDeHorarios = null;
                ObtemListaHorarios();
            } catch (Exception ex) {
                eventLog1.WriteEntry("Erro geral ao determinar lista de horários: " + ex.Message + ".");
            }
        }

        private void OnTimer(object sender, System.Timers.ElapsedEventArgs e)
        {

            //
            // Verifica se as listas estão *realmente* preenchidas 
            // (pode ter havido falha de rede na hora de iniciar o serviço, por exemplo).
            //
            try {
                ObtemListaPastasMaquina();
            } catch (Exception ex) {
                eventLog1.WriteEntry("Erro geral ao determinar lista de pastas do computador: " + ex.Message + ".");
                return;
            }

            try {
                ObtemListaHorarios();
            } catch (Exception ex) {
                eventLog1.WriteEntry("Erro geral ao determinar lista de horários: " + ex.Message + ".");
                return;
            }

            if (usarModoDebug) 
                eventLog1.WriteEntry("Verificação às " + DateTime.Now.ToString("HH:mm") + 
                                    "\nEsta mensagem está sendo apresentada porque o serviço está usando o Modo Debug.");

            //
            // Para cada hora existente na lista de horários, verifica se há uma série com arquivos
            //
            try
            {
                foreach (DisysDAO.CAIHorario hora in listaDeHorarios)
                {

                    if (usarModoDebug)
                        eventLog1.WriteEntry("Verificando horários da série " + hora.Cod_Assunto, EventLogEntryType.Information);

                    if (Convert.ToDateTime(hora.HORA1).ToString("HH:mm") == DateTime.Now.ToString("HH:mm") ||
                        Convert.ToDateTime(hora.HORA2).ToString("HH:mm") == DateTime.Now.ToString("HH:mm") ||
                        Convert.ToDateTime(hora.HORA3).ToString("HH:mm") == DateTime.Now.ToString("HH:mm") ||
                        Convert.ToDateTime(hora.HORA4).ToString("HH:mm") == DateTime.Now.ToString("HH:mm") ||
                        Convert.ToDateTime(hora.HORA5).ToString("HH:mm") == DateTime.Now.ToString("HH:mm")) {

                        eventLog1.WriteEntry("Horário monitorado atingido para a série " + hora.Cod_Assunto, EventLogEntryType.Information);

                        foreach (DisysDAO.CAIMaquina maq in listaDeMaquinas)
                        {

                            if (maq.Cod_Assunto == hora.Cod_Assunto) {

                                eventLog1.WriteEntry("Série " + hora.Cod_Assunto + " será processada agora.", EventLogEntryType.Information);

                                // Processa cada arquivo do path informado
                                DirectoryInfo di = new DirectoryInfo(maq.Pasta_Base);
                                foreach (var file in di.EnumerateFiles())
                                {
                                    string arq = Path.Combine(maq.Pasta_Base, file.Name);
                                    string param = sINI + "|" + maq.Cod_Operador + "|" + arq;
                                    if (usarModoDebug) param += "|/DEBUG";
                                    eventLog1.WriteEntry("Disparando worker para <" + arq + ">.", EventLogEntryType.Information);
                                    ProcessStartInfo proggie = new ProcessStartInfo(@"C:\Disys\AutoImportWorker.exe", param);
                                    proggie.UseShellExecute = false; // evitar uac?
                                    proggie.WindowStyle = ProcessWindowStyle.Hidden;
                                    Process.Start(proggie);
                                    System.Threading.Thread.Sleep(500); // evitar workers com mesma hora de execução...
                                }

                            }

                        } // foreach

                    } // if
                } // foreach
            } catch (Exception ex) {
                eventLog1.WriteEntry("Erro geral no processamento dos horários: " + ex.Message + ".");
                return;
            }
            
        }

        protected override void OnStop()
        {
            eventLog1.WriteEntry("Serviço finalizado.");
        }
    }
}
