using EmailSender;
using ReportService.Core.Repositories;
using ReportService.Core;    
using ReportService.Core.Domains;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using System.Configuration;
using Cipher;

namespace ReportService
{
    public partial class ReportService : ServiceBase
    {
        private static readonly NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();

        private const int SendHour = 8;
        private const int IntervalMinutes = 1;
        private Timer _timer = new Timer(IntervalMinutes * 60000);
        private ErrorRepository _errorRepository = new ErrorRepository();
        private ReportRepository _reportRepository = new ReportRepository();
        private Email _email;
        private GenerateHtmlEmail _htmlEmail = new GenerateHtmlEmail();
        private string _emailReceiver;
        private StringCipher _stringCipher = new StringCipher("73C09781-855C-4574-8928-95D772C04904");
        public ReportService()
        {
            InitializeComponent();

            try
            {
                _emailReceiver = ConfigurationManager.AppSettings["ReceiverEmail"];

                var encryptedPassword = ConfigurationManager.AppSettings["SenderEmailPassword"];
                if (encryptedPassword.StartsWith("encrypt:"))
                {
                    encryptedPassword = _stringCipher.Encrypt(encryptedPassword.Replace("encrypt:", ""));

                    var configFile = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                    configFile.AppSettings.Settings["SenderEmailPassword"].Value = encryptedPassword;
                    configFile.Save();
                }

                _email = new Email(new EmailParams
                {
                    HostSmtp = ConfigurationManager.AppSettings["HostSmtp"],
                    Port = Convert.ToInt32(ConfigurationManager.AppSettings["Port"]),
                    EnableSsl = Convert.ToBoolean(ConfigurationManager.AppSettings["EnableSsl"]),
                    SenderName = ConfigurationManager.AppSettings["SenderName"],
                    SenderEmail = ConfigurationManager.AppSettings["SenderEmail"],
                    SenderEmailPassword = DecryptSenderEmailPassword()
                });
            }
            catch (Exception ex)
            {
                Logger.Error(ex, ex.Message);
                throw new Exception(ex.Message);
            }
        }

        private string DecryptSenderEmailPassword()
        {
            var encryptedPassword = ConfigurationManager.AppSettings["SenderEmailPassword"];
            if (encryptedPassword.StartsWith("encrypt:"))
            {
                encryptedPassword = _stringCipher.Encrypt(encryptedPassword.Replace("encrypt:", ""));

                var configFile = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                configFile.AppSettings.Settings["SenderEmailPassword"].Value = encryptedPassword;
                configFile.Save();
            }

            return _stringCipher.Decrypt(encryptedPassword);
        }

        protected override void OnStart(string[] args)
        {
            _timer.Elapsed += DoWork;
            _timer.Start();

            Logger.Info("Service started...");
        }

        private async void DoWork(object sender, ElapsedEventArgs e)
        {
            

            try
            {
                await SendError();
                SendReport();
            }
            catch (Exception ex)
            {
                Logger.Error(ex, ex.Message);
                throw new Exception(ex.Message);
            }

        }

        private async Task SendError()
        {
            var errors = _errorRepository.GetLastErrors(IntervalMinutes);

            if (errors == null || !errors.Any())
                return;

            await _email.Send("Błędy w aplikacji", _htmlEmail.GenerateErrors(errors, IntervalMinutes), _emailReceiver);

            Logger.Info("Error sent...");
        }

        private async void SendReport()
        {
            var actualHour = DateTime.Now.Hour;

            if (actualHour < SendHour)
                return;

            var report = _reportRepository.GetLastNotSentReport();

            if (report == null)
                return;

            await _email.Send("Raport dobowy", _htmlEmail.GenerateReport(report), _emailReceiver);
            _reportRepository.ReportSent(report);

            Logger.Info("Report sent...");

        }

        protected override void OnStop()
        {
            Logger.Info("Service stopped...");

        }
    }
}
