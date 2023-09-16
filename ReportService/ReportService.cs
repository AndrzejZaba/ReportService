using ReportService.Repositories;
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

namespace ReportService
{
    public partial class ReportService : ServiceBase
    {
        private const int SendHour = 8;
        private const int IntervalMinutes = 60;
        private Timer _timer = new Timer(IntervalMinutes * 60000);
        private ErrorRepository _errorRepository = new ErrorRepository();
        private ReportRepository _reportRepository = new ReportRepository();
        public ReportService()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            _timer.Elapsed += DoWork;
            _timer.Start();
        }

        private void DoWork(object sender, ElapsedEventArgs e)
        {
            SendError();
            SendReport();
        }

        private void SendError()
        {
            var errors = _errorRepository.GetLastErrors(IntervalMinutes);

            if (errors == null || !errors.Any())
                return;

            //send mail
        }

        private void SendReport()
        {
            var actualHour = DateTime.Now.Hour;

            if (actualHour < SendHour)
                return;

            var report = _reportRepository.GetLastNotSentReport();

            if (report == null)
                return;

            //send mail
            _reportRepository.ReportSent(report);

        }

        protected override void OnStop()
        {
        }
    }
}
