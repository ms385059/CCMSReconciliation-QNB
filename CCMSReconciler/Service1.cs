using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using Microsoft.Win32;
using Avanza.iSuite.DAL;
using Avanza.CCMS.DAL;
using System.Reflection;
using System.Data.SqlClient;
using DAL;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Diagnostics.Eventing.Reader;
using static Avanza.CCMS.DAL.BatchSchedule;
using System.Threading.Tasks;
using System.Configuration;
using System.Collections;

namespace CCMSReconciler
{
    public partial class Service1 : ServiceBase
    {
        Stopwatch stopWatch = new Stopwatch();
        Timer timerDoWork;
        Timer timerScheduleThreadForExecution;
        public static AppSetting appSettings;

        int withdrawalTransactionType = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["withdrawalTransactionType"]);
        int depositTransactionType = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["depositTransactionType"]);

        int defaultInterval = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["RescheduleReconciler"]); //minutes

        public static bool isBna = false;

        public Service1()
        {
            InitializeComponent();
        }

        public void OnDebug()
        {
            OnStart(null);
        }
        protected override void OnStart(string[] args)
        {
            timerScheduleThreadForExecution = new Timer(ScheduleThreadForExecution, null, new TimeSpan(0, 0, 25), new TimeSpan(0, 0, 0, 0, -1));
            EventLog.WriteEntry("CCMSReconciler", "Thread schedular sent startup request", EventLogEntryType.Information);
            //
        }
        void ScheduleThreadForExecution(object state)
        {
            try
            {

                string connectionStr = (string)Registry.LocalMachine.OpenSubKey("SOFTWARE\\NCR\\CCMS").GetValue("ConnectionString", "");
                connectionStr = Encryption.Cryptic.DecryptString(connectionStr);
                ConnectionFactory.Initialize(connectionStr, true);
                appSettings = AppSetting.LoadAppSetting("1=1");
                XmlLogWriter.InitXmlLogWriter(String.Format("{0}\\CCMSReconciler_{1:yyMMMdd}.txt", appSettings.LogFilePath, DateTime.Now));
                LogableTask.LogMonoActivityTask("DisplayVersion", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "Version : CCMS Reconciler 2.0.0.0, Creation Date 11/08/2026");

                LogableTask.LogMonoActivityTask("Schedular", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "Worker Threads begin execution after 15 seconds.");

                timerDoWork = new Timer(DoWork, null, new TimeSpan(0, 0, 15), new TimeSpan(0, 0, 0, 0, -1));

                EventLog.WriteEntry("CCMSReconciler", "Service Started Successfully", EventLogEntryType.Information);

            }
            catch (Exception ex)
            {
                //trying to log error in event log if its not full.
                try
                {
                    EventLog.WriteEntry("CCMSReconciler", ex.Message + " " + ex.StackTrace, EventLogEntryType.Error);
                    //EventLog.WriteEntry("CurrencyMngServer", "Service is idle", EventLogEntryType.Warning);
                    // timerScheduleThreadForExecution.Change(new TimeSpan(0, appSettings.RefreshInterval, 0), new TimeSpan(0, 0, 0, 0, -1));
                    timerScheduleThreadForExecution.Change(new TimeSpan(0, 15, 0), new TimeSpan(0, 0, 0, 0, -1));//updated refreshinterval as batch rerun takes more time
                }
                catch (Exception innerException)
                {
                    LogableTask.LogMonoActivityTask("Schedular", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "Worker Threads error on Exception timerScheduleThreadForExecution");
                }
            }
        }
        void DoWork(object state)
        {
            //try
            //{
            //    if (CancelScheduled == 1)
            //    {
            //        StringBuilder builder = new StringBuilder();//Host_Cardless_Transactions = 0, Host_Amount_Cardless = 0,
            //        builder.Append("update reconciliation_batch set status='' where status = 'Scheduled';");
            //        ConnectionFactory.ExecuteQuery(builder.ToString());
            //    }
            //}
            //catch (Exception ex)
            //{
            //    // task.Log(MethodBase.GetCurrentMethod(), TraceLevel.Error, "CancelScheduled:" + ex);
            //}
            timerDoWork.Change(-1, -1);

            try
            {
                XmlLogWriter.InitXmlLogWriter(String.Format("{0}\\CCMSReconciler_{1:yyMMMdd}.txt", appSettings.LogFilePath, DateTime.Now));
                ThreadPool.SetMaxThreads(25, 25);

                if (appSettings.RefreshInterval == 0)
                    appSettings.RefreshInterval = 10;
                else if (appSettings.RefreshInterval > 720)
                    appSettings.RefreshInterval = 720;
                try
                {
                    LogableTask.DefaultTraceLevel = (TraceLevel)Enum.Parse(typeof(TraceLevel), appSettings.ServiceLogLevel);
                }
                catch
                {
                    LogableTask.DefaultTraceLevel = TraceLevel.Info;
                    LogableTask.LogMonoActivityTask("GetTraceLevel", MethodBase.GetCurrentMethod(), TraceLevel.Error, "Failed to extract trace level from database");
                }
                DoReconcile();
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry("CCMSReconciler", "Error in DoWork(). detail: " + ex.Message + ex.StackTrace, EventLogEntryType.Error);
            }

            try
            {
                if (appSettings == null)
                    EventLog.WriteEntry("CCMSReconciler", "Error in DoWork(). detail: AppSettings not loaded");
                else
                {
                    //defaultInterval = appSettings.RefreshInterval;
                    LogableTask.LogMonoActivityTask("Dowork", MethodBase.GetCurrentMethod(), TraceLevel.Info, "Going to sleep for " + defaultInterval + " sec");
                }
            }
            catch (Exception ex)
            {
                LogableTask.LogMonoActivityTask("Dowork", MethodBase.GetCurrentMethod(), TraceLevel.Error, ex);
            }
            finally
            {
                timerDoWork.Change(new TimeSpan(0, 0, defaultInterval), new TimeSpan(0, 0, 0, 0, -1));
            }
        }
        private void DoReconcile()
        {

            LogableTask task = LogableTask.NewTask("DoReconcile");
            ReconciliationBatch.ReconciliationBatchReader batchScheduleReader = null;
            SqlCommand cmd = null;

            try
            {
                cmd = ConnectionFactory.GetNewCommand(false);

                // NOTE: The Reconciliation Scheduler changes the status of a batch to "Scheduled" whenever any Dispense or Deposit files are uploaded for that particular batch; no matter the existing status is Completed, Dispense Completed or Deposit Completed. 

                batchScheduleReader = ReconciliationBatch.ExecuteReader($"status = '{BatchStatus.Scheduled}' and retry_count>0 and atm_id in (select atm_id from atm where is_active=1) order by last_invoked_at desc");
                cmd.Connection.Open();

                while (batchScheduleReader.Read())
                {
                    try
                    {
                        int currentBatchId = batchScheduleReader.CurrentReconciliationBatch.ReconciliationBatchId;

                        LogableTask.LogMonoActivityTask("prcBatch", MethodBase.GetCurrentMethod(), TraceLevel.Info, "Processing batch id " + currentBatchId);

                        batchScheduleReader.CurrentReconciliationBatch.RetryCount--;
                        batchScheduleReader.CurrentReconciliationBatch.LastInvokedAt = DateTime.Now;
                        batchScheduleReader.CurrentReconciliationBatch.Save();

                        ReconBatchInfo.DeleteReconBatchInfos("Reconciliation_Batch_Id = " + currentBatchId);

                        var reconBatchInfo = new ReconBatchInfo();
                        reconBatchInfo.ReconciliationBatchId = currentBatchId;
                        reconBatchInfo.AtmId = batchScheduleReader.CurrentReconciliationBatch.AtmId;
                        reconBatchInfo.TranDate = batchScheduleReader.CurrentReconciliationBatch.TransactionStartDate;
                        reconBatchInfo.ProcessingDateTime = DateTime.Now;

                        #region DISPENSE

                        LogableTask.LogMonoActivityTask("Processing", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "Processing Dispense Transactions.");

                        cmd.CommandText = "select * from reconciliation_host_data where reconciliation_batch_id =  " + currentBatchId;

                        var dtHost = PopulateDataTable(cmd);

                        cmd.CommandText = "select * from reconciliation_switch_data where reconciliation_batch_id = "
                            + currentBatchId;

                        var dtSwitch = PopulateDataTable(cmd);

                        cmd.CommandText = "select * from vejparsedtransactions ej_parsed_transactions inner join atm on ej_parsed_transactions.atm_id = atm.atm_id and amount is not null and trxn_Datetime >= convert(datetime,'" + batchScheduleReader.CurrentReconciliationBatch.TransactionStartDate.ToString("dd/MM/yyyy") + "',103) " + " and trxn_datetime<=convert(datetime,'" + batchScheduleReader.CurrentReconciliationBatch.TransactionStartDate.ToString("dd/MM/yyyy") + " 23:59:59',103) and atm.ATM_id = " + batchScheduleReader.CurrentReconciliationBatch.AtmId +
                          " and transaction_type_id in (" + withdrawalTransactionType + ")"; // Just fethcing Withdrawal trxns. 38035

                        var dtEJ = PopulateDataTable(cmd);

                        LogableTask.LogMonoActivityTask("Datatables Population", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "All 3 Dispense Datatables populated from Database.");

                        LogableTask.LogMonoActivityTask("Reconcile() Calling", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "Calling Reconcile() Method for Dispense Transactions.");

                        isBna = false;
                        var result = Reconcile(dtEJ, dtSwitch, dtHost, currentBatchId);

                        LogableTask.LogMonoActivityTask("Reconcile() Executed", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "Reconcile() Method has executed for Dispense Transactions.");

                        if (result.ReconciliationTransactions.Count > 0)
                        {
                            // Amounts Disputed
                            reconBatchInfo.DisputedAmount = result.Summary.TotalAmountDisputed;
                            reconBatchInfo.DisputedOnUsEJAmount = result.Summary.AmountDisputedOnEj;
                            reconBatchInfo.DisputedOnUsSwitchAmount = result.Summary.AmountDisputedOnSwitch;
                            reconBatchInfo.DisputedOnUsHostAmount = result.Summary.AmountDisputedOnHost;

                            // Transactions Disputed
                            reconBatchInfo.DisputedOnUsTransactions = result.Summary.TotalTransDisputed;
                            reconBatchInfo.DisputedOnUsEJTransactions = result.Summary.TransDisputedOnEj;
                            reconBatchInfo.DisputedOnUsSwitchTransactions = result.Summary.TransDisputedOnSwitch;
                            reconBatchInfo.DisputedOnUsHostTransactions = result.Summary.TransDisputedOnHost;

                            batchScheduleReader.CurrentReconciliationBatch.AutoReconciledAmount = result.Summary.TotalAmountReconciled;
                            batchScheduleReader.CurrentReconciliationBatch.NoOfRecordsProcessed = result.Summary.NoOfRecordsProcessed;
                            batchScheduleReader.CurrentReconciliationBatch.NoOfRecordsReconciled = result.Summary.NoOfRecordsReconciledCompletely;
                            batchScheduleReader.CurrentReconciliationBatch.NoOfRecordsFailedToReconciled = result.Summary.NoOfRecordsFailedToReconciled;
                            batchScheduleReader.CurrentReconciliationBatch.FailureReason = "";

                            batchScheduleReader.CurrentReconciliationBatch.Status = batchScheduleReader.CurrentReconciliationBatch.Status == BatchStatus.DepositCompleted ? BatchStatus.Completed : BatchStatus.DispenseCompleted;

                            //ManageTx(result.ReconciliationTransactions, result.ReconciledTransactions, batchScheduleReader, reconBatchInfo);
                        }
                        else
                        {
                            LogableTask.LogMonoActivityTask("DoReconcile", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "Reconcile() Method didnot produce any Reconciliation Data for DISPENSE.");
                        }

                        #endregion

                        #region DEPOSIT

                        LogableTask.LogMonoActivityTask("Processing", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "Processing Deposit Transactions.");


                        cmd.CommandText = "select * from reconciliation_bna_host_data where reconciliation_batch_id =  " + currentBatchId;

                        var dtBnaHost = PopulateDataTable(cmd);

                        cmd.CommandText = "select * from reconciliation_bna_switch_data where reconciliation_batch_id = "
                            + currentBatchId;

                        var dtBnaSwitch = PopulateDataTable(cmd);

                        cmd.CommandText = "select * from vEjParsedBNATransaction ej_parsed_bna_transaction inner join atm on ej_parsed_bna_transaction.atm_id = atm.atm_id and amount_authorized is not null and trxn_Datetime >= convert(datetime,'" + batchScheduleReader.CurrentReconciliationBatch.TransactionStartDate.ToString("dd/MM/yyyy") + "',103) " + " and trxn_datetime<=convert(datetime,'" + batchScheduleReader.CurrentReconciliationBatch.TransactionStartDate.ToString("dd/MM/yyyy") + " 23:59:59',103) and atm.ATM_id = " + batchScheduleReader.CurrentReconciliationBatch.AtmId +
                          " and transaction_type_id in (" + depositTransactionType + ")"; // Just fethcing Deposit trxns. 

                        var dtBnaEJ = PopulateDataTable(cmd);

                        LogableTask.LogMonoActivityTask("Datatables Population", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "All 3 Deposit Datatables populated from Database.");

                        LogableTask.LogMonoActivityTask("Reconcile() Calling", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "Calling Reconcile() Method for Deposit Transactions.");

                        isBna = true;
                        var resultBna = Reconcile(dtBnaEJ, dtBnaSwitch, dtBnaHost, currentBatchId);

                        LogableTask.LogMonoActivityTask("Reconcile() Executed", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "Reconcile() Method has executed for Deposit Transactions.");

                        if (resultBna.ReconciliationTransactions.Count > 0)
                        {
                            // Amounts Disputed
                            reconBatchInfo.DisputedAmountBNA = resultBna.Summary.TotalAmountDisputed;
                            reconBatchInfo.DisputedBNAEJAmount = resultBna.Summary.AmountDisputedOnEj;
                            reconBatchInfo.DisputedBNASwitchAmount = resultBna.Summary.AmountDisputedOnSwitch;
                            reconBatchInfo.DisputedBNAHostAmount = resultBna.Summary.AmountDisputedOnHost;

                            // Transactions Disputed
                            reconBatchInfo.DisputedBNATransactions = resultBna.Summary.TotalTransDisputed;
                            reconBatchInfo.DisputedBNAEJTransactions = resultBna.Summary.TransDisputedOnEj;
                            reconBatchInfo.DisputedBNASwitchTransactions = resultBna.Summary.TransDisputedOnSwitch;
                            reconBatchInfo.DisputedBNAHostTransactions = resultBna.Summary.TransDisputedOnHost;

                            batchScheduleReader.CurrentReconciliationBatch.AutoReconciledAmount =
                                (batchScheduleReader.CurrentReconciliationBatch.AutoReconciledAmount ?? 0)
                                + resultBna.Summary.TotalAmountReconciled;

                            batchScheduleReader.CurrentReconciliationBatch.NoOfRecordsProcessed =
                                (batchScheduleReader.CurrentReconciliationBatch.NoOfRecordsProcessed ?? 0)
                                + resultBna.Summary.NoOfRecordsProcessed;

                            batchScheduleReader.CurrentReconciliationBatch.NoOfRecordsReconciled =
                                (batchScheduleReader.CurrentReconciliationBatch.NoOfRecordsReconciled ?? 0)
                                + resultBna.Summary.NoOfRecordsReconciledCompletely;

                            batchScheduleReader.CurrentReconciliationBatch.NoOfRecordsFailedToReconciled =
                                (batchScheduleReader.CurrentReconciliationBatch.NoOfRecordsFailedToReconciled ?? 0)
                                + resultBna.Summary.NoOfRecordsFailedToReconciled;

                            batchScheduleReader.CurrentReconciliationBatch.FailureReason = "";

                            batchScheduleReader.CurrentReconciliationBatch.Status = batchScheduleReader.CurrentReconciliationBatch.Status == BatchStatus.DispenseCompleted ? BatchStatus.Completed : BatchStatus.DepositCompleted;

                        }
                        else
                        {
                            LogableTask.LogMonoActivityTask("DoReconcile", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "Reconcile() Method didnot produce any Reconciliation Data for DEPOSIT.");
                        }

                        #endregion

                        var totalReconciliationTransactions = new List<ReconciliationTransactions>();
                        totalReconciliationTransactions.AddRange(result.ReconciliationTransactions);
                        totalReconciliationTransactions.AddRange(resultBna.ReconciliationTransactions);

                        var totalReconciledTransactions = new List<ReconciledTransactions>();
                        totalReconciledTransactions.AddRange(result.ReconciledTransactions);
                        totalReconciledTransactions.AddRange(resultBna.ReconciledTransactions);

                        ManageTx(totalReconciliationTransactions, totalReconciledTransactions, batchScheduleReader, reconBatchInfo);

                    }
                    catch (Exception ex)
                    {
                        task.Log(MethodBase.GetCurrentMethod(), TraceLevel.Error, ex);
                        if (ex.InnerException != null)
                        {
                            task.Log(MethodBase.GetCurrentMethod(),
                                TraceLevel.Error,
                                "Inner: " + ex.InnerException);
                        }
                        if (batchScheduleReader != null)
                        {
                            if (ex.Message.Length > 500)
                                batchScheduleReader.CurrentReconciliationBatch.FailureReason = ex.Message.Substring(0, 499);
                            else
                                batchScheduleReader.CurrentReconciliationBatch.FailureReason = ex.Message;
                            batchScheduleReader.CurrentReconciliationBatch.FailureReason = batchScheduleReader.CurrentReconciliationBatch.FailureReason.Replace("'", "''");
                            if (batchScheduleReader.CurrentReconciliationBatch.RetryCount == 0)
                                batchScheduleReader.CurrentReconciliationBatch.Status = BatchStatus.Failed;
                            else
                                batchScheduleReader.CurrentReconciliationBatch.Status = BatchStatus.Scheduled;

                            batchScheduleReader.CurrentReconciliationBatch.Save();

                        }
                    }

                }
            }
            catch (Exception ex)
            {
                task.Log(MethodBase.GetCurrentMethod(), TraceLevel.Error, ex);
                if (ex.InnerException != null)
                {
                    task.Log(MethodBase.GetCurrentMethod(),
                        TraceLevel.Error,
                        "Inner: " + ex.InnerException);
                }
                if (batchScheduleReader != null)
                {
                    if (ex.Message.Length > 500)
                        batchScheduleReader.CurrentReconciliationBatch.FailureReason = ex.Message.Substring(0, 499);
                    else
                        batchScheduleReader.CurrentReconciliationBatch.FailureReason = ex.Message;
                    batchScheduleReader.CurrentReconciliationBatch.FailureReason = batchScheduleReader.CurrentReconciliationBatch.FailureReason.Replace("'", "''");
                    if (batchScheduleReader.CurrentReconciliationBatch.RetryCount == 0)
                        batchScheduleReader.CurrentReconciliationBatch.Status = BatchStatus.Failed;
                    else
                        batchScheduleReader.CurrentReconciliationBatch.Status = BatchStatus.Scheduled;

                    batchScheduleReader.CurrentReconciliationBatch.Save();

                }
            }
            finally
            {
                if (batchScheduleReader != null)
                    batchScheduleReader.Close();
                if (cmd != null)
                    if (cmd.Connection != null)
                        cmd.Connection.Close();

                task.EndTask();

            }
        }
        private DataTable PopulateDataTable(SqlCommand cmd)
        {
            var dt = new DataTable();
            StringBuilder builder = new StringBuilder();
            stopWatch.Start();
            var adapter = new SqlDataAdapter(cmd);
            builder.Append("q->:" + adapter.SelectCommand.CommandText + "\r\n");
            builder.Append("Start Time:" + stopWatch.Elapsed.ToString() + ",");
            adapter.Fill(dt);
            stopWatch.Stop();
            builder.Append("End Time:" + stopWatch.Elapsed.ToString());
            LogableTask.LogMonoActivityTask("", MethodBase.GetCurrentMethod(), TraceLevel.Info, builder.ToString());

            return dt;
        }
        private ReconciledTransactions CreateReconciledTrxnEntry(int batchID, decimal amount, string title, string tsn, DateTime trxnDatetime, string pan, string source)
        {
            ReconciledTransactions reconciledTrxn = new ReconciledTransactions();
            reconciledTrxn.GeneratedAt = DateTime.Now;
            reconciledTrxn.BatchId = batchID;
            reconciledTrxn.Amount = amount;
            reconciledTrxn.Title = title;
            reconciledTrxn.Tsn = tsn;
            reconciledTrxn.TrxnDatetime = trxnDatetime;
            reconciledTrxn.Pan = pan;
            reconciledTrxn.Source = source;
            //reconciledTrxn.Save(trxn.Connection, trxn);
            return reconciledTrxn;
        }
        public static void ManageTx(List<ReconciliationTransactions> listReconciliationTransactions,
            List<ReconciledTransactions> listReconciledTransactions, ReconciliationBatch.ReconciliationBatchReader reconBatch, ReconBatchInfo reconBatchInfo)
        {
            SqlCommand cmd = null;
            SqlTransaction SqlTrxn = null;

            try
            {
                cmd = ConnectionFactory.GetNewCommand(true);
                SqlTrxn = cmd.Connection.BeginTransaction();
                cmd.Transaction = SqlTrxn;
                if (listReconciliationTransactions.Count > 0)
                    ReconciliationTransactions.BulkSave(listReconciliationTransactions, SqlTrxn);

                if (listReconciledTransactions.Count > 0)
                    ReconciledTransactions.BulkSave(listReconciledTransactions, SqlTrxn);

                reconBatch.CurrentReconciliationBatch.Save(SqlTrxn.Connection, SqlTrxn);
                reconBatchInfo.Save(SqlTrxn.Connection, SqlTrxn);
                SqlTrxn.Commit();

                LogableTask.LogMonoActivityTask("Saving Transactions ", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "Successfully Reconciled and results saved to database.");
            }
            catch (Exception ex)
            {
                LogableTask.LogMonoActivityTask("", MethodBase.GetCurrentMethod(), TraceLevel.Error, ex);
                if (SqlTrxn != null)
                    SqlTrxn.Rollback();

            }
            finally
            {
                if (cmd != null)
                    if (cmd.Connection != null)
                        cmd.Connection.Close();
            }
        }

        private decimal GetAcceptableDiffValue(ReconciliationBatch batch, decimal val1, decimal val2)
        {
            decimal temp = 0;
            if (batch.AcceptableDifferenceType == "Fixed" || batch.AcceptableDifferenceType == "fixed")
                temp = Math.Abs(val1 - val2);
            else
            {
                if (val1 > val2)
                    temp = val2 / val1 * 100;
                else
                    temp = val1 / val2 * 100;
            }
            return temp;
        }

        public static void ManageTx(List<ReconciledTransactions> list)
        {
            SqlCommand cmd = null;
            SqlTransaction SqlTrxn = null;

            try
            {
                cmd = ConnectionFactory.GetNewCommand(true);
                SqlTrxn = cmd.Connection.BeginTransaction();
                cmd.Transaction = SqlTrxn;
                ReconciledTransactions.BulkSave(list, SqlTrxn);
                SqlTrxn.Commit();
            }
            catch (Exception ex)
            {
                LogableTask.LogMonoActivityTask("", MethodBase.GetCurrentMethod(), TraceLevel.Error, ex);
                if (SqlTrxn != null)
                    SqlTrxn.Rollback();
            }
            finally
            {
                if (cmd != null)
                    if (cmd.Connection != null)
                        cmd.Connection.Close();
            }
        }

        public static void ManageTx(List<ReconciliationTransactions> list)
        {
            SqlCommand cmd = null;
            SqlTransaction SqlTrxn = null;

            try
            {
                cmd = ConnectionFactory.GetNewCommand(true);
                SqlTrxn = cmd.Connection.BeginTransaction();
                cmd.Transaction = SqlTrxn;
                ReconciliationTransactions.BulkSave(list, SqlTrxn);
                SqlTrxn.Commit();
            }
            catch (Exception ex)
            {
                LogableTask.LogMonoActivityTask("", MethodBase.GetCurrentMethod(), TraceLevel.Error, ex);
                if (SqlTrxn != null)
                    SqlTrxn.Rollback();
            }
            finally
            {
                if (cmd != null)
                    if (cmd.Connection != null)
                        cmd.Connection.Close();
            }
        }




        protected override void OnStop()
        {
        }

        #region SAAD's RECONCILIATION VERSION

        public ReconciliationResult Reconcile(DataTable dtEj, DataTable dtSwitch, DataTable dtHost, int batchId)
        {
            var (ejLookup, switchLookup, hostLookup) = BuildLookups(dtEj, dtSwitch, dtHost);

            var allKeys = new HashSet<string>(ejLookup.Keys.Union(switchLookup.Keys).Union(hostLookup.Keys));

            var results = new List<ReconciliationTransactions>();
            var reconciledResults = new List<ReconciledTransactions>();
            var summary = new ReconciliationSummary();

            foreach (var key in allKeys)
            {
                try
                {
                    var unified = new UnifiedTransaction
                    {
                        Key = key,
                        EjRow = ejLookup.ContainsKey(key) ? ejLookup[key] : null,
                        SwitchRow = switchLookup.ContainsKey(key) ? switchLookup[key] : null,
                        HostRow = hostLookup.ContainsKey(key) ? hostLookup[key] : null
                    };

                    var result = EvaluateTransaction(unified, batchId);

                    decimal ejAmount;
                    if (!isBna)
                    {
                        ejAmount = unified.EjExists ? Convert.ToDecimal(unified.EjRow["amount"]) : 0;
                    }
                    else
                    {
                        ejAmount = unified.EjExists ? Convert.ToDecimal(unified.EjRow["amount_authorized"]) : 0;

                    }
                    decimal swAmount = unified.SwitchExists ? Convert.ToDecimal(unified.SwitchRow["transaction_amount"]) : 0;
                    decimal hostAmount = unified.HostExists ? Convert.ToDecimal(unified.HostRow["transaction_amount"]) : 0;

                    var maxAmount = Math.Max(ejAmount, Math.Max(swAmount, hostAmount));

                    #region REPORTING / AGGREGATION PART

                    var tempSummary = new ReconciliationSummary();
                    tempSummary.NoOfRecordsProcessed++;

                    if (result.IsReconciled == true)
                        tempSummary.NoOfRecordsReconciledCompletely++;
                    else
                        tempSummary.NoOfRecordsFailedToReconciled++;

                    if (result.Status == "Disputed" || result.Status == "Suspected")
                    {
                        tempSummary.TotalTransDisputed++;

                        // Based on reason, increment appropriate bucket
                        if (result.Reason.Contains("EJ"))
                        {
                            tempSummary.TransDisputedOnEj++;
                            tempSummary.AmountDisputedOnEj += maxAmount;
                        }

                        if (result.Reason.Contains("Switch"))
                        {
                            tempSummary.TransDisputedOnSwitch++;
                            tempSummary.AmountDisputedOnSwitch += maxAmount;
                        }

                        if (result.Reason.Contains("Host"))
                        {
                            tempSummary.TransDisputedOnHost++;
                            tempSummary.AmountDisputedOnHost += maxAmount;
                        }

                        if (result.Reason.Contains("All three"))
                        {
                            tempSummary.TransDisputedOnEj++;
                            tempSummary.AmountDisputedOnEj += maxAmount;

                            tempSummary.TransDisputedOnSwitch++;
                            tempSummary.AmountDisputedOnSwitch += maxAmount;

                            tempSummary.TransDisputedOnHost++;
                            tempSummary.AmountDisputedOnHost += maxAmount;
                        }

                        tempSummary.TotalAmountDisputed += maxAmount;
                    }
                    else if (result.Status == "Matched")
                    {
                        tempSummary.TotalAmountReconciled += maxAmount;
                    }

                    #endregion

                    // commiting details at the end so that everything is successful

                    results.Add(result);
                    MergeSummary(summary, tempSummary);

                    if (result.Status == "Matched")
                    {
                        string atmTitle = null;
                        string tsn = null;
                        object trxnDate = null;
                        string cardNo = null;
                        string source = null;

                        if (unified.EjExists)
                        {
                            atmTitle = unified.EjRow["title"]?.ToString();
                            tsn = !isBna ? unified.EjRow["tsn"]?.ToString() : unified.EjRow["seq"]?.ToString();
                            trxnDate = unified.EjRow["trxn_datetime"];
                            cardNo = unified.EjRow["pan"]?.ToString();
                            source = "EJ";
                        }
                        else
                        {
                            atmTitle = unified.HostRow["atm_id"]?.ToString();
                            tsn = unified.HostRow["transaction_sequence"]?.ToString();
                            trxnDate = unified.HostRow["transaction_date"];
                            cardNo = unified.HostRow["card_number"]?.ToString();
                            source = "Switch/Host";
                        }

                        DateTime parsedDate;
                        DateTime.TryParse(trxnDate?.ToString(), out parsedDate);

                        reconciledResults.Add(
                        CreateReconciledTrxnEntry(
                            batchId,
                            maxAmount,
                            atmTitle,
                            tsn,
                            parsedDate,
                            cardNo,
                            source
                        ));
                    }

                }
                catch (Exception ex)
                {
                    LogableTask.LogMonoActivityTask("Evaluating Transactions", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Error, $"Error while evaluating Transaction Key: {key} => " + ex);

                    EventLog.WriteEntry("CCMSReconciler", $"Error in EvaluateTransaction() Transaction Key: {key} =>  " + ex.Message + " " + ex.StackTrace, EventLogEntryType.Error);

                    continue;
                }

            }

            return new ReconciliationResult
            {
                ReconciliationTransactions = results,
                ReconciledTransactions = reconciledResults,
                Summary = summary
            };
        }

        private ReconciliationTransactions EvaluateTransaction(UnifiedTransaction ut, int batchId)
        {
            var record = new ReconciliationTransactions
            {
                ReconciliationBatchId = batchId,
                ComparisonType = !isBna ? 1 : 2,
                GeneratedAt = DateTime.Now,
                UserId = 1
            };

            AttachSourceIds(record, ut);

            int presenceCount =
                (ut.EjExists ? 1 : 0) +
                (ut.SwitchExists ? 1 : 0) +
                (ut.HostExists ? 1 : 0);

            if (presenceCount == 3)
                EvaluateThreeWay(ut, record);
            else if (presenceCount == 2)
                EvaluateTwoWay(ut, record);
            else
                EvaluateSingleSource(ut, record);

            return record;
        }

        private void EvaluateThreeWay(UnifiedTransaction ut, ReconciliationTransactions record)
        {
            decimal ejAmount;
            int ejStatus;
            if (!isBna)
            {
                ejAmount = ut.EjRow["amount"] == DBNull.Value ? 0 : Math.Round(Convert.ToDecimal(ut.EjRow["amount"]), 2);

                ejStatus = Convert.ToInt32(ut.EjRow["status"]) != 0 ? 1 : 0;
            }
            else
            {
                ejAmount = ut.EjRow["amount_authorized"] == DBNull.Value ? 0 : Math.Round(Convert.ToDecimal(ut.EjRow["amount_authorized"]), 2);

                ejStatus = ut.EjRow["status"].ToString() != "Successful" ? 1 : 0;
            }

            decimal swAmount = ut.SwitchRow["transaction_amount"] == DBNull.Value ? 0 : Math.Round(Convert.ToDecimal(ut.SwitchRow["transaction_amount"]), 2);

            decimal hostAmount = ut.HostRow["transaction_amount"] == DBNull.Value ? 0 : Math.Round(Convert.ToDecimal(ut.HostRow["transaction_amount"]), 2);

            int swStatus = ut.SwitchRow["transaction_response"].ToString().ToUpper().Contains("APPROVED") ? 0 : 1;
            int hostStatus = ut.HostExists ? 0 : 1;

            var amountResult = CompareThreeAmounts(ejAmount, swAmount, hostAmount);
            var statusResult = CompareThreeStatuses(ejStatus, swStatus, hostStatus);

            if (amountResult.IsMatch && statusResult.IsMatch)
            {
                record.IsReconciled = true;
                record.Status = "Matched";
            }
            else
            {
                record.IsReconciled = false;
                record.Status = "Disputed";
                record.Reason = $"{amountResult.Reason} {statusResult.Reason}".Trim();


            }
        }

        private void EvaluateTwoWay(UnifiedTransaction ut, ReconciliationTransactions record)
        {
            var sources = new List<(string Name, DataRow Row)>();

            if (ut.EjExists) sources.Add(("EJ", ut.EjRow));
            if (ut.SwitchExists) sources.Add(("Switch", ut.SwitchRow));
            if (ut.HostExists) sources.Add(("Host", ut.HostRow));

            var first = sources[0];
            var second = sources[1];

            decimal amount1, amount2;
            int status1, status2;

            if (!isBna)
            {
                amount1 = first.Name == "EJ" ? Math.Round(Convert.ToDecimal(first.Row["amount"]), 2) : Math.Round(Convert.ToDecimal(first.Row["transaction_amount"]), 2);

                amount2 = second.Name == "EJ" ? Math.Round(Convert.ToDecimal(second.Row["amount"]), 2) : Math.Round(Convert.ToDecimal(second.Row["transaction_amount"]), 2);

                status1 = first.Name == "EJ" ? Convert.ToInt32(first.Row["status"]) != 0 ? 1 : 0 : first.Name == "Switch" ? first.Row["transaction_response"].ToString().ToUpper().Contains("APPROVED") ? 0 : 1 : 0;

                status2 = second.Name == "EJ" ? Convert.ToInt32(second.Row["status"]) != 0 ? 1 : 0 : second.Name == "Switch" ? second.Row["transaction_response"].ToString().ToUpper().Contains("APPROVED") ? 0 : 1 : 0;
            }
            else
            {
                amount1 = first.Name == "EJ" ? Math.Round(Convert.ToDecimal(first.Row["amount_authorized"]), 2) : Math.Round(Convert.ToDecimal(first.Row["transaction_amount"]), 2);

                amount2 = second.Name == "EJ" ? Math.Round(Convert.ToDecimal(second.Row["amount_authorized"]), 2) : Math.Round(Convert.ToDecimal(second.Row["transaction_amount"]), 2);

                status1 = first.Name == "EJ" ? first.Row["status"].ToString() != "Successful" ? 1 : 0 : first.Name == "Switch" ? first.Row["transaction_response"].ToString().ToUpper().Contains("APPROVED") ? 0 : 1 : 0;

                status2 = second.Name == "EJ" ? second.Row["status"].ToString() != "Successful" ? 1 : 0 : second.Name == "Switch" ? second.Row["transaction_response"].ToString().ToUpper().Contains("APPROVED") ? 0 : 1 : 0;
            }


            //bool amountMatch = amount1 == amount2;

            string missingSource = !ut.EjExists ? "EJ" : !ut.SwitchExists ? "Switch" : "Host";

            if (status1 == 1 && status2 == 1) // if status of both present sources is failure
            {
                record.IsReconciled = true;
                record.Status = "Matched";
                record.Reason = $"Missing in {missingSource}";
            }
            else
            {
                record.IsReconciled = false;
                record.Status = "Disputed";
                record.Reason = $"Missing in {missingSource}";
            }
        }

        private void EvaluateSingleSource(UnifiedTransaction ut, ReconciliationTransactions record)
        {
            string source = ut.EjExists ? "EJ" : ut.SwitchExists ? "Switch" : "Host";
            DataRow row = ut.EjExists ? ut.EjRow : ut.SwitchExists ? ut.SwitchRow : ut.HostRow;

            record.IsReconciled = false;
            record.Status = "Suspected";

            record.Reason = source == "EJ" ? "Missing in Switch and Host" :
                            source == "Switch" ? "Missing in EJ and Host" :
                                      "Missing in EJ and Switch";

            

            //int status = ut.EjExists ? Convert.ToInt32(ut.EjRow["status"]) != 0 ? 1 : 0 : ut.SwitchExists ? ut.SwitchRow["transaction_response"].ToString().ToUpper().Contains("APPROVED") ? 0 : 1 : ut.HostExists ? 0 : 1;

            //if (status == 1) // failed
            //{
            //    record.IsReconciled = true;
            //    record.Status = "Matched";
            //    record.Reason = $"Transation only present in {source} but failed";
            //}
            //else
            //{
            //    record.IsReconciled = false;
            //    record.Status = "Disputed";
            //    record.Reason = $"Transaction exists only in {source}";
            //}
        }

        private void AttachSourceIds(ReconciliationTransactions record, UnifiedTransaction ut)
        {
            if (!isBna)
            {
                if (ut.EjRow != null)
                    record.EjParsedTransactionsId = Convert.ToInt32(ut.EjRow["ej_parsed_transactions_id"]);

                if (ut.SwitchRow != null)
                    record.ReconciliationSwitchDataId = Convert.ToInt32(ut.SwitchRow["reconciliation_switch_data_id"]);

                if (ut.HostRow != null)
                    record.ReconciliationHostDataId = Convert.ToInt32(ut.HostRow["reconciliation_host_data_id"]);
            }
            else
            {
                if (ut.EjRow != null)
                    record.EjParsedBnaTransactionId = Convert.ToInt32(ut.EjRow["ej_parsed_bna_transaction_id"]);

                if (ut.SwitchRow != null)
                    record.ReconciliationBnaSwitchDataId = Convert.ToInt32(ut.SwitchRow["reconciliation_bna_switch_data_id"]);

                if (ut.HostRow != null)
                    record.ReconciliationBnaHostDataId = Convert.ToInt32(ut.HostRow["reconciliation_bna_host_data_id"]);
            }

        }
        private void MergeSummary(ReconciliationSummary main, ReconciliationSummary temp)
        {
            main.NoOfRecordsProcessed += temp.NoOfRecordsProcessed;
            main.NoOfRecordsReconciledCompletely += temp.NoOfRecordsReconciledCompletely;
            main.NoOfRecordsFailedToReconciled += temp.NoOfRecordsFailedToReconciled;

            main.TotalTransDisputed += temp.TotalTransDisputed;
            main.TotalAmountDisputed += temp.TotalAmountDisputed;
            main.TotalAmountReconciled += temp.TotalAmountReconciled;

            main.TransDisputedOnEj += temp.TransDisputedOnEj;
            main.AmountDisputedOnEj += temp.AmountDisputedOnEj;

            main.TransDisputedOnSwitch += temp.TransDisputedOnSwitch;
            main.AmountDisputedOnSwitch += temp.AmountDisputedOnSwitch;

            main.TransDisputedOnHost += temp.TransDisputedOnHost;
            main.AmountDisputedOnHost += temp.AmountDisputedOnHost;
        }
        private (bool IsMatch, string Reason) CompareThreeAmounts(decimal a, decimal b, decimal c)
        {
            // a is EJ
            // b is Switch
            // c is Host
            bool ab = a == b;
            bool ac = a == c;
            bool bc = b == c;

            if (ab && ac)
                return (true, "");

            if (ab && !ac)
                return (false, "Host amount mismatch.");

            if (ac && !ab)
                return (false, "Switch amount mismatch.");

            if (bc && !ab)
                return (false, "EJ amount mismatch.");

            return (false, "All three amounts differ.");
        }

        private (bool IsMatch, string Reason) CompareThreeStatuses(int a, int b, int c)
        {
            // a is EJ
            // b is Switch
            // c is Host
            if (a == b && b == c)
                return (true, "");

            if (a == b && a != c)
                return (false, "Host status mismatch.");

            if (a == c && a != b)
                return (false, "Switch status mismatch.");

            if (b == c && a != b)
                return (false, "EJ status mismatch.");

            return (false, "All three statuses differ.");
        }

        private string BuildKey(string atm, string pan, string tsn)
        {
            return $"{atm.Trim()}|{pan.Trim()}|{tsn.Trim()}";
        }

        private (Dictionary<string, DataRow>, Dictionary<string, DataRow>, Dictionary<string, DataRow>) BuildLookups(DataTable dtEj, DataTable dtSwitch, DataTable dtHost)
        {
            var ejLookup = new Dictionary<string, DataRow>();
            var duplicateTsn = new StringBuilder();
            var duplicateSet = new HashSet<string>();

            foreach (DataRow ejRow in dtEj.Rows)
            {
                string tsn = !isBna ? ejRow["tsn"].ToString() : ejRow["seq"].ToString();

                string key = BuildKey(
                    ejRow["title"].ToString(),
                    ejRow["pan"].ToString(),
                    tsn);

                if (!ejLookup.ContainsKey(key))
                {
                    ejLookup.Add(key, ejRow);

                }
                else
                {
                    if (duplicateSet.Add(tsn)) // ensures tsn added only once
                        duplicateTsn.Append(", ").Append(tsn);
                }
            }

            if (duplicateTsn.Length > 0)
            {
                duplicateTsn.Remove(0, 2); // remove leading comma + space

                LogableTask.LogMonoActivityTask("Duplicates in EJ", MethodBase.GetCurrentMethod(), TraceLevel.Info, $"Duplicate transaction sequence in EJ for ATM: {dtEj.Rows[0]["title"]}:" + Environment.NewLine + duplicateTsn + Environment.NewLine);
            }

            var switchLookup = new Dictionary<string, DataRow>();
            duplicateTsn = new StringBuilder();
            duplicateSet = new HashSet<string>();

            foreach (DataRow switchRow in dtSwitch.Rows)
            {
                string cardNoSwitch = switchRow["card_number"].ToString();
                string cardPrefixSixDig = cardNoSwitch.Substring(0, 6);
                string cardSuffixFourDig = cardNoSwitch.Substring(cardNoSwitch.Length - 4, 4);
                string cardNoMasked = cardPrefixSixDig + "******" + cardSuffixFourDig;

                string tsn = switchRow["transaction_sequence"].ToString();
                string key = BuildKey(
                    switchRow["atm_id"].ToString(),
                    cardNoMasked,
                    tsn);

                if (!switchLookup.ContainsKey(key))
                {
                    switchLookup.Add(key, switchRow);
                }
                else
                {
                    if (duplicateSet.Add(tsn)) // ensures tsn added only once
                        duplicateTsn.Append(", ").Append(tsn);
                }
            }

            if (duplicateTsn.Length > 0)
            {
                duplicateTsn.Remove(0, 2); // remove leading comma + space

                LogableTask.LogMonoActivityTask("Duplicates in Switch", MethodBase.GetCurrentMethod(), TraceLevel.Info, $"Duplicate transaction sequence in Switch for ATM: {dtSwitch.Rows[0]["atm_id"]}:" + Environment.NewLine + duplicateTsn + Environment.NewLine);
            }

            var hostLookup = new Dictionary<string, DataRow>();
            duplicateTsn = new StringBuilder();
            duplicateSet = new HashSet<string>();

            foreach (DataRow hostRow in dtHost.Rows)
            {
                string tsn = hostRow["transaction_sequence"].ToString();

                string key = BuildKey(
                    hostRow["atm_id"].ToString(),
                    hostRow["card_number"].ToString(),
                    tsn);

                if (!hostLookup.ContainsKey(key))
                {
                    hostLookup.Add(key, hostRow);
                }
                else
                {
                    if (duplicateSet.Add(tsn)) // ensures tsn added only once
                        duplicateTsn.Append(", ").Append(tsn);
                }
            }

            if (duplicateTsn.Length > 0)
            {
                duplicateTsn.Remove(0, 2); // remove leading comma + space

                LogableTask.LogMonoActivityTask("Duplicates in Host", MethodBase.GetCurrentMethod(), TraceLevel.Info, $"Duplicate transaction sequence in Host for ATM: {dtHost.Rows[0]["atm_id"]}:" + Environment.NewLine + duplicateTsn + Environment.NewLine);
            }

            return (ejLookup, switchLookup, hostLookup);
        }

        #endregion

    }

    public class UnifiedTransaction
    {
        public string Key { get; set; }
        public DataRow EjRow { get; set; }
        public DataRow SwitchRow { get; set; }
        public DataRow HostRow { get; set; }

        public bool EjExists => EjRow != null;
        public bool SwitchExists => SwitchRow != null;
        public bool HostExists => HostRow != null;
    }
    public class ReconciliationSummary
    {
        public int NoOfRecordsProcessed { get; set; }
        public int NoOfRecordsFailedToReconciled { get; set; }
        public int NoOfRecordsReconciledCompletely { get; set; }
        public decimal TotalAmountReconciled { get; set; }
        public int TotalTransDisputed { get; set; }
        public decimal TotalAmountDisputed { get; set; }
        public decimal AmountDisputedOnEj { get; set; }
        public decimal AmountDisputedOnSwitch { get; set; }
        public decimal AmountDisputedOnHost { get; set; }
        public int TransDisputedOnEj { get; set; }
        public int TransDisputedOnSwitch { get; set; }
        public int TransDisputedOnHost { get; set; }
    }
    public class ReconciliationResult
    {
        public List<ReconciliationTransactions> ReconciliationTransactions { get; set; }
        public List<ReconciledTransactions> ReconciledTransactions { get; set; }
        public ReconciliationSummary Summary { get; set; }
    }
    public static class BatchStatus
    {
        public const string Scheduled = "Scheduled";
        public const string DispenseCompleted = "Dispense Completed";
        public const string DepositCompleted = "Deposit Completed";
        public const string Completed = "Completed";
        public const string Failed = "Failed";
    }

}
