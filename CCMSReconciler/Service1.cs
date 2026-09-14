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
        //int noOfRecordsProcessed = 0;
        //int noOfRecordsReconciled = 0;
        //decimal amountReconciledCompletely = 0;
        //int noOfRecordsReconciledCompletely = 0;
        //int noOfRecordsFailedToReconciled = 0;

        //int withdrawalTransactionType = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["withdrawalTransactionType"]);
        string withdrawalTransactionType = System.Configuration.ConfigurationManager.AppSettings["withdrawalTransactionType"];
        int defaultInterval = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["RescheduleReconciler"]); //minutes

        string OnUsCode;// = System.Configuration.ConfigurationManager.AppSettings["OnUsCode"];
        //string reversalCodes = System.Configuration.ConfigurationManager.AppSettings["reversalCodes"];
        string CardLessCode = System.Configuration.ConfigurationManager.AppSettings["CardLessCode"];
        string CardLessCardnoinHOST = System.Configuration.ConfigurationManager.AppSettings["CardLessCardno_Host"];
        string CardLessCardno_CivilID = System.Configuration.ConfigurationManager.AppSettings["CardLessCardno_CivilID"];
        string TellerxCardCode = System.Configuration.ConfigurationManager.AppSettings["TellerxCardCode"];
        string CreditCardCode;// = System.Configuration.ConfigurationManager.AppSettings["CreditCardCode"]; //ignore credit card transactions as they are settled outside view360
        string Cardless_Tellerx_TerminalID = System.Configuration.ConfigurationManager.AppSettings["Cardless_Tellerx_TerminalID"];
        int Cardless_Tellerx_BatchID = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["Cardless_Tellerx_BatchID"]);
        int RemoteOnUs_BatchID = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["RemoteOnUs_BatchID"]);
        int CancelScheduled = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["CancelScheduled"]);

        // string KNET_ST_Approved_ResponseCodes = System.Configuration.ConfigurationManager.AppSettings["KNET_ST_Approved_ResponseCodes"]; //add hardcode response for successful host-local-gcc totals
        string KNETdrvnATMSAG = "'002' ,'003' ,'005' ,'006' ,'007' ,'009' ,'011' ,'012' ,'013' ,'015' ,'016' ,'017' ,'018' ,'019' ,'020' ,'021' ,'022' ,'023' ,'027' ,'029' ,'032' ,'034' ,'035' ,'037' ,'038' ,'039' ,'040' ,'042' ,'043' ,'044' ,'045' ,'046' ,'047' ,'048' ,'049' ,'050' ,'051' ,'052' ,'054' ,'055' ,'056' ,'057' ,'058' ,'059' ,'060' ,'061' ,'062' ,'063' ,'064' ,'065' ,'066' ,'067' ,'068' ,'070' ,'071' ,'072' ,'073' ,'076' ,'077' ,'079' ,'080' ,'081' ,'082' ,'083' ,'084' ,'085' ,'086' ,'087' ,'088' ,'089' ,'090' ,'091' ,'092' ,'093' ,'095' ,'097' ,'098' ,'099' ,'100' ,'101' ,'102' ,'103' ,'105' ,'106' ,'108' ,'109' ,'110' ,'111' ,'112' ,'113' ,'114' ,'116' ,'117' ,'120' ,'122' ,'123' ,'124' ,'126' ,'127' ,'128' ,'129' ,'130' ,'131' ,'132' ,'134' ,'135' ,'136' ,'137' ,'138' ,'139' ,'140' ,'141' ,'142' ,'143' ,'144' ,'145' ,'146' ,'147' ,'148' ,'149' ,'150' ,'151' ,'152' ,'153' ,'154' ,'155' ,'156' ,'157' ,'158' ,'159' ,'160' ,'161' ,'162' ,'163' ,'164' ,'165'";
        string KNETdrvnATMSA = "'702' ,'715' ,'716' ,'717' ,'727' ,'735' ,'741' ,'743' ,'747' ,'748' ,'749' ,'751' ,'753' ,'755' ,'758' ,'760' ,'762' ,'767' ,'768' ,'770' ,'771' ,'772' ,'773' ,'776' ,'779' ,'782' ,'783' ,'784' ,'790' ,'791'";
        //bool _isFilesinUSE = false;
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
                batchScheduleReader = ReconciliationBatch.ExecuteReader("status = 'Scheduled' and retry_count>0 and atm_id in (select atm_id from atm where is_active=1) order by atm_id desc"); //to process REMOTE BATCH ALWAYS AT LAST
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

                        cmd.CommandText = "select * from reconciliation_host_data where reconciliation_batch_id =  " + currentBatchId; //TODO: Need transaction_type for transactions from bank to filter here.

                        var dtHost = PopulateDataTable(cmd);

                        cmd.CommandText = "select * from reconciliation_switch_data where reconciliation_batch_id = "
                            + currentBatchId + " and transaction_type like '%Withdrawal%'";

                        var dtSwitch = PopulateDataTable(cmd);

                        cmd.CommandText = "select * from vejparsedtransactions ej_parsed_transactions inner join atm on ej_parsed_transactions.atm_id = atm.atm_id and amount is not null and trxn_Datetime >= convert(datetime,'" + batchScheduleReader.CurrentReconciliationBatch.TransactionStartDate.ToString("dd/MM/yyyy") + "',103) " + " and trxn_datetime<=convert(datetime,'" + batchScheduleReader.CurrentReconciliationBatch.TransactionStartDate.ToString("dd/MM/yyyy") + " 23:59:59',103) and atm.ATM_id = " + batchScheduleReader.CurrentReconciliationBatch.AtmId +
                          " and transaction_type_id in (38035)"; // Just fethcing Withdrawal trxns. 38035

                        var dtEJ = PopulateDataTable(cmd);

                        LogableTask.LogMonoActivityTask("Datatables Population", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "All 3 Datatables populated from Database.");

                        ReconBatchInfo.DeleteReconBatchInfos("Reconciliation_Batch_Id = " + batchScheduleReader.CurrentReconciliationBatch.ReconciliationBatchId);

                        var reconBatchInfo = new ReconBatchInfo();
                        reconBatchInfo.ReconciliationBatchId = currentBatchId;
                        reconBatchInfo.AtmId = batchScheduleReader.CurrentReconciliationBatch.AtmId;
                        reconBatchInfo.TranDate = batchScheduleReader.CurrentReconciliationBatch.TransactionStartDate;
                        reconBatchInfo.ProcessingDateTime = DateTime.Now;


                        var result = Reconcile(dtEJ, dtSwitch, dtHost, currentBatchId);

                        LogableTask.LogMonoActivityTask("Reconcile() Executed", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "Reconcile() Method has executed completely.");

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
                        batchScheduleReader.CurrentReconciliationBatch.Status = "Completed";
                        batchScheduleReader.CurrentReconciliationBatch.FailureReason = "";


                        ManageTx(result.ReconciliationTransactions, result.ReconciledTransactions, batchScheduleReader, reconBatchInfo);
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
                                batchScheduleReader.CurrentReconciliationBatch.Status = "Failed";
                            else
                                batchScheduleReader.CurrentReconciliationBatch.Status = "Scheduled";

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
                        batchScheduleReader.CurrentReconciliationBatch.Status = "Failed";
                    else
                        batchScheduleReader.CurrentReconciliationBatch.Status = "Scheduled";

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

                    decimal ejAmount = unified.EjExists ? Convert.ToDecimal(unified.EjRow["amount"]) : 0;
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

                    if (result.Status == "Disputed")
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

                    if(result.Status == "Matched")
                    {
                        string atmTitle = null;
                        string tsn = null;
                        object trxnDate = null;
                        string cardNo = null;
                        string source = null;

                        if (unified.EjExists)
                        {
                            atmTitle = unified.EjRow["title"]?.ToString();
                            tsn = unified.EjRow["tsn"]?.ToString();
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
                ComparisonType = 1, // Indicates Transaction Type: Withdrawal
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
            decimal ejAmount = ut.EjRow["amount"] == DBNull.Value ? 0 : Math.Round(Convert.ToDecimal(ut.EjRow["amount"]), 2);

            decimal swAmount = ut.SwitchRow["transaction_amount"] == DBNull.Value ? 0 : Math.Round(Convert.ToDecimal(ut.SwitchRow["transaction_amount"]), 2);

            decimal hostAmount = ut.HostRow["transaction_amount"] == DBNull.Value ? 0 : Math.Round(Convert.ToDecimal(ut.HostRow["transaction_amount"]), 2);

            int ejStatus = Convert.ToInt32(ut.EjRow["status"]) != 0 ? 1 : 0;
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

            decimal amount1 = first.Name == "EJ" ? Math.Round(Convert.ToDecimal(first.Row["amount"]), 2) : Math.Round(Convert.ToDecimal(first.Row["transaction_amount"]), 2);

            decimal amount2 = second.Name == "EJ" ? Math.Round(Convert.ToDecimal(second.Row["amount"]), 2) : Math.Round(Convert.ToDecimal(second.Row["transaction_amount"]), 2);

            int status1 = first.Name == "EJ" ? Convert.ToInt32(first.Row["status"]) != 0 ? 1 : 0 : first.Name == "Switch" ? first.Row["transaction_response"].ToString().ToUpper().Contains("APPROVED") ? 0 : 1 : 0;

            int status2 = second.Name == "EJ" ? Convert.ToInt32(second.Row["status"]) != 0 ? 1 : 0 : second.Name == "Switch" ? second.Row["transaction_response"].ToString().ToUpper().Contains("APPROVED") ? 0 : 1 : 0;

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
            record.Reason = $"Transaction found in only one source: {source}";

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
            if (ut.EjRow != null)
                record.EjParsedTransactionsId = Convert.ToInt32(ut.EjRow["ej_parsed_transactions_id"]);

            if (ut.SwitchRow != null)
                record.ReconciliationSwitchDataId = Convert.ToInt32(ut.SwitchRow["reconciliation_switch_data_id"]);

            if (ut.HostRow != null)
                record.ReconciliationHostDataId = Convert.ToInt32(ut.HostRow["reconciliation_host_data_id"]);
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
                string tsn = ejRow["tsn"].ToString();

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

    #region OLD REONCILIATION IMPLEMENTATION
    //        private void CompareEjAndSwitchData(DataTable dtEj, DataTable dtSwitch, int batchID, ref decimal amountDisputed, ref decimal amountDisputedOffUs, ref decimal amountDisputedOnUsSwitch,
    //            ref decimal amountDisputedOnUsEJ, ref decimal amountDisputedOffUsSwitch, ref decimal amountDisputedOffUsEJ, ref decimal amountDisputedCardless, ref decimal amountDisputedCardlessSwitch,
    //            ref decimal amountDisputedCardlessEJ, ref decimal transDisputed, ref decimal transDisputedoffus, ref decimal transDisputedcardless, ref decimal transDisputedonusSwitch, ref decimal transDisputedoffusSwitch, ref decimal transDisputedcardlessSwitch, ref decimal transDisputedonusEJ, ref decimal transDisputedoffusEJ, ref decimal transDisputedcardlessEJ,
    //             ref decimal transDisputedOFFUS_KNET, ref decimal amountDisputedOFFUS_KNET, ref decimal transDisputedOFFUS_GCC, ref decimal amountDisputedOFFUS_GCC)
    //        {
    //            List<int> switchIdProcessed = new List<int>();
    //            ReconciliationBatch batch = ReconciliationBatch.LoadReconciliationBatchByPk(batchID);
    //            foreach (DataRow dr in dtEj.Rows)
    //            {
    //                //if (!OnUsCode.Contains(dr["pan"].ToString().Substring(0, 6)))
    //                //{
    //                //    continue;
    //                //}

    //                // Exclude Internationl OFF US cards
    //                //if (!
    //                //    (isLocalBankPAN.Contains(dr["pan"].ToString().Substring(0, 6))
    //                //    || isGCCBankPAN.Contains(dr["pan"].ToString().Substring(0, 6))
    //                //    || isGBKBankPAN.Contains(dr["pan"].ToString().Substring(0, 6))
    //                //    || dr["pan"].ToString().Substring(0, 6) == "999999"
    //                //    || isAMEXBankPAN.Contains(dr["pan"].ToString().Substring(0, 2)))
    //                //    )
    //                //{
    //                //    continue;
    //                //}

    //                bool saveTrn = true;
    //                noOfRecordsProcessed++;
    //                ReconciliationTransactions trxn = new ReconciliationTransactions();
    //                trxn.ReconciliationBatchId = batchID;
    //                trxn.EjParsedTransactionsId = int.Parse(dr["ej_parsed_transactions_id"].ToString());

    //                DateTime trxnDatetime = DateTime.Parse(dr["trxn_datetime"].ToString());
    //                string date = trxnDatetime.ToString("yyyyMMdd");
    //                string time = trxnDatetime.ToString("HH:mm");

    //                //            LogableTask.LogMonoActivityTask("match", MethodBase.GetCurrentMethod(), TraceLevel.Info,
    //                //" Going to check switch data with this criteria " +
    //                //string.Format("atm_id = '{0}' and card_number like '%{1}' and transaction_date='{2}' and transaction_time = '{3}' and transaction_sequence like '%{4}'",
    //                //               dr["title"], dr["pan"].ToString().Substring(dr["pan"].ToString().Length - 4), date, time, dr["tsn"]));

    //                //DataRow[] drArray = dtSwitch.Select(
    //                //    string.Format("atm_id = '{0}' and card_number like '%{1}' and transaction_date='{2}' and transaction_sequence like '%{3}'",
    //                //                   dr["title"], dr["pan"].ToString(), date,  dr["tsn"]));

    //                DataRow[] drArray = dtSwitch.Select(
    //                    string.Format("atm_id = '{0}' and card_number ='{1}' and transaction_sequence like '%{2}'",
    //                                   dr["title"], dr["pan"].ToString(), dr["tsn"]));

    //                if (dr["pan"].ToString() == "507808******5224")
    //                {
    //                    int c = 0;
    //                }
    //                if (dr["tsn"].ToString().Contains("2612"))
    //                {
    //                    int c = 0;
    //                }
    //                if (drArray.Length > 0) //Transaction found in switch.
    //                {
    //                    // int ejResponseCode = Convert.ToInt32(dr["result"].ToString());
    //                    int switchResponseCode = Convert.ToInt32(drArray[0]["transaction_response"].ToString());
    //                    int status = Convert.ToInt32(dr["status"].ToString());
    //                    bool check = true;
    //                    //if (ejResponseCode == 0 && switchResponseCode == 0 && status != 0)
    //                    //{
    //                    //    check = false;
    //                    //}

    //                    if (status != 0)
    //                        status = 1;

    //                    if (switchResponseCode != 0)
    //                        switchResponseCode = 1;

    //                    trxn.ReconciliationSwitchDataId = int.Parse(drArray[0]["reconciliation_switch_data_id"].ToString());
    //                    switchIdProcessed.Add(trxn.ReconciliationSwitchDataId.Value);
    //                    decimal ejAmount = decimal.Parse(dr["amount"].ToString());
    //                    decimal switchAmount = decimal.Parse(drArray[0]["transaction_amount"].ToString());
    //                    //string ejStatus = dr["status"].ToString();
    //                    //string switchStatus = drArray[0]["transaction_status"].ToString();

    //                    decimal temp = GetAcceptableDiffValue(batch, switchAmount, ejAmount);


    //                    //                    if (((switchAmount == ejAmount) || (temp <= batch.AcceptableDifference && temp != 0)) & (ejResponseCode == switchResponseCode) & check)

    //                    if (((switchAmount == ejAmount) || (temp <= batch.AcceptableDifference && temp != 0)) & (status == switchResponseCode))
    //                    {
    //                        if (temp <= batch.AcceptableDifference && temp != 0)
    //                            trxn.Reason = "Transaction Amount Mismatched On Switch But reconciled because of acceptable difference";

    //                        if (listReconciledSwitchTrxn.Contains(trxn.ReconciliationSwitchDataId.Value))
    //                        {
    //                            trxn.IsReconciled = true;
    //                            trxn.Status = "Matched";
    //                            noOfRecordsReconciled++;
    //                            noOfRecordsReconciledCompletely++;

    //                            ReconciliationTransactions reconTran = listReconciliationTransactions.Where(i => i.ReconciliationSwitchDataId == trxn.ReconciliationSwitchDataId.Value).SingleOrDefault();

    //                            //var reconTran = ReconciliationTransactions.LoadReconciliationTransactions("reconciliation_switch_data_id = " + trxn.ReconciliationSwitchDataId.Value);

    //                            reconTran.EjParsedTransactionsId = int.Parse(dr["ej_parsed_transactions_id"].ToString());
    //                            reconTran.ComparisonType = 0;
    //                            //reconTran.Save(dbTrxn.Connection, dbTrxn);
    //                            saveTrn = false;
    //                            //amountReconciledCompletely += ejAmount;
    //                            //CreateGLAccountSummary(trxn.ReconciliationBatchId.Value, int.Parse(dr["atm_id"].ToString()), ejAmount, dbTrxn);
    //                            listReconciledTransactions.Add(CreateReconciledTrxnEntry(trxn.ReconciliationBatchId.Value, ejAmount, dr["title"].ToString(), dr["tsn"].ToString(), trxnDatetime,
    //                                dr["pan"].ToString(), "EJ"));

    //                        }
    //                        else if (listDisputedSwitchTrxn.Contains(trxn.ReconciliationSwitchDataId.Value))
    //                        {
    //                            List<ReconciliationTransactions> reconTran = listReconciliationTransactions.Where(i => i.ReconciliationSwitchDataId == trxn.ReconciliationSwitchDataId.Value).ToList();//.SingleOrDefault();
    //                            for (int k = 0; k < reconTran.Count; k++)
    //                                reconTran[k].EjParsedTransactionsId = int.Parse(dr["ej_parsed_transactions_id"].ToString());
    //                            //var reconTran = ReconciliationTransactions.LoadReconciliationTransactions("reconciliation_switch_data_id = " + trxn.ReconciliationSwitchDataId.Value);
    //                            //reconTran.EjParsedTransactionsId = int.Parse(dr["ej_parsed_transactions_id"].ToString());
    //                            //reconTran.ComparisonType = 2;
    //                            //reconTran.Status = "Disputed";
    //                            //reconTran.Reason = "Transaction matched but status is failed on EJ";
    //                            //reconTran.Save(dbTrxn.Connection, dbTrxn);
    //                            saveTrn = false;
    //                        }
    //                        else
    //                        {
    //                            trxn.IsReconciled = true;
    //                            trxn.Status = "Matched";
    //                            noOfRecordsReconciled++;
    //                        }
    //                    }
    //                    else if (((switchAmount == ejAmount) || (temp <= batch.AcceptableDifference && temp != 0)) & (status == switchResponseCode) & !check)
    //                    {
    //                        trxn.Reason = "Transaction matched but status is failed on EJ.";
    //                        trxn.Status = "Disputed";
    //                        trxn.IsReconciled = false;
    //                        noOfRecordsFailedToReconciled++;

    //                        if (CardLessCode.Contains(dr["pan"].ToString()))
    //                        {
    //                            amountDisputedCardless += ejAmount;
    //                            amountDisputedCardlessEJ += ejAmount;

    //                            transDisputedcardless++;
    //                            transDisputedcardlessEJ++;
    //                        }
    //                        else if (!OnUsCode.Contains(dr["pan"].ToString().Substring(0, 6)))
    //                        {
    //                            amountDisputedOffUs += ejAmount;
    //                            amountDisputedOffUsEJ += ejAmount;

    //                            transDisputedoffus++;
    //                            transDisputedoffusEJ++;
    //                        }
    //                        else
    //                        {
    //                            amountDisputed += ejAmount;
    //                            amountDisputedOnUsEJ += ejAmount;

    //                            transDisputed++;
    //                            transDisputedonusEJ++;
    //                        }
    //                        if (listReconciledSwitchTrxn.Contains(trxn.ReconciliationSwitchDataId.Value))
    //                        {
    //                            ReconciliationTransactions reconTran = listReconciliationTransactions.Where(i => i.ReconciliationSwitchDataId == trxn.ReconciliationSwitchDataId.Value).SingleOrDefault();

    //                            //var reconTran = ReconciliationTransactions.LoadReconciliationTransactions("reconciliation_switch_data_id = " + trxn.ReconciliationSwitchDataId.Value);
    //                            reconTran.EjParsedTransactionsId = int.Parse(dr["ej_parsed_transactions_id"].ToString());
    //                            reconTran.ComparisonType = 0;
    //                            reconTran.Status = "Disputed";
    //                            reconTran.IsReconciled = false;
    //                            reconTran.Reason = "Transaction matched but status is failed on EJ_";
    //                            //reconTran.Save(dbTrxn.Connection, dbTrxn);
    //                            saveTrn = false;

    //                            //if (isLocalBankPAN.Contains(dr["pan"].ToString().Substring(0, 6)) || isAMEXBankPAN.Contains(dr["pan"].ToString().Substring(0, 2)))
    //                            //{
    //                            //    transDisputedOFFUS_KNET++;
    //                            //    amountDisputedOFFUS_KNET += switchAmount;
    //                            //}
    //                            //else if (isGCCBankPAN.Contains(dr["pan"].ToString().Substring(0, 6)))
    //                            //{
    //                            //    transDisputedOFFUS_GCC++;
    //                            //    amountDisputedOFFUS_GCC += switchAmount;
    //                            //}

    //                        }
    //                        else if (listDisputedSwitchTrxn.Contains(trxn.ReconciliationSwitchDataId.Value))
    //                        {
    //                            ReconciliationTransactions reconTran = listReconciliationTransactions.Where(i => i.ReconciliationSwitchDataId == trxn.ReconciliationSwitchDataId.Value).SingleOrDefault();

    //                            //                            var reconTran = ReconciliationTransactions.LoadReconciliationTransactions("reconciliation_switch_data_id = " + trxn.ReconciliationSwitchDataId.Value);
    //                            reconTran.EjParsedTransactionsId = int.Parse(dr["ej_parsed_transactions_id"].ToString());
    //                            //reconTran.ComparisonType = 2;
    //                            //reconTran.Status = "Disputed";
    //                            //reconTran.Reason = "Transaction matched but status is failed on EJ";
    //                            //                          reconTran.Save(dbTrxn.Connection, dbTrxn);
    //                            saveTrn = false;
    //                        }
    //                    }
    //                    else if (((switchAmount == ejAmount) || (temp <= batch.AcceptableDifference && temp != 0)) & (status != switchResponseCode))
    //                    {
    //                        trxn.Reason = "Transaction Response Code Mismatch on Switch";
    //                        trxn.Status = "Disputed";
    //                        trxn.IsReconciled = false;
    //                        noOfRecordsFailedToReconciled++;

    //                        if (listDisputedSwitchTrxnResponse.Contains(trxn.ReconciliationSwitchDataId.Value))
    //                        {
    //                            ReconciliationTransactions reconTran = listReconciliationTransactions.Where(i => i.ReconciliationSwitchDataId == trxn.ReconciliationSwitchDataId.Value).SingleOrDefault();

    //                            //                            var reconTran = ReconciliationTransactions.LoadReconciliationTransactions("reconciliation_switch_data_id = " + trxn.ReconciliationSwitchDataId.Value);
    //                            reconTran.EjParsedTransactionsId = int.Parse(dr["ej_parsed_transactions_id"].ToString());
    //                            reconTran.Reason += trxn.Reason;
    //                            reconTran.IsReconciled = false;
    //                            //reconTran.ComparisonType = 2;
    //                            //reconTran.Status = "Disputed";
    //                            //reconTran.Reason = "Transaction matched but status is failed on EJ";
    //                            //                          reconTran.Save(dbTrxn.Connection, dbTrxn);
    //                            saveTrn = false;
    //                        }


    //                        if (!listDisputedSwitchTrxn.Contains(trxn.ReconciliationSwitchDataId.Value))
    //                        {

    //                            {
    //                                amountDisputed += ejAmount;

    //                                transDisputed++;
    //                            }
    //                        }


    //                        {
    //                            amountDisputedOnUsSwitch += ejAmount;

    //                            transDisputedonusSwitch++;
    //                        }

    //                    }
    //                    else
    //                    {

    //                        trxn.Reason = "Transaction Amount Mismatch on Switch. Amount Reported on Switch is " + drArray[0]["transaction_amount"].ToString();
    //                        trxn.Status = "Disputed";
    //                        trxn.IsReconciled = false;

    //                        if (listDisputedSwitchTrxnResponse.Contains(trxn.ReconciliationSwitchDataId.Value))
    //                        {
    //                            ReconciliationTransactions reconTran = listReconciliationTransactions.Where(i => i.ReconciliationSwitchDataId == trxn.ReconciliationSwitchDataId.Value).SingleOrDefault();

    //                            //                            var reconTran = ReconciliationTransactions.LoadReconciliationTransactions("reconciliation_switch_data_id = " + trxn.ReconciliationSwitchDataId.Value);
    //                            reconTran.EjParsedTransactionsId = int.Parse(dr["ej_parsed_transactions_id"].ToString());
    //                            reconTran.Reason += trxn.Reason;
    //                            reconTran.IsReconciled = false;
    //                            //reconTran.ComparisonType = 2;
    //                            //reconTran.Status = "Disputed";
    //                            //reconTran.Reason = "Transaction matched but status is failed on EJ";
    //                            //                          reconTran.Save(dbTrxn.Connection, dbTrxn);
    //                            saveTrn = false;
    //                        }
    //                        noOfRecordsFailedToReconciled++;
    //                        if (!listDisputedSwitchTrxn.Contains(trxn.ReconciliationSwitchDataId.Value))
    //                        {
    //                            listDisputedSwitchTrxn.Add(trxn.ReconciliationSwitchDataId.Value);


    //                            {
    //                                amountDisputed += ejAmount;

    //                                transDisputed++;
    //                            }
    //                        }


    //                        {
    //                            amountDisputedOnUsSwitch += ejAmount;

    //                            transDisputedonusSwitch++;
    //                        }
    //                    }
    //                }
    //                else
    //                {
    //                    if (dr["status"].ToString() != "1") //ignore if ej is failed and switch dont have entry
    //                    {
    //                        trxn.ReconciliationSwitchDataId = 0;
    //                        trxn.Reason = "Transaction missing on Switch";
    //                        trxn.Status = "Disputed";
    //                        trxn.IsReconciled = false;
    //                        noOfRecordsFailedToReconciled++;


    //                        {
    //                            amountDisputed += decimal.Parse(dr["amount"].ToString());
    //                            amountDisputedOnUsSwitch += decimal.Parse(dr["amount"].ToString());

    //                            transDisputed++;
    //                            transDisputedonusSwitch++;
    //                        }
    //                    }
    //                }
    //                if (saveTrn)
    //                {
    //                    trxn.UserId = 1;
    //                    trxn.GeneratedAt = DateTime.Now;
    //                    trxn.ComparisonType = 2;
    //                    //trxn.Save(dbTrxn.Connection, dbTrxn);
    //                    listReconciliationTransactions.Add(trxn);
    //                }

    //            }
    //            foreach (DataRow dr in dtSwitch.Rows)
    //            {

    //                // Exclude Internationl OFF US cards
    //                //if (!
    //                //    (isLocalBankPAN.Contains(dr["card_number"].ToString().Substring(0, 6))
    //                //    || isGCCBankPAN.Contains(dr["card_number"].ToString().Substring(0, 6))
    //                //    || isGBKBankPAN.Contains(dr["card_number"].ToString().Substring(0, 6))
    //                //    || dr["card_number"].ToString().Substring(0, 6) == "999999"
    //                //    || isAMEXBankPAN.Contains(dr["card_number"].ToString().Substring(0, 2)))
    //                //    )
    //                //{
    //                //    continue;
    //                //}
    //                noOfRecordsProcessed++;
    //                ReconciliationTransactions trxn = new ReconciliationTransactions();
    //                trxn.ReconciliationBatchId = batchID;
    //                trxn.ReconciliationSwitchDataId = int.Parse(dr["reconciliation_switch_data_id"].ToString());
    //                if (switchIdProcessed.Contains(trxn.ReconciliationSwitchDataId.Value))
    //                    continue;

    //                string date = dr["transaction_date"].ToString();
    //                string time = dr["transaction_time"].ToString();
    //                time = "00:00:00";
    //                string[] dateSplit = dr["transaction_date"].ToString().Split(' ');

    //                var dateParts = DateTime.ParseExact(dateSplit[0], "dd/MM/yyyy", null); ;
    //                string[] timeParts = time.Split(':');
    //                DateTime dateTime = new DateTime(dateParts.Year, dateParts.Month, dateParts.Day, int.Parse(timeParts[0]), int.Parse(timeParts[1]), int.Parse("00"));
    //                DateTime TodateTime = new DateTime(dateParts.Year, dateParts.Month, dateParts.Day, 23, 59, 59);

    //                int trn = Convert.ToInt32(dr["transaction_sequence"].ToString().Substring(dr["transaction_sequence"].ToString().Length - 4));
    //                DataRow[] drArray = dtEj.Select(
    //                    string.Format("title = '{0}' and pan = '{1}' and tsn like '%{2}'",
    //                                   dr["atm_id"], dr["card_number"].ToString(), trn));

    //                if (dr["card_number"].ToString() == "507808******7867")
    //                {
    //                    int c = 0;
    //                }

    //                //DataRow[] drArray = dtEj.Select(
    //                //    string.Format("title = '{0}' and pan = '{1}' and trxn_datetime>= '{2}' and trxn_datetime<= '{3}' and tsn like '%{4}'",
    //                //                   dr["atm_id"], dr["card_number"].ToString(), dateTime, TodateTime, trn));

    //                //dr["pan"].ToString().Substring(dr["pan"].ToString().Length - 4)

    //                if (drArray.Length == 0)
    //                {
    //                    if (!listDisputedSwitchTrxn.Contains(trxn.ReconciliationSwitchDataId.Value))
    //                    {
    //                        noOfRecordsFailedToReconciled++;
    //                        trxn.EjParsedTransactionsId = 0;
    //                        trxn.IsReconciled = false;
    //                        trxn.Reason = "Transaction missing on Ej";
    //                        trxn.Status = "Disputed";
    //                        trxn.UserId = 1;
    //                        trxn.GeneratedAt = DateTime.Now;
    //                        trxn.ComparisonType = 2;
    //                        //trxn.Save(dbTrxn.Connection, dbTrxn);
    //                        listReconciliationTransactions.Add(trxn);
    //                        {
    //                            amountDisputed += decimal.Parse(dr["transaction_amount"].ToString());
    //                            transDisputed++;
    //                        }

    //                        {
    //                            amountDisputedOnUsEJ += decimal.Parse(dr["transaction_amount"].ToString());
    //                            transDisputedonusEJ++;
    //                        }
    //                    }


    //                }
    //                else
    //                {
    //                    decimal ejAmount = decimal.Parse(drArray[0]["amount"].ToString());
    //                    decimal switchAmount = decimal.Parse(dr["transaction_amount"].ToString());
    //                    // trxn.EjParsedTransactionsId = int.Parse(drArray[0]["ej_parsed_transactions_id"].ToString());

    //                    if (ejAmount != switchAmount)
    //                    {
    //                        decimal temp = GetAcceptableDiffValue(batch, decimal.Parse(dr["transaction_amount"].ToString()), decimal.Parse(drArray[0]["amount"].ToString()));
    //                        if (!(temp <= batch.AcceptableDifference && temp != 0))
    //                        {
    //                            //                            if (!listDisputedSwitchTrxn.Contains(trxn.ReconciliationSwitchDataId.Value))

    //                            if (!listDisputedSwitchTrxn.Contains(trxn.ReconciliationSwitchDataId.Value)
    //                            && listReconciledSwitchTrxn.Contains(trxn.ReconciliationSwitchDataId.Value))
    //                            {
    //                                noOfRecordsFailedToReconciled++;
    //                                trxn.EjParsedTransactionsId = int.Parse(drArray[0]["ej_parsed_transactions_id"].ToString());
    //                                trxn.IsReconciled = false;
    //                                trxn.Reason = "Transaction Amount Mismatched On Switch and EJ";
    //                                trxn.Status = "Disputed";
    //                                trxn.UserId = 1;
    //                                trxn.GeneratedAt = DateTime.Now;
    //                                trxn.ComparisonType = 2;
    //                                //trxn.Save(dbTrxn.Connection, dbTrxn);
    //                                listReconciliationTransactions.Add(trxn);


    //                                {
    //                                    amountDisputed += decimal.Parse(dr["transaction_amount"].ToString());

    //                                    transDisputed++;
    //                                }



    //                                {
    //                                    amountDisputedOnUsEJ += decimal.Parse(dr["transaction_amount"].ToString());

    //                                    transDisputedonusEJ++;
    //                                }
    //                            }
    //                        }
    //                    }
    //                }

    //            }

    //        }
    //        private void CompareSwitchAndHost(DataTable dtEj, DataTable dtSwitch, DataTable dtSwitchCardless_Revesal, DataTable dtHost, int batchID, ref decimal amountDisputed, ref decimal amountDisputedOnUsHost, ref decimal amountDisputedOnUsSwitch, ref decimal amountDisputedCardless, ref decimal amountDisputedCardlessHost, ref decimal amountDisputedCardlessSwitch, ref decimal amountDisputedCardlessEJ, ref decimal transDisputed, ref decimal transDisputedonusHost, ref decimal transDisputedonusSwitch, ref decimal transDisputedcardless, ref decimal transDisputedcardlessHost, ref decimal transDisputedcardlessSwitch, ref decimal transDisputedcardlessEJ, ref decimal transDisputedonusEJ, ref decimal amountDisputedOnUsEJ, ref decimal transDisputedoffusHost, ref decimal amountDisputedOffUsHost)
    //        {
    //            ReconciliationBatch batch = ReconciliationBatch.LoadReconciliationBatchByPk(batchID);
    //            List<int> processedHostID = new List<int>();
    //            int c = 0;
    //            foreach (DataRow dr in dtSwitch.Rows)
    //            {
    //                noOfRecordsProcessed++;
    //                ReconciliationTransactions trxn = new ReconciliationTransactions();
    //                trxn.ReconciliationBatchId = batchID;
    //                trxn.ReconciliationSwitchDataId = int.Parse(dr["reconciliation_switch_data_id"].ToString());

    //                string terminalID = dr["atm_id"].ToString();

    //                string cardNoSwitch = dr["card_number"].ToString();
    //                string cardPrefixSixDig = cardNoSwitch.Substring(0, 6);
    //                string cardSuffixFourDig = cardNoSwitch.Substring(cardNoSwitch.Length - 4, 4);
    //                string cardNoMasked = cardPrefixSixDig + "******" + cardSuffixFourDig;

    //                string amount = dr["transaction_amount"].ToString();
    //                string tsn = dr["transaction_sequence"].ToString();
    //                //if (tsn != "8718" && terminalID == "QNB0521")
    //                //{
    //                //    c++;
    //                //    continue;
    //                //}
    //                //if (amount.Contains("7530"))
    //                //{
    //                //    int c = 0;
    //                //}
    //                string[] dateSplit = dr["transaction_date"].ToString().Split(' ');
    //                string qDate = DateTime.ParseExact(dateSplit[0], "dd/MM/yyyy", null).ToString("MM/dd/yyyy");

    //                DataRow[] drArrayHostEntry = dtHost.Select(string.Format("atm_id = '{0}' and card_number='{1}' and transaction_sequence = '{2}' and transaction_amount = '{3}'", terminalID, cardNoMasked, tsn, amount)).ToArray();

    //                if (drArrayHostEntry.Length > 0) //Transaction found in host.
    //                {
    //                    trxn.ReconciliationHostDataId = int.Parse(drArrayHostEntry[0]["reconciliation_host_data_id"].ToString());
    //                    processedHostID.Add(trxn.ReconciliationHostDataId.Value);

    //                    //new
    //                    int switchResponseCode = dr["transaction_response"].ToString().Contains("APPROVED") ? 0 : 1;
    //                    int hostResponseCode = 0; // Convert.ToInt32(drArrayHostEntry[0]["transaction_response"].ToString());
    //                                              //
    //                    if (switchResponseCode != 0)
    //                        switchResponseCode = 1;

    //                    if (hostResponseCode != 0)
    //                        hostResponseCode = 1;


    //                    decimal switchAmount = Math.Round(decimal.Parse(dr["transaction_amount"].ToString()), 2);
    //                    decimal hostAmount = Math.Round(decimal.Parse(drArrayHostEntry[0]["transaction_amount"].ToString()), 2);

    //                    //check if failed in EJ, --Iyju GBK
    //                    string date = dr["transaction_date"].ToString();
    //                    string time = dr["transaction_time"].ToString();
    //                    time = "00:00:00";

    //                    var dateParts = DateTime.ParseExact(dateSplit[0], "dd/MM/yyyy", null);// DateTime.ParseExact(date, "yyyyMMdd", null); 

    //                    string[] timeParts = time.Split(':');
    //                    DateTime dateTime = new DateTime(dateParts.Year, dateParts.Month, dateParts.Day, int.Parse(timeParts[0]), int.Parse(timeParts[1]), int.Parse("00"));
    //                    DateTime TodateTime = new DateTime(dateParts.Year, dateParts.Month, dateParts.Day, 23, 59, 59);

    //                    int trn = Convert.ToInt32(tsn);

    //                    DataRow[] drEJArray = dtEj.Select(string.Format("title = '{0}' and pan = '{1}'  and tsn like'%{2}'", dr["atm_id"], cardNoMasked, trn));

    //                    //amount == switchamount
    //                    if (switchAmount == hostAmount && switchResponseCode == hostResponseCode)
    //                    {

    //                        if (drEJArray.Length == 0) //Transaction NOT found in EJ.
    //                        {
    //                            trxn.IsReconciled = false;
    //                            trxn.Status = "Disputed";
    //                            trxn.Reason = "Transaction missing on Ej";
    //                            listDisputedSwitchTrxn.Add(trxn.ReconciliationSwitchDataId.Value);
    //                            noOfRecordsFailedToReconciled++;

    //                            amountDisputed += switchAmount;
    //                            amountDisputedOnUsEJ += switchAmount;
    //                            transDisputed++;
    //                            transDisputedonusEJ++;

    //                        }
    //                        else
    //                        {
    //                            //Added by IK on 27/11/2024 to update EjParsedTransactionsId if all 3 sources are matched.
    //                            //***********************************************************************************************************************************************************
    //                            trxn.EjParsedTransactionsId = int.Parse(drEJArray[0]["ej_parsed_transactions_id"].ToString());
    //                            //***********************************************************************************************************************************************************
    //                            int ejStatus = int.Parse(drEJArray[0]["status"].ToString());
    //                            if (ejStatus == switchResponseCode)
    //                            {
    //                                trxn.IsReconciled = true;
    //                                trxn.Status = "Matched";
    //                                noOfRecordsReconciled++;
    //                                listReconciledSwitchTrxn.Add(trxn.ReconciliationSwitchDataId.Value);
    //                            }
    //                            else
    //                            {
    //                                trxn.IsReconciled = false;
    //                                trxn.Status = "Disputed";
    //                                trxn.Reason = "Status Mismatched";
    //                                listDisputedSwitchTrxn.Add(trxn.ReconciliationSwitchDataId.Value);
    //                            }
    //                        }
    //                    }
    //                    //new
    //                    else if (switchResponseCode != hostResponseCode && amount == hostAmount.ToString())
    //                    {
    //                        // trxn.ReconciliationHostDataId = int.Parse(drArray[0]["reconciliation_host_data_id"].ToString());
    //                        if (switchResponseCode != 0 && hostResponseCode != 0)
    //                        {

    //                            if (drEJArray.Length == 0) //Transaction NOT found in EJ.
    //                            {
    //                                trxn.IsReconciled = false;
    //                                trxn.Status = "Disputed";
    //                                trxn.Reason = "Transaction missing on Ej";
    //                                listDisputedSwitchTrxn.Add(trxn.ReconciliationSwitchDataId.Value);
    //                                noOfRecordsFailedToReconciled++;

    //                                amountDisputedOnUsEJ += decimal.Parse(dr["transaction_amount"].ToString());
    //                                transDisputedonusEJ++;
    //                            }
    //                            else
    //                            {
    //                                trxn.IsReconciled = true;
    //                                trxn.Status = "Matched";
    //                                noOfRecordsReconciled++;
    //                                listReconciledSwitchTrxn.Add(trxn.ReconciliationSwitchDataId.Value);
    //                            }
    //                        }
    //                        else
    //                        {

    //                            trxn.IsReconciled = false;
    //                            trxn.Status = "Disputed";
    //                            trxn.Reason = "Transaction Response Code Mismatch on Host.";
    //                            noOfRecordsFailedToReconciled++;

    //                            {
    //                                if (hostResponseCode != 0)
    //                                {
    //                                    amountDisputedOnUsHost += switchAmount;
    //                                    transDisputedonusHost++;
    //                                }
    //                                else if (switchResponseCode != 0)
    //                                {
    //                                    amountDisputedOnUsSwitch += switchAmount;
    //                                    transDisputedonusSwitch++;


    //                                    int ejResponseCode = drEJArray.Count() > 0 ? Convert.ToInt32(drEJArray[0]["status"].ToString()) : -1;

    //                                    if (ejResponseCode != 0)
    //                                    {
    //                                        amountDisputedOnUsEJ += switchAmount;
    //                                        transDisputedonusEJ++;
    //                                    }
    //                                }
    //                                amountDisputed += switchAmount;
    //                                transDisputed++;


    //                            }
    //                            trxn.IsReconciled = true;
    //                            //trxn.Status = "Matched";
    //                            //noOfRecordsReconciled++;
    //                            //listReconciledSwitchTrxn.Add(trxn.ReconciliationSwitchDataId.Value);

    //                            listDisputedSwitchTrxn.Add(trxn.ReconciliationSwitchDataId.Value);
    //                            listDisputedSwitchTrxnResponse.Add(trxn.ReconciliationSwitchDataId.Value);


    //                        }
    //                    }
    //                    //
    //                    else
    //                    {
    //                        // trxn.ReconciliationHostDataId = int.Parse(drArray[0]["reconciliation_host_data_id"].ToString());
    //                        trxn.IsReconciled = false;
    //                        trxn.Status = "Disputed";
    //                        trxn.Reason = "Transaction Amount Mismatch, amount on Host is " + drArrayHostEntry[0]["transaction_amount"].ToString();
    //                        noOfRecordsFailedToReconciled++;
    //                        if (CardLessCode.Contains(dr["card_number"].ToString()))
    //                        {
    //                            amountDisputedCardless += switchAmount;
    //                            amountDisputedCardlessHost += switchAmount;

    //                            transDisputedcardless++;
    //                            transDisputedcardlessHost++;
    //                        }
    //                        else
    //                        {
    //                            amountDisputed += switchAmount;
    //                            amountDisputedOnUsHost += switchAmount;

    //                            transDisputed++;
    //                            transDisputedonusHost++;
    //                        }
    //                        listDisputedSwitchTrxn.Add(trxn.ReconciliationSwitchDataId.Value);
    //                    }
    //                }
    //                else
    //                {//missing in host without reversals dthost
    //                 //check if all entreis failed, then ignore--Iyju GBK
    //                    string date = dr["transaction_date"].ToString();
    //                    string time = dr["transaction_time"].ToString();
    //                    time = "00:00:00";
    //                    var dateParts = DateTime.ParseExact(date, "dd/MM/yyyy hh:mm:ss tt", null); ;
    //                    string[] timeParts = time.Split(':');
    //                    DateTime dateTime = new DateTime(dateParts.Year, dateParts.Month, dateParts.Day, int.Parse(timeParts[0]), int.Parse(timeParts[1]), int.Parse("00"));
    //                    DateTime TodateTime = new DateTime(dateParts.Year, dateParts.Month, dateParts.Day, 23, 59, 59);


    //                    int trn = Convert.ToInt32(dr["transaction_sequence"].ToString().Substring(dr["transaction_sequence"].ToString().Length - 4));
    //                    DataRow[] drEJArray = dtEj.Select(
    //                                       string.Format("title = '{0}' and pan = '{1}' and trxn_datetime>= '{2}' and trxn_datetime<= '{3}' and tsn='{4}'",
    //                                                      dr["atm_id"], dr["card_number"].ToString(), dateTime, TodateTime, trn));


    //                    int switchResponseCode = Convert.ToInt32(dr["transaction_response"].ToString());
    //                    int ejResponseCode = drEJArray.Count() > 0 ? Convert.ToInt32(drEJArray[0]["status"].ToString()) : -1;
    //                    if (ejResponseCode != 0 && switchResponseCode != 0 && drEJArray.Count() > 0)
    //                    {
    //                        //both are failed... no dispute

    //                        {
    //                            trxn.ReconciliationHostDataId = 0;
    //                            trxn.IsReconciled = false;
    //                            // trxn.Status = "Disputed";
    //                            trxn.Reason = "Transaction missing on Host & Failed in EJ and Switch";
    //                            noOfRecordsReconciled++;
    //                            listReconciledSwitchTrxn.Add(trxn.ReconciliationSwitchDataId.Value);
    //                        }
    //                    }
    //                    else
    //                    {
    //                        trxn.ReconciliationHostDataId = 0;
    //                        trxn.IsReconciled = false;
    //                        trxn.Status = "Disputed";
    //                        trxn.Reason = "Transaction missing on Host";
    //                        if (drEJArray.Length > 0)
    //                            trxn.EjParsedTransactionsId = int.Parse(drEJArray[0]["ej_parsed_transactions_id"].ToString());

    //                        noOfRecordsFailedToReconciled++;

    //                        {
    //                            amountDisputed += decimal.Parse(dr["transaction_amount"].ToString());
    //                            amountDisputedOnUsHost += decimal.Parse(dr["transaction_amount"].ToString());

    //                            transDisputed++;
    //                            transDisputedonusHost++;
    //                        }
    //                        listDisputedSwitchTrxn.Add(trxn.ReconciliationSwitchDataId.Value);
    //                    }
    //                }
    //                trxn.UserId = 1;
    //                trxn.GeneratedAt = DateTime.Now;
    //                trxn.ComparisonType = 1;
    //                //trxn.Save(dbTrxn.Connection, dbTrxn);
    //                listReconciliationTransactions.Add(trxn);
    //            }



    //            foreach (DataRow dr in dtHost.Rows)
    //            {


    //                noOfRecordsProcessed++;
    //                ReconciliationTransactions trxn = new ReconciliationTransactions();
    //                trxn.ReconciliationBatchId = batchID;
    //                trxn.ReconciliationHostDataId = int.Parse(dr["reconciliation_host_data_id"].ToString());

    //                //DateTime hostTimePlus = DateTime.ParseExact(dr["transaction_time"].ToString(), "HH:mm",
    //                //                      CultureInfo.InvariantCulture);


    //                //string selectQuery_Switch = string.Format("atm_id = '{0}' and card_number='{1}' and transaction_date='{2}' " +
    //                //    " and transaction_amount='{3}' and transaction_sequence='{4}'",
    //                //   dr["atm_id"], dr["card_number"], dr["transaction_date"]
    //                //   , dr["transaction_amount"], dr["transaction_sequence"]);

    //                string selectQuery_Switch = string.Format("atm_id = '{0}' and card_number='{1}' " +
    //                    " and transaction_sequence like '%{2}'",
    //                   dr["atm_id"], dr["card_number"], dr["transaction_sequence"]);


    //                DataRow[] drArray = dtSwitch.Select(selectQuery_Switch);

    //                if (drArray.Length == 0) //Transaction NOT found in switch.
    //                {
    //                    //if (CardLessCode.Contains(dr["card_number"].ToString())) //include cardless as cardless entry reversals not support by OBAPI and will come as dispute



    //                    //check if missing in EJ &switch and  failed in HOST (Duplicate entry) , then ignore 
    //                    int hostResponseCode = Convert.ToInt32(dr["transaction_response"].ToString());

    //                    string date = dr["transaction_date"].ToString();
    //                    string time = dr["transaction_time"].ToString();
    //                    time = "00:00:00";
    //                    string[] dateSplit = dr["transaction_date"].ToString().Split(' ');
    //                    var dateParts = DateTime.ParseExact(dateSplit[0], "MM/dd/yyyy", null);

    //                    //var dateParts = DateTime.ParseExact(date, "yyyyMMdd", null); ;
    //                    string[] timeParts = time.Split(':');
    //                    /* string terminalID = dr["atm_id"].ToString();
    //                string cardno = dr["card_number"].ToString();
    //                string amount = dr["transaction_amount"].ToString().Replace(".", "") + ".000";
    //                string tsn = dr["transaction_sequence"].ToString().Substring(1);
    //                string[] dateSplit = dr["transaction_date"].ToString().Split(' ');
    //                string qDate = DateTime.ParseExact(dateSplit[0], "dd/MM/yyyy", null).ToString("MM/dd/yyyy");
    //*/
    //                    DateTime dateTime = new DateTime(dateParts.Year, dateParts.Month, dateParts.Day, int.Parse(timeParts[0]), int.Parse(timeParts[1]), int.Parse("00"));
    //                    DateTime TodateTime = new DateTime(dateParts.Year, dateParts.Month, dateParts.Day, 23, 59, 59);

    //                    int trn = Convert.ToInt32(dr["transaction_sequence"].ToString().Substring(dr["transaction_sequence"].ToString().Length - 4));

    //                    string selectQuery_EJ = string.Format("title = '{0}' and pan = '{1}' and trxn_datetime>= '{2}' and trxn_datetime<= '{3}'  and tsn='{4}'",
    //                     dr["atm_id"], dr["card_number"].ToString(), dateTime, TodateTime, trn);

    //                    DataRow[] drEJArray = dtEj.Select(selectQuery_EJ);

    //                    if (drEJArray.Length == 0) //Transaction NOT found in EJ.
    //                    {
    //                        if (hostResponseCode != 0) //failed on host
    //                            continue;

    //                    }


    //                    //Added by IK on 01/12/2024 to update EjParsedTransactionsId if all 3 sources are matched.
    //                    //***********************************************************************************************************************************************************
    //                    if (drEJArray.Length > 0)
    //                        trxn.EjParsedTransactionsId = int.Parse(drEJArray[0]["ej_parsed_transactions_id"].ToString());
    //                    //***********************************************************************************************************************************************************


    //                    trxn.ReconciliationSwitchDataId = 0;
    //                    trxn.Reason = "Transaction missing on Switch";
    //                    trxn.IsReconciled = false;
    //                    trxn.Status = "Disputed";
    //                    noOfRecordsFailedToReconciled++;
    //                    trxn.UserId = 1;
    //                    trxn.GeneratedAt = DateTime.Now;
    //                    trxn.ComparisonType = 1;
    //                    //trxn.Save(dbTrxn.Connection, dbTrxn);
    //                    listReconciliationTransactions.Add(trxn);

    //                    {
    //                        amountDisputed += decimal.Parse(dr["transaction_amount"].ToString());
    //                        amountDisputedOnUsSwitch += decimal.Parse(dr["transaction_amount"].ToString());

    //                        transDisputed++;
    //                        transDisputedonusSwitch++;


    //                        if (drEJArray.Length == 0) //Transaction NOT found in EJ.
    //                        {
    //                            amountDisputedOnUsEJ += decimal.Parse(dr["transaction_amount"].ToString());
    //                            transDisputedonusEJ++;
    //                        }
    //                    }
    //                }
    //                else
    //                {
    //                    decimal hostAmount = decimal.Parse(dr["transaction_amount"].ToString());
    //                    decimal switchAmount = decimal.Parse(drArray[0]["transaction_amount"].ToString());

    //                    if (hostAmount.ToString().Replace(".", "").IndexOf(switchAmount.ToString().Replace(".", "")) < 0)

    //                    //if (hostAmount != switchAmount)
    //                    {
    //                        decimal temp = GetAcceptableDiffValue(batch, hostAmount, switchAmount);
    //                        if (!(temp <= batch.AcceptableDifference && temp != 0))
    //                        {
    //                            trxn.ReconciliationSwitchDataId = 0;
    //                            trxn.Reason = "Transaction missing on Switch";
    //                            trxn.IsReconciled = false;
    //                            trxn.Status = "Disputed";
    //                            noOfRecordsFailedToReconciled++;
    //                            trxn.UserId = 1;
    //                            trxn.GeneratedAt = DateTime.Now;
    //                            trxn.ComparisonType = 1;
    //                            //                            trxn.Save(dbTrxn.Connection, dbTrxn);
    //                            listReconciliationTransactions.Add(trxn);

    //                            {
    //                                amountDisputed += hostAmount;
    //                                amountDisputedOnUsSwitch += hostAmount;

    //                                transDisputed++;
    //                                transDisputedonusSwitch++;



    //                                //check if entreis failed in EJ, then update EJ amount & count--Iyju GBK
    //                                string date = dr["transaction_date"].ToString();
    //                                string time = dr["transaction_time"].ToString();
    //                                time = "00:00:00";
    //                                var dateParts = DateTime.ParseExact(date, "yyyyMMdd", null); ;
    //                                string[] timeParts = time.Split(':');

    //                                DateTime dateTime = new DateTime(dateParts.Year, dateParts.Month, dateParts.Day, int.Parse(timeParts[0]), int.Parse(timeParts[1]), int.Parse("00"));
    //                                DateTime TodateTime = new DateTime(dateParts.Year, dateParts.Month, dateParts.Day, 23, 59, 59);

    //                                int trn = Convert.ToInt32(dr["transaction_sequence"].ToString().Substring(dr["transaction_sequence"].ToString().Length - 4));

    //                                string selectQuery_EJ = string.Format("title = '{0}' and pan= '{1}' and trxn_datetime>= '{2}' and trxn_datetime<= '{3}' and tsn='{4}'",
    //                                 dr["atm_id"], dr["card_number"].ToString(), dateTime, TodateTime, trn);

    //                                DataRow[] drEJArray = dtEj.Select(selectQuery_EJ);
    //                                if (drEJArray.Length == 0) //Transaction NOT found in EJ.
    //                                {
    //                                    amountDisputedOnUsEJ += hostAmount;
    //                                    transDisputedonusEJ++;
    //                                }
    //                            }
    //                        }
    //                    }

    //                }
    //            }



    //        }
    //    private void CompareEjandSwitchBnaData(DataTable dtBnaEj, DataTable dtBnaSwitch, DataTable dtBnaHost, int batchID, ref decimal amountDisputedBNA, ref decimal transDisputedbna, ref decimal amountDisputedBNASwitch, ref decimal transDisputedbnaSwitch, ref decimal amountDisputedBNAEJ, ref decimal transDisputedbnaEJ, ref decimal amountDisputedBNAHost, ref decimal transDisputedbnaHost)
    //    {

    //        ReconciliationBatch batch = ReconciliationBatch.LoadReconciliationBatchByPk(batchID);
    //        foreach (DataRow dr in dtBnaEj.Rows)
    //        {
    //            if (CardLessCode.Contains(dr["pan"].ToString()))
    //            {
    //                continue;
    //            }
    //            bool saveTrn = true;
    //            noOfRecordsProcessed++;
    //            ReconciliationTransactions trxn = new ReconciliationTransactions();
    //            trxn.ReconciliationBatchId = batchID;
    //            trxn.EjParsedBnaTransactionId = int.Parse(dr["ej_parsed_bna_transaction_id"].ToString());
    //            string status = dr["status"].ToString().Trim();

    //            DateTime trxnDatetime = DateTime.Parse(dr["trxn_datetime"].ToString());
    //            string date = trxnDatetime.ToString("yyyyMMdd");
    //            string time = trxnDatetime.ToString("HH:mm");

    //            DataRow[] drArray = dtBnaSwitch.Select(
    //                string.Format("atm_id = '{0}' and card_number like '%{1}' and transaction_date='{2}' and transaction_time = '{3}' and transaction_sequence like '%{4}'",
    //                               dr["title"], dr["pan"].ToString().Substring(dr["pan"].ToString().Length - 4), date, time, dr["seq"]));
    //            if (drArray.Length > 0) //Transaction found in switch.
    //            {
    //                int ejResponseCode = dr["processed_tran"].ToString().Trim() != "" ? Convert.ToInt32(dr["processed_tran"].ToString()) : 1;
    //                int switchResponseCode = Convert.ToInt32(drArray[0]["transaction_response"].ToString());

    //                bool check = true;
    //                if (ejResponseCode == 0 && switchResponseCode == 0 && status != "Successful")
    //                {
    //                    check = false;
    //                }

    //                trxn.ReconciliationBnaSwitchDataId = int.Parse(drArray[0]["reconciliation_bna_switch_data_id"].ToString());
    //                decimal ejAmount = decimal.Parse(dr["amount_authorized"].ToString());
    //                decimal switchAmount = decimal.Parse(drArray[0]["transaction_amount"].ToString());
    //                decimal temp = GetAcceptableDiffValue(batch, switchAmount, ejAmount);

    //                if (((switchAmount == ejAmount) || (temp <= batch.AcceptableDifference && temp != 0)) & (ejResponseCode == switchResponseCode) & check)
    //                {
    //                    if (temp <= batch.AcceptableDifference && temp != 0)
    //                        trxn.Reason = "Transaction Amount Mismatched On Switch But reconciled because of acceptable difference";
    //                    trxn.IsReconciled = true;
    //                    trxn.Status = "Matched";
    //                    noOfRecordsReconciled++;
    //                    if (listReconciledSwitchBnaTrxn.Contains(trxn.ReconciliationBnaSwitchDataId.Value))
    //                    {
    //                        noOfRecordsReconciledCompletely++;
    //                        //amountReconciledCompletely += ejAmount;
    //                        //CreateGLAccountSummary(trxn.ReconciliationBatchId.Value, int.Parse(dr["atm_id"].ToString()), ejAmount, dbTrxn);
    //                        listReconciledTransactions.Add(CreateReconciledTrxnEntry(trxn.ReconciliationBatchId.Value, ejAmount, dr["title"].ToString(), dr["seq"].ToString(), trxnDatetime,
    //                            dr["pan"].ToString(), "EJ"));

    //                        ReconciliationTransactions reconTran = listReconciliationTransactions.Where(i => i.ReconciliationBnaSwitchDataId == trxn.ReconciliationBnaSwitchDataId.Value).SingleOrDefault();

    //                        //var reconTran = ReconciliationTransactions.LoadReconciliationTransactions("reconciliation_bna_switch_data_id = " + trxn.ReconciliationBnaSwitchDataId.Value);
    //                        reconTran.EjParsedBnaTransactionId = int.Parse(dr["ej_parsed_bna_transaction_id"].ToString());
    //                        reconTran.ComparisonType = 4;
    //                        // reconTran.Save(dbTrxn.Connection, dbTrxn);
    //                        saveTrn = false;

    //                    }
    //                    else if (listDisputedSwitchBNATrxn.Contains(trxn.ReconciliationBnaSwitchDataId.Value))
    //                    {
    //                        ReconciliationTransactions reconTran = listReconciliationTransactions.Where(i => i.ReconciliationBnaSwitchDataId == trxn.ReconciliationBnaSwitchDataId.Value).SingleOrDefault();

    //                        //  var reconTran = ReconciliationTransactions.LoadReconciliationTransactions("reconciliation_bna_switch_data_id = " + trxn.ReconciliationBnaSwitchDataId.Value);
    //                        reconTran.EjParsedBnaTransactionId = int.Parse(dr["ej_parsed_bna_transaction_id"].ToString());
    //                        //reconTran.ComparisonType = 2;
    //                        //reconTran.Status = "Disputed";
    //                        //reconTran.Reason = "Transaction matched but status is failed on EJ";
    //                        //reconTran.Save(dbTrxn.Connection, dbTrxn);
    //                        saveTrn = false;
    //                    }
    //                    else
    //                    {
    //                        //trxn.Reason = "Transaction Response Code Mismatch on Switch";
    //                        trxn.Status = "Matched";
    //                        trxn.IsReconciled = true;
    //                        noOfRecordsReconciled++;
    //                    }
    //                }
    //                else if (((switchAmount == ejAmount) || (temp <= batch.AcceptableDifference && temp != 0)) & (ejResponseCode == switchResponseCode) & !check)
    //                {
    //                    if (!listDisputedSwitchBNATrxn.Contains(trxn.ReconciliationBnaSwitchDataId.Value))
    //                    {

    //                        trxn.Reason = "Transaction matched but status is failed on EJ";
    //                        trxn.Status = "Disputed";
    //                        trxn.IsReconciled = false;
    //                        noOfRecordsFailedToReconciled++;
    //                        amountDisputedBNA += ejAmount;
    //                        amountDisputedBNAEJ += ejAmount;

    //                        transDisputedbna++;
    //                        transDisputedbnaEJ++;
    //                    }
    //                    else
    //                    {
    //                        saveTrn = false;
    //                    }


    //                    if (listReconciledSwitchBnaTrxn.Contains(trxn.ReconciliationBnaSwitchDataId.Value))
    //                    {
    //                        ReconciliationTransactions reconTran = listReconciliationTransactions.Where(i => i.ReconciliationBnaSwitchDataId == trxn.ReconciliationBnaSwitchDataId.Value).SingleOrDefault();

    //                        //var reconTran = ReconciliationTransactions.LoadReconciliationTransactions("reconciliation_bna_switch_data_id = " + trxn.ReconciliationBnaSwitchDataId.Value);
    //                        reconTran.EjParsedBnaTransactionId = int.Parse(dr["ej_parsed_bna_transaction_id"].ToString());
    //                        reconTran.ComparisonType = 4;
    //                        reconTran.Status = "Disputed";
    //                        reconTran.IsReconciled = false;
    //                        reconTran.Reason = "Transaction matched but status is failed on EJ";
    //                        //reconTran.Save(dbTrxn.Connection, dbTrxn);
    //                        saveTrn = false;
    //                    }

    //                }
    //                else if (((switchAmount == ejAmount) || (temp <= batch.AcceptableDifference && temp != 0)) & (ejResponseCode != switchResponseCode))
    //                {

    //                    if (!listDisputedSwitchBNATrxn.Contains(trxn.ReconciliationBnaSwitchDataId.Value))
    //                    {
    //                        LogableTask.LogMonoActivityTask("match1", MethodBase.GetCurrentMethod(), TraceLevel.Info, "switchAmount:" + switchAmount
    //                            + ", ejAmount:" + ejAmount + ", switchResponseCode:" + switchResponseCode + ", ejResponseCode:" + ejResponseCode);

    //                        trxn.Reason = "Transaction Response Code Mismatch on Switch.";
    //                        trxn.Status = "Disputed";
    //                        trxn.IsReconciled = false;
    //                        noOfRecordsFailedToReconciled++;



    //                        amountDisputedBNA += ejAmount;
    //                        amountDisputedBNASwitch += ejAmount;

    //                        transDisputedbna++;
    //                        transDisputedbnaSwitch++;
    //                    }
    //                    else
    //                    {
    //                        saveTrn = false;
    //                    }
    //                }
    //                else
    //                {
    //                    trxn.Reason = "Transaction Amount Mismatch on Switch.Amount Reported on Switch is " + drArray[0]["transaction_amount"].ToString();
    //                    trxn.Status = "Disputed";
    //                    trxn.IsReconciled = false;
    //                    noOfRecordsFailedToReconciled++;
    //                    if (!listDisputedSwitchBNATrxn.Contains(trxn.ReconciliationBnaSwitchDataId.Value))
    //                    {
    //                        amountDisputedBNA += ejAmount;
    //                        amountDisputedBNASwitch += ejAmount;

    //                        transDisputedbna++;
    //                        transDisputedbnaSwitch++;
    //                    }
    //                }
    //            }
    //            else
    //            {
    //                DataRow[] drArrayHost = dtBnaHost.Select(
    //              string.Format("atm_id = '{0}' and card_number='{1}' and transaction_date='{2}' and transaction_time = '{3}' and customer_account_no like '%{4}'",
    //                             dr["title"], dr["pan"], date, time, dr["account_no"]));

    //                if (drArrayHost.Length > 0 && status != "Failed")
    //                {

    //                    trxn.Reason = "Transaction missing on Switch";
    //                    trxn.Status = "Disputed";
    //                    trxn.IsReconciled = false;
    //                    noOfRecordsFailedToReconciled++;
    //                    amountDisputedBNA += decimal.Parse(dr["amount_authorized"].ToString());
    //                    amountDisputedBNASwitch += decimal.Parse(dr["amount_authorized"].ToString());

    //                    transDisputedbna++;
    //                    transDisputedbnaSwitch++;


    //                    if (drArrayHost.Length == 0)
    //                    {
    //                        amountDisputedBNAHost += decimal.Parse(dr["amount_authorized"].ToString());
    //                        transDisputedbnaHost++;
    //                    }
    //                }
    //                else
    //                {//missing in host and switch

    //                    if (status != "Failed")  //== "Suspicious")
    //                    {
    //                        trxn.Reason = "Transaction missing on Switch & Host";// Suspicious/Successful in EJ";
    //                        trxn.Status = "Disputed";
    //                        trxn.IsReconciled = false;
    //                        noOfRecordsFailedToReconciled++;

    //                        amountDisputedBNA += decimal.Parse(dr["amount_authorized"].ToString());
    //                        amountDisputedBNASwitch += decimal.Parse(dr["amount_authorized"].ToString());
    //                        amountDisputedBNAHost += decimal.Parse(dr["amount_authorized"].ToString());

    //                        transDisputedbna++;
    //                        transDisputedbnaSwitch++;
    //                        transDisputedbnaHost++;

    //                        if (status == "Suspicious")
    //                        {
    //                            amountDisputedBNAEJ += decimal.Parse(dr["amount_authorized"].ToString());
    //                            transDisputedbnaEJ++;
    //                        }
    //                    }
    //                    //else ignore as it failed in EJ
    //                }
    //            }
    //            if (saveTrn)
    //            {
    //                trxn.UserId = 1;
    //                trxn.GeneratedAt = DateTime.Now;
    //                trxn.ComparisonType = 5;
    //                //trxn.Save(dbTrxn.Connection, dbTrxn);
    //                listReconciliationTransactions.Add(trxn);
    //            }

    //        }

    //        foreach (DataRow dr in dtBnaSwitch.Rows)
    //        {
    //            if (CardLessCode.Contains(dr["card_number"].ToString()))
    //            {
    //                continue;
    //            }
    //            noOfRecordsProcessed++;
    //            ReconciliationTransactions trxn = new ReconciliationTransactions();
    //            trxn.ReconciliationBatchId = batchID;
    //            trxn.ReconciliationBnaSwitchDataId = int.Parse(dr["reconciliation_bna_switch_data_id"].ToString());


    //            string date = dr["transaction_date"].ToString();
    //            string time = dr["transaction_time"].ToString();

    //            var dateParts = DateTime.ParseExact(date, "yyyyMMdd", null);
    //            string[] timeParts = time.Split(':');
    //            DateTime dateTime = new DateTime(dateParts.Year, dateParts.Month, dateParts.Day, int.Parse(timeParts[0]), int.Parse(timeParts[1]), int.Parse("00"));

    //            //string c = dr["card_number"].ToString().Substring(dr["card_number"].ToString().Length - 4);
    //            //string t = dr["transaction_sequence"].ToString().Substring(dr["transaction_sequence"].ToString().Length - 4);

    //            int trn = Convert.ToInt32(dr["transaction_sequence"].ToString().Substring(dr["transaction_sequence"].ToString().Length - 4));

    //            //            LogableTask.LogMonoActivityTask("match", MethodBase.GetCurrentMethod(), TraceLevel.Info,
    //            //" Going to check switch data with this criteria " +
    //            //string.Format("title = '{0}' and pan like '%{1}' and trxn_datetime= '{2}' and tsn='{3}'",
    //            //                               dr["atm_id"], dr["card_number"].ToString().Trim().Substring(dr["card_number"].ToString().Trim().Length - 4),
    //            //                               dateTime, trn));

    //            DataRow[] drArray = dtBnaEj.Select(
    //                string.Format("title = '{0}' and pan like '%{1}' and trxn_datetime= '{2}' and seq='{3}'",
    //                               dr["atm_id"], dr["card_number"].ToString().Trim().Substring(dr["card_number"].ToString().Trim().Length - 4), dateTime, trn));


    //            if (drArray.Length == 0)
    //            {

    //                if (!listDisputedSwitchBNATrxn.Contains(trxn.ReconciliationBnaSwitchDataId.Value))
    //                {
    //                    noOfRecordsFailedToReconciled++;
    //                    trxn.EjParsedBnaTransactionId = 0;
    //                    trxn.IsReconciled = false;
    //                    trxn.Reason = "Transaction missing on Ej";
    //                    trxn.Status = "Disputed";
    //                    trxn.UserId = 1;
    //                    trxn.GeneratedAt = DateTime.Now;
    //                    trxn.ComparisonType = 5;
    //                    //trxn.Save(dbTrxn.Connection, dbTrxn);
    //                    listReconciliationTransactions.Add(trxn);
    //                    amountDisputedBNA += decimal.Parse(dr["transaction_amount"].ToString());
    //                    amountDisputedBNAEJ += decimal.Parse(dr["transaction_amount"].ToString());

    //                    transDisputedbna++;
    //                    transDisputedbnaEJ++;
    //                }

    //            }
    //            else
    //            {
    //                decimal ejAmount = decimal.Parse(drArray[0]["amount_authorized"].ToString());
    //                decimal switchAmount = decimal.Parse(dr["transaction_amount"].ToString());

    //                if (ejAmount != switchAmount)
    //                {
    //                    decimal temp = GetAcceptableDiffValue(batch, decimal.Parse(dr["transaction_amount"].ToString()), decimal.Parse(drArray[0]["amount_authorized"].ToString()));
    //                    if (!(temp <= batch.AcceptableDifference && temp != 0))
    //                    {

    //                        if (!listDisputedSwitchBNATrxn.Contains(trxn.ReconciliationBnaSwitchDataId.Value))
    //                        {
    //                            noOfRecordsFailedToReconciled++;
    //                            trxn.EjParsedBnaTransactionId = int.Parse(drArray[0]["ej_parsed_bna_transaction_id"].ToString());
    //                            trxn.IsReconciled = false;
    //                            trxn.Reason = "Transaction Amount Mismatched On Switch and EJ";
    //                            trxn.Status = "Disputed";
    //                            trxn.UserId = 1;
    //                            trxn.GeneratedAt = DateTime.Now;
    //                            trxn.ComparisonType = 5;
    //                            //                                trxn.Save(dbTrxn.Connection, dbTrxn);
    //                            listReconciliationTransactions.Add(trxn);

    //                            amountDisputedBNA += decimal.Parse(dr["transaction_amount"].ToString());
    //                            amountDisputedBNAEJ += decimal.Parse(dr["transaction_amount"].ToString());

    //                            transDisputedbna++;
    //                            transDisputedbnaEJ++;
    //                        }
    //                    }
    //                }
    //            }

    //        }
    //    }


    //    private void CompareSwitchAndHostBnaData(DataTable dtEj, DataTable dtBnaSwitch, DataTable dtBnaHost, DataTable dtCorrectedHostBNA, int batchID, ref decimal amountDisputedBNA, ref decimal transDisputedbna, ref decimal amountDisputedBNAHost, ref decimal transDisputedbnaHost, ref decimal amountDisputedBNASwitch, ref decimal transDisputedbnaSwitch, ref decimal amountDisputedBNAEJ, ref decimal transDisputedbnaEJ)
    //    {
    //        ReconciliationBatch batch = ReconciliationBatch.LoadReconciliationBatchByPk(batchID);
    //        foreach (DataRow dr in dtBnaSwitch.Rows)
    //        {
    //            if (CardLessCode.Contains(dr["card_number"].ToString()))
    //            {
    //                continue;
    //            }
    //            noOfRecordsProcessed++;
    //            ReconciliationTransactions trxn = new ReconciliationTransactions();
    //            trxn.ReconciliationBatchId = batchID;
    //            trxn.ReconciliationBnaSwitchDataId = int.Parse(dr["reconciliation_bna_switch_data_id"].ToString());

    //            DateTime switchTimePlus = DateTime.ParseExact(dr["transaction_time"].ToString(), "HH:mm",
    //                                   CultureInfo.InvariantCulture);
    //            switchTimePlus = switchTimePlus.AddMinutes(1);

    //            DataRow[] drArray = dtBnaHost.Select(
    //                string.Format("atm_id = '{0}' and card_number='{1}' and transaction_date='{2}' " +
    //                " and customer_account_no like '%{3}'" +
    //                " and transaction_amount = '{4}'  and transaction_sequence='{5}'",
    //                               dr["atm_id"], dr["card_number"].ToString().Trim(), dr["transaction_date"],
    //                               dr["customer_account_no"],
    //                               dr["transaction_amount"], dr["transaction_sequence"]));

    //            if (drArray.Length > 0) //Transaction found in host.
    //            {
    //                int switchResponseCode = Convert.ToInt32(dr["transaction_response"].ToString());
    //                int hostResponseCode = Convert.ToInt32(drArray[0]["transaction_response"].ToString());

    //                decimal switchAmount = decimal.Parse(dr["transaction_amount"].ToString());
    //                decimal hostAmount = decimal.Parse(drArray[0]["transaction_amount"].ToString());
    //                decimal temp = GetAcceptableDiffValue(batch, switchAmount, hostAmount);

    //                //check if all entreis failed, then ignore--Iyju GBK
    //                string date = dr["transaction_date"].ToString();
    //                string time = dr["transaction_time"].ToString();

    //                var dateParts = DateTime.ParseExact(date, "yyyyMMdd", null); ;
    //                string[] timeParts = time.Split(':');
    //                DateTime dateTime = new DateTime(dateParts.Year, dateParts.Month, dateParts.Day, int.Parse(timeParts[0]), int.Parse(timeParts[1]), int.Parse("00"));
    //                DateTime SwitchTimeMinus = dateTime.AddMinutes(-1);

    //                int trn = Convert.ToInt32(dr["transaction_sequence"].ToString().Substring(dr["transaction_sequence"].ToString().Length - 4));
    //                DataRow[] drEJArray = dtEj.Select(
    //                                   string.Format("title = '{0}' and pan like '%{1}' and( trxn_datetime= '{2}' or  trxn_datetime= '{3}') and seq='{4}'",
    //                                                  dr["atm_id"], dr["card_number"].ToString().Trim().Substring(dr["card_number"].ToString().Trim().Length - 4), dateTime, SwitchTimeMinus
    //                                                  , trn));


    //                if (((switchAmount == hostAmount) || (temp <= batch.AcceptableDifference && temp != 0)) & (switchResponseCode == hostResponseCode))
    //                //if (switchAmount == hostAmount)
    //                {
    //                    trxn.ReconciliationBnaHostDataId = int.Parse(drArray[0]["reconciliation_bna_host_data_id"].ToString());




    //                    string ejResponseCode = drEJArray.Count() > 0 ? drEJArray[0]["status"].ToString().Trim() : "";
    //                    if (ejResponseCode == "Failed" && switchResponseCode == 0)
    //                    {
    //                        trxn.EjParsedBnaTransactionId = int.Parse(drEJArray[0]["ej_parsed_bna_transaction_id"].ToString());
    //                        trxn.IsReconciled = false;
    //                        trxn.Status = "Disputed";
    //                        trxn.Reason = "Transaction failed on EJ";
    //                        noOfRecordsFailedToReconciled++;
    //                        amountDisputedBNA += decimal.Parse(dr["transaction_amount"].ToString());
    //                        amountDisputedBNAEJ += decimal.Parse(dr["transaction_amount"].ToString());

    //                        transDisputedbna++;
    //                        transDisputedbnaEJ++;
    //                        listDisputedSwitchBNATrxn.Add(trxn.ReconciliationBnaSwitchDataId.Value);
    //                    }
    //                    else if (ejResponseCode == "" && switchResponseCode == 0)
    //                    {
    //                        trxn.IsReconciled = false;
    //                        trxn.Status = "Disputed";
    //                        trxn.Reason = "Transaction missing on EJ";
    //                        noOfRecordsFailedToReconciled++;
    //                        amountDisputedBNA += decimal.Parse(dr["transaction_amount"].ToString());
    //                        amountDisputedBNAEJ += decimal.Parse(dr["transaction_amount"].ToString());

    //                        transDisputedbna++;
    //                        transDisputedbnaEJ++;
    //                        listDisputedSwitchBNATrxn.Add(trxn.ReconciliationBnaSwitchDataId.Value);
    //                    }
    //                    else
    //                    {

    //                        if (temp <= batch.AcceptableDifference && temp != 0)
    //                            trxn.Reason = "Transaction Amount Mismatched On Host But reconciled because of acceptable difference";

    //                        if (drEJArray.Length == 0) //Transaction NOT found in EJ.
    //                        {
    //                            trxn.IsReconciled = false;
    //                            trxn.Status = "Disputed";
    //                            trxn.Reason = "Transaction missing on Ej";
    //                            listDisputedSwitchBNATrxn.Add(trxn.ReconciliationBnaSwitchDataId.Value);
    //                            noOfRecordsFailedToReconciled++;

    //                            amountDisputedBNA += decimal.Parse(dr["transaction_amount"].ToString());
    //                            amountDisputedBNAEJ += decimal.Parse(dr["transaction_amount"].ToString());

    //                            transDisputedbna++;
    //                            transDisputedbnaEJ++;
    //                        }
    //                        else
    //                        {
    //                            trxn.IsReconciled = true;
    //                            trxn.Status = "Matched";
    //                            noOfRecordsReconciled++;
    //                            listReconciledSwitchBnaTrxn.Add(trxn.ReconciliationBnaSwitchDataId.Value);
    //                        }
    //                    }
    //                }
    //                else if ((switchResponseCode != hostResponseCode) & ((switchAmount == hostAmount) || (temp <= batch.AcceptableDifference && temp != 0)))
    //                {
    //                    trxn.ReconciliationBnaHostDataId = int.Parse(drArray[0]["reconciliation_bna_host_data_id"].ToString());
    //                    string ejResponseCode = drEJArray.Count() > 0 ? drEJArray[0]["status"].ToString().Trim() : "";

    //                    if (switchResponseCode != 0 && hostResponseCode != 0)
    //                    {
    //                        if (temp <= batch.AcceptableDifference && temp != 0)
    //                            trxn.Reason = "Transaction Amount Mismatched On Host But reconciled because of acceptable difference.";

    //                        if (ejResponseCode == "Failed")
    //                        {
    //                            trxn.Reason = "";
    //                            trxn.IsReconciled = true;
    //                            trxn.Status = "Matched";
    //                            noOfRecordsReconciled++;
    //                            listReconciledSwitchBnaTrxn.Add(trxn.ReconciliationBnaSwitchDataId.Value);
    //                        }
    //                        else if (ejResponseCode == "Successful")
    //                        {
    //                            trxn.EjParsedBnaTransactionId = int.Parse(drEJArray[0]["ej_parsed_bna_transaction_id"].ToString());
    //                            trxn.IsReconciled = false;
    //                            trxn.Status = "Disputed";
    //                            trxn.Reason = "Transaction failed on Switch and Host";
    //                            noOfRecordsFailedToReconciled++;
    //                            amountDisputedBNA += decimal.Parse(dr["transaction_amount"].ToString());
    //                            amountDisputedBNAEJ += decimal.Parse(dr["transaction_amount"].ToString());

    //                            transDisputedbna++;
    //                            transDisputedbnaEJ++;
    //                            listDisputedSwitchBNATrxn.Add(trxn.ReconciliationBnaSwitchDataId.Value);
    //                        }
    //                    }
    //                    else
    //                    {
    //                        if (hostResponseCode != 0 && switchResponseCode != 0 && ejResponseCode == "Failed")
    //                        {//ignore if failed on all ej switch and host
    //                            trxn.IsReconciled = true;
    //                            trxn.Status = "Matched";
    //                            noOfRecordsReconciled++;
    //                            listReconciledSwitchBnaTrxn.Add(trxn.ReconciliationBnaSwitchDataId.Value);
    //                        }
    //                        else
    //                        {
    //                            DataRow[] drcorrectedArray = dtCorrectedHostBNA.Select(
    //                          string.Format("atm_id = '{0}' and card_number='{1}' and transaction_date='{2}' " +
    //                          "and customer_account_no like '%{3}' and transaction_amount = '{4}' and transaction_sequence='{5}' ",
    //                                 dr["atm_id"], dr["card_number"].ToString().Trim(), dr["transaction_date"],
    //                                  dr["customer_account_no"], dr["transaction_amount"], dr["transaction_sequence"]));

    //                            if (drcorrectedArray.Length > 0)
    //                            {
    //                                if (int.Parse(drcorrectedArray[0]["transaction_response"].ToString()) == 0)
    //                                {
    //                                    trxn.ReconciliationBnaHostDataId = int.Parse(drcorrectedArray[0]["reconciliation_bna_host_data_id"].ToString());

    //                                    trxn.Reason = "Reversal Transaction";
    //                                    trxn.IsReconciled = true;
    //                                    trxn.Status = "Matched";
    //                                    noOfRecordsReconciled++;
    //                                    listReconciledSwitchBnaTrxn.Add(trxn.ReconciliationBnaSwitchDataId.Value);
    //                                }
    //                                else
    //                                {
    //                                    trxn.IsReconciled = false;
    //                                    trxn.Status = "Disputed";
    //                                    trxn.Reason = "Transaction Response Code Mismatch on Host.";
    //                                    noOfRecordsFailedToReconciled++;

    //                                    amountDisputedBNA += switchAmount;
    //                                    transDisputedbna++;

    //                                    if (hostResponseCode != 0)
    //                                    {
    //                                        amountDisputedBNAHost += switchAmount;
    //                                        transDisputedbnaHost++;
    //                                    }
    //                                    if (switchResponseCode != 0)
    //                                    {
    //                                        amountDisputedBNASwitch += hostAmount;
    //                                        transDisputedbnaSwitch++;
    //                                    }
    //                                    //check EJ status
    //                                    if (ejResponseCode == "Failed")
    //                                    {
    //                                        amountDisputedBNAEJ += hostAmount;
    //                                        transDisputedbnaEJ++;
    //                                    }

    //                                    listDisputedSwitchBNATrxn.Add(trxn.ReconciliationBnaSwitchDataId.Value);
    //                                }
    //                            }
    //                            else
    //                            {
    //                                if (hostResponseCode == 0 && switchResponseCode != 0)
    //                                {
    //                                    trxn.IsReconciled = false;
    //                                    trxn.Status = "Disputed";
    //                                    trxn.Reason = "Transaction Failed in Switch";
    //                                    noOfRecordsFailedToReconciled++;
    //                                    amountDisputedBNA += hostAmount;
    //                                    amountDisputedBNASwitch += hostAmount;

    //                                    transDisputedbna++;
    //                                    transDisputedbnaSwitch++;
    //                                    listDisputedSwitchBNATrxn.Add(trxn.ReconciliationBnaSwitchDataId.Value);

    //                                    if (ejResponseCode == "Failed")
    //                                    {
    //                                        trxn.Reason = "Transaction Failed in EJ and Switch";
    //                                        amountDisputedBNAEJ += hostAmount;
    //                                        transDisputedbnaEJ++;
    //                                    }
    //                                }
    //                                else if (hostResponseCode != 0 && switchResponseCode == 0)
    //                                {
    //                                    trxn.IsReconciled = false;
    //                                    trxn.Status = "Disputed";
    //                                    trxn.Reason = "Transaction Failed in Host";
    //                                    noOfRecordsFailedToReconciled++;
    //                                    amountDisputedBNA += switchAmount;
    //                                    amountDisputedBNAHost += switchAmount;

    //                                    transDisputedbna++;
    //                                    transDisputedbnaHost++;
    //                                    listDisputedSwitchBNATrxn.Add(trxn.ReconciliationBnaSwitchDataId.Value);

    //                                    if (ejResponseCode == "Failed")
    //                                    {
    //                                        trxn.Reason = "Transaction Failed in EJ and Host";
    //                                        amountDisputedBNAEJ += switchAmount;
    //                                        transDisputedbnaEJ++;
    //                                    }
    //                                }

    //                            }
    //                        }
    //                    }
    //                }
    //                else
    //                {
    //                    trxn.ReconciliationBnaHostDataId = int.Parse(drArray[0]["reconciliation_bna_host_data_id"].ToString());
    //                    trxn.IsReconciled = false;
    //                    trxn.Status = "Disputed";
    //                    trxn.Reason = "Transaction Amount Mismatch on Host.Amount Reported on Host is " + drArray[0]["transaction_amount"].ToString();
    //                    noOfRecordsFailedToReconciled++;
    //                    amountDisputedBNA += switchAmount;
    //                    amountDisputedBNAHost += switchAmount;

    //                    transDisputedbna++;
    //                    transDisputedbnaHost++;
    //                    listDisputedSwitchBNATrxn.Add(trxn.ReconciliationBnaSwitchDataId.Value);
    //                }
    //            }
    //            else
    //            {
    //                //check if all entreis failed, then ignore--Iyju GBK
    //                string date = dr["transaction_date"].ToString();
    //                string time = dr["transaction_time"].ToString();

    //                var dateParts = DateTime.ParseExact(date, "yyyyMMdd", null); ;
    //                string[] timeParts = time.Split(':');
    //                DateTime dateTime = new DateTime(dateParts.Year, dateParts.Month, dateParts.Day, int.Parse(timeParts[0]), int.Parse(timeParts[1]), int.Parse("00"));


    //                int trn = Convert.ToInt32(dr["transaction_sequence"].ToString().Substring(dr["transaction_sequence"].ToString().Length - 4));
    //                DataRow[] drEJArray = dtEj.Select(
    //                                   string.Format("title = '{0}' and pan like '%{1}' and trxn_datetime= '{2}' and seq='{3}'",
    //                                                  dr["atm_id"], dr["card_number"].ToString().Trim().Substring(dr["card_number"].ToString().Trim().Length - 4), dateTime, trn));


    //                int switchResponseCode = Convert.ToInt32(dr["transaction_response"].ToString());
    //                string ejResponseCode = drEJArray.Count() > 0 ? drEJArray[0]["status"].ToString().Trim() : "";
    //                if (ejResponseCode == "Failed" && switchResponseCode != 0)
    //                {
    //                    //both are failed... no dispute

    //                    trxn.ReconciliationBnaHostDataId = 0;
    //                    trxn.IsReconciled = false;
    //                    // trxn.Status = "Disputed";
    //                    trxn.Reason = "Transaction missing on Host & Failed in EJ and Switch";
    //                    noOfRecordsReconciled++;
    //                    listDisputedSwitchBNATrxn.Add(trxn.ReconciliationBnaSwitchDataId.Value);
    //                }
    //                else
    //                {
    //                    trxn.ReconciliationBnaHostDataId = 0;
    //                    trxn.IsReconciled = false;
    //                    trxn.Status = "Disputed";
    //                    trxn.Reason = "Transaction missing on Host";
    //                    noOfRecordsFailedToReconciled++;
    //                    amountDisputedBNA += decimal.Parse(dr["transaction_amount"].ToString());
    //                    amountDisputedBNAHost += decimal.Parse(dr["transaction_amount"].ToString());

    //                    transDisputedbna++;
    //                    transDisputedbnaHost++;
    //                    listDisputedSwitchBNATrxn.Add(trxn.ReconciliationBnaSwitchDataId.Value);

    //                    if (ejResponseCode == "Failed")
    //                    {
    //                        transDisputedbnaEJ++;
    //                        amountDisputedBNAEJ += decimal.Parse(dr["transaction_amount"].ToString());

    //                        trxn.Reason = "Transaction missing on Host & Failed in EJ";
    //                    }

    //                    if (switchResponseCode != 0)
    //                    {
    //                        transDisputedbnaSwitch++;
    //                        amountDisputedBNASwitch += decimal.Parse(dr["transaction_amount"].ToString());

    //                        trxn.Reason = "Transaction missing on Host & Failed in Switch";
    //                    }
    //                }
    //            }
    //            trxn.UserId = 1;
    //            trxn.GeneratedAt = DateTime.Now;
    //            trxn.ComparisonType = 6;
    //            //trxn.Save(dbTrxn.Connection, dbTrxn);
    //            //htReconciliationTransactions.Add(trxn.ReconciliationTransactionsId, trxn);
    //            listReconciliationTransactions.Add(trxn);
    //        }

    //        foreach (DataRow dr in dtBnaHost.Rows)
    //        {
    //            if (CardLessCode.Contains(dr["card_number"].ToString()))
    //            {
    //                continue;
    //            }

    //            noOfRecordsProcessed++;
    //            ReconciliationTransactions trxn = new ReconciliationTransactions();
    //            trxn.ReconciliationBatchId = batchID;
    //            trxn.ReconciliationBnaHostDataId = int.Parse(dr["reconciliation_bna_host_data_id"].ToString());

    //            DateTime hostTimePlus = DateTime.ParseExact(dr["transaction_time"].ToString(), "HH:mm",
    //                                 CultureInfo.InvariantCulture);
    //            hostTimePlus = hostTimePlus.AddMinutes(-1);

    //            DataRow[] drArray = dtBnaSwitch.Select(
    //string.Format("atm_id = '{0}' and card_number='{1}' and transaction_date='{2}' and customer_account_no='{3}' " +
    //" and transaction_amount = '{4}'  and transaction_sequence='{5}'",
    //               dr["atm_id"], dr["card_number"].ToString().Trim(), dr["transaction_date"]
    //               , dr["customer_account_no"].ToString().TrimStart('0'), dr["transaction_amount"], dr["transaction_sequence"]));

    //            if (drArray.Length == 0) //Transaction NOT found in switch.
    //            {
    //                trxn.ReconciliationBnaSwitchDataId = 0;
    //                trxn.Reason = "Transaction missing on Switch";
    //                trxn.IsReconciled = false;
    //                trxn.Status = "Disputed";
    //                noOfRecordsFailedToReconciled++;
    //                trxn.UserId = 1;
    //                trxn.GeneratedAt = DateTime.Now;
    //                trxn.ComparisonType = 6;
    //                //trxn.Save(dbTrxn.Connection, dbTrxn);
    //                listReconciliationTransactions.Add(trxn);

    //                amountDisputedBNA += decimal.Parse(dr["transaction_amount"].ToString());
    //                amountDisputedBNASwitch += decimal.Parse(dr["transaction_amount"].ToString());

    //                transDisputedbna++;
    //                transDisputedbnaSwitch++;

    //                //check if entreis failed in EJ, then update EJ amount & count--Iyju GBK
    //                string date = dr["transaction_date"].ToString();
    //                string time = dr["transaction_time"].ToString();

    //                var dateParts = DateTime.ParseExact(date, "yyyyMMdd", null); ;
    //                string[] timeParts = time.Split(':');
    //                DateTime dateTime = new DateTime(dateParts.Year, dateParts.Month, dateParts.Day, int.Parse(timeParts[0]), int.Parse(timeParts[1]), int.Parse("00"));

    //                int trn = Convert.ToInt32(dr["transaction_sequence"].ToString().Substring(dr["transaction_sequence"].ToString().Length - 4));
    //                DataRow[] drEJArray = dtEj.Select(string.Format("title = '{0}' and pan like '%{1}' and trxn_datetime= '{2}' and seq='{3}'",
    //                 dr["atm_id"], dr["card_number"].ToString().Trim().Substring(dr["card_number"].ToString().Trim().Length - 4), dateTime, trn));
    //                if (drEJArray.Length == 0) //Transaction NOT found in EJ.
    //                {
    //                    amountDisputedBNAEJ += decimal.Parse(dr["transaction_amount"].ToString());
    //                    transDisputedbnaEJ++;
    //                }
    //            }
    //            else
    //            {
    //                decimal hostAmount = decimal.Parse(dr["transaction_amount"].ToString());
    //                decimal switchAmount = decimal.Parse(drArray[0]["transaction_amount"].ToString());

    //                if (hostAmount != switchAmount)
    //                {
    //                    decimal temp = GetAcceptableDiffValue(batch, hostAmount, switchAmount);
    //                    if (!(temp <= batch.AcceptableDifference && temp != 0))
    //                    {
    //                        trxn.ReconciliationBnaSwitchDataId = 0;
    //                        trxn.Reason = "Transaction missing on Switch";
    //                        trxn.IsReconciled = false;
    //                        trxn.Status = "Disputed";
    //                        noOfRecordsFailedToReconciled++;
    //                        trxn.UserId = 1;
    //                        trxn.GeneratedAt = DateTime.Now;
    //                        trxn.ComparisonType = 6;
    //                        //trxn.Save(dbTrxn.Connection, dbTrxn);
    //                        listReconciliationTransactions.Add(trxn);

    //                        amountDisputedBNA += hostAmount;
    //                        amountDisputedBNASwitch += hostAmount;

    //                        transDisputedbna++;
    //                        transDisputedbnaSwitch++;

    //                        //check if entreis failed in EJ, then update EJ amount & count--Iyju GBK
    //                        string date = dr["transaction_date"].ToString();
    //                        string time = dr["transaction_time"].ToString();

    //                        var dateParts = DateTime.ParseExact(date, "yyyyMMdd", null); ;
    //                        string[] timeParts = time.Split(':');
    //                        DateTime dateTime = new DateTime(dateParts.Year, dateParts.Month, dateParts.Day, int.Parse(timeParts[0]), int.Parse(timeParts[1]), int.Parse("00"));

    //                        int trn = Convert.ToInt32(dr["transaction_sequence"].ToString().Substring(dr["transaction_sequence"].ToString().Length - 4));
    //                        DataRow[] drEJArray = dtEj.Select(string.Format("title = '{0}' and pan like '%{1}' and trxn_datetime= '{2}' and tsn='{3}'",
    //                         dr["atm_id"], dr["card_number"].ToString().Trim().Substring(dr["card_number"].ToString().Trim().Length - 4), dateTime, trn));
    //                        if (drEJArray.Length == 0) //Transaction NOT found in EJ.
    //                        {
    //                            amountDisputedBNAEJ += decimal.Parse(dr["transaction_amount"].ToString());
    //                            transDisputedbnaEJ++;
    //                        }
    //                    }
    //                }

    //            }
    //        }

    //    }
    #endregion

}
