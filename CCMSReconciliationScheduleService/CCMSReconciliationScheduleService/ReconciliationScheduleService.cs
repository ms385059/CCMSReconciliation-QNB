using Avanza.CCMS.DAL;
using Avanza.iSuite.DAL;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2021.DocumentTasks;
using Encryption;
using ExcelDataReader;
using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Reflection;
using System.ServiceProcess;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace ReconciliationScheduleService
{
    public partial class ReconciliationScheduleService : ServiceBase
    {
        Timer timerUpdateReconBatchStatus;
        Timer timer;
        Timer timerScheduleThreadForExecution;
        public static AppSetting appSettings;
        public static ReconciliationConfiguration recon;
        //public static MappingAtmTxns atmTxns;
        public static DateTime appSettingLastLoadedAt = DateTime.MinValue;
        public Dictionary<string, string> withdrawalTran = new Dictionary<string, string>();
        public Dictionary<string, string> depositTran = new Dictionary<string, string>();
        string minuteSchedule = System.Configuration.ConfigurationManager.AppSettings["minuteSchedule"];
        string Cardless_Tellerx_TerminalID = System.Configuration.ConfigurationManager.AppSettings["Cardless_Tellerx_TerminalID"];
        string Customer_FIT_BankName = System.Configuration.ConfigurationManager.AppSettings["Customer_FIT_BankName"];
        int Cardless_Tellerx_BatchID = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["Cardless_Tellerx_BatchID"]);
        int RemoteOnUs_BatchID = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["RemoteOnUs_BatchID"]);
        string BatchRerunNotified_ToEmailIDs = System.Configuration.ConfigurationManager.AppSettings["BatchRerunNotified_ToEmailIDs"];
        string RawHostDuplicateDeleteInterval = System.Configuration.ConfigurationManager.AppSettings["RawHostDuplicateDeleteInterval"];
        int BatchToFetchInDays = int.Parse(System.Configuration.ConfigurationManager.AppSettings["BatchToFetchInDays"]);
        Dictionary<string, int> atmData = null;
        Dictionary<string, ReconciliationBatchModel> ReconBatchDict = new Dictionary<string, ReconciliationBatchModel>();
        MappingAtmTxns.MappingAtmTxnsReader atmTxns = null;
        Atm.AtmReader _allAtms = null;
        ReconciliationBatch.ReconciliationBatchReader _batchRecons = null;
        bool isInitialized = false;
        public ReconciliationScheduleService()
        {
            InitializeComponent();
        }

        void ScheduleThreadForExecution(object state)
        {
            timer = new Timer(DoWork, null, new TimeSpan(0, 0, 15), new TimeSpan(0, 0, 0, 0, -1));
            timerUpdateReconBatchStatus = new Timer(UpdateReconBatchStatus, null, new TimeSpan(0, 1, 0), new TimeSpan(0, 0, 0, 0, -1));
        }
        private void DoInitialize()
        {
            try
            {
                string connectionStr = (string)Registry.LocalMachine.OpenSubKey("SOFTWARE\\NCR\\CCMS").GetValue("ConnectionString", "");
                connectionStr = Encryption.Cryptic.DecryptString(connectionStr);
                ConnectionFactory.Initialize(connectionStr, true);
                appSettings = AppSetting.LoadAppSetting("1=1");
                appSettingLastLoadedAt = DateTime.Now;

                XmlLogWriter.InitXmlLogWriter(String.Format("{0}\\CCMSReconciliationScheduleService_{1:yyMMMdd}.txt", appSettings.LogFilePath, DateTime.Now));
                try
                {
                    LogableTask.DefaultTraceLevel = (TraceLevel)Enum.Parse(typeof(TraceLevel), appSettings.ServiceLogLevel);
                }
                catch
                {
                    LogableTask.DefaultTraceLevel = TraceLevel.Info;
                    LogableTask.LogMonoActivityTask("GetTraceLevel", MethodBase.GetCurrentMethod(), TraceLevel.Error, "Failed to extract trace level from database");
                }
                LogableTask.LogMonoActivityTask("DisplayVersion", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "Version: Reconciliation Schedular 6.2.0, Build date : 11/08/2026");
                //LogableTask.LogMonoActivityTask("Schedular", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "Worker Threads begin execution after 15 seconds.");

                recon = ReconciliationConfiguration.LoadReconciliationConfiguration("1=1");

                LogableTask.LogMonoActivityTask("Reconciliation Configuration", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "Reconciliation Configuration loaded from Database.");

                atmTxns = MappingAtmTxns.ExecuteReader("1=1");
                while (atmTxns.Read())
                {
                    if (atmTxns.CurrentMappingAtmTxns.TransactionIndicator == "W")
                    {
                        if (!withdrawalTran.ContainsKey(atmTxns.CurrentMappingAtmTxns.TransactionCode))
                        {
                            withdrawalTran[atmTxns.CurrentMappingAtmTxns.TransactionCode] = atmTxns.CurrentMappingAtmTxns.SmallDescription.Trim();
                        }
                    }
                    else if (atmTxns.CurrentMappingAtmTxns.TransactionIndicator == "D")
                    {
                        if (!depositTran.ContainsKey(atmTxns.CurrentMappingAtmTxns.TransactionCode))
                        {
                            depositTran[atmTxns.CurrentMappingAtmTxns.TransactionCode] = atmTxns.CurrentMappingAtmTxns.SmallDescription.Trim();
                        }
                    }
                }

                if (!Directory.Exists(recon.BackupFolderPath))
                {
                    Directory.CreateDirectory(recon.BackupFolderPath);
                }
                if (atmData == null)
                {
                    atmData = new Dictionary<string, int>();

                    _allAtms = Atm.ExecuteReader("is_active = 1");

                    if (_allAtms != null)
                    {
                        while (_allAtms.Read())
                        {
                            string title = _allAtms.CurrentAtm.Title;
                            int atmId = _allAtms.CurrentAtm.ATMId;

                            // Store the title and corresponding ATMId in the dictionary
                            atmData[title] = atmId;
                        }
                    }

                }

                LogableTask.LogMonoActivityTask("ATM Data", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "ATM Dictionary loaded.");

                ReconBatchDict = new Dictionary<string, ReconciliationBatchModel>();
                _batchRecons = ReconciliationBatch.ExecuteReader(
                    $@"transaction_start_date >= convert(datetime,'{DateTime.Now.AddDays(-BatchToFetchInDays).ToString("dd/MM/yyyy")}',103) 
                     and transaction_start_date<=convert(datetime,'{DateTime.Now.ToString("dd/MM/yyyy")}',103)");

                if (_batchRecons != null)
                {
                    while (_batchRecons.Read())
                    {
                        InsertIntoReconciliationBatchDictionary(_batchRecons.CurrentReconciliationBatch.TransactionStartDate, _batchRecons.CurrentReconciliationBatch.AtmId,
                            _batchRecons.CurrentReconciliationBatch.ReconciliationBatchId);
                    }
                }

                LogableTask.LogMonoActivityTask("Reconciliation Batch", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "Reconciliation Batch Dictionary loaded.");

            }
            finally
            {
                if (atmTxns != null)
                    atmTxns.Close();
                if (_allAtms != null)
                    _allAtms.Close();
                if (_batchRecons != null)
                    _batchRecons.Close();

            }

        }

        void UpdateReconBatchStatus(object state)
        {
            timerUpdateReconBatchStatus.Change(-1, -1);
            try
            {
                string query = string.Format("creation_time>=convert(datetime,'{0}',103) and creation_time<=convert(datetime,'{0} 23:59:59',103) and status = 'Processed'" +
                   "and file_type in ('Switch','Host','KNET') ",
                   DateTime.Today.ToString("dd/MM/yyyy"));

                if ((int)ConnectionFactory.ExecuteScalar("select count(*) from Reconciliation_Data_Feed_Files where " + query) >= 3)
                {
                    query = "update reconciliation_batch set status = 'Scheduled',Is_KnetST_Parsed=1,is_host_parsed=1,is_switch_parsed=1 where transaction_start_date = convert(datetime, '" +
                        DateTime.Now.AddDays(-1).ToString("dd/MM/yyyy") + "',103) and status like '%pending%'";
                    ConnectionFactory.ExecuteQuery(query);
                    LogableTask.LogMonoActivityTask("BatchId", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info,
                           "Current day batches scheduled for reconciliation");

                }
            }
            catch (Exception ex)
            {
                LogableTask.LogMonoActivityTask("UpdateReconBatchStatus", MethodBase.GetCurrentMethod(), TraceLevel.Info, ex);
            }

            finally
            {
                timerUpdateReconBatchStatus.Change(new TimeSpan(0, 10, 0), new TimeSpan(0, 0, 0, 0, -1));
            }
        }
        //update batch creation logic independant of host/switch files and mange empty or missing host/switch file issue : iyju for GBK 11DEC2022
        void DoWork(object state)
        {
            timer.Change(-1, -1);
            bool hasONUS_SwitchTransactionONLY = true;
            bool noONUS_SwitchTransactionONLY = true;
            try
            {
                if (!isInitialized || (DateTime.Now - appSettingLastLoadedAt).Days > 7)
                {
                    DoInitialize();
                    isInitialized = true;
                }
                XmlLogWriter.InitXmlLogWriter(String.Format("{0}\\CCMSReconciliationScheduleService_{1:yyMMMdd}.txt", appSettings.LogFilePath, DateTime.Now));
                LogableTask.LogMonoActivityTask("writeVersion", MethodBase.GetCurrentMethod(), TraceLevel.Info, "Version: Reconciliation Schedular 6.2.0, Build date : 11/08/2026");
                //List<string> trxDateforKnetParse = new List<string>();
                List<int> batchIDforKnetParse = new List<int>();




                //TimeSpan _time = TimeSpan.Parse(recon.ServiceRunTime);
                //var date = DateTime.Now;
                //var _current = TimeSpan.Parse(date.ToString("HH:mm:ss"));
                //var temp = new TimeSpan(24, 0, 0) + _time - _current;
                //LogableTask.LogMonoActivityTask("HoursRemaining", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, " Hours Remaining: " + temp.ToString());
                //                DoWork_Ver2();
                //if (minuteSchedule != "1")
                //{
                //    //timer = new Timer(SendEmail, null, temp, new TimeSpan(0, 24, 0, 0));
                //    timer = new Timer(DoWork_Ver2, null, temp, new TimeSpan(0, 24, 0, 0));
                //    //timer = new Timer(DoWork, null, temp, new TimeSpan(0, 24, 0, 0));
                //}
                //else
                //{
                //    //timer = new Timer(SendEmail, null, new TimeSpan(0, 1, 0), new TimeSpan(0, 0, 5, 0));
                //    timer = new Timer(DoWork_Ver2, null, new TimeSpan(1, 0, 0), new TimeSpan(0, 0, 1, 0));//360
                //    //  timer = new Timer(DoWork, null, new TimeSpan(0, 1, 0), new TimeSpan(0, 0, 5, 0));
                //}

                //               EventLog.WriteEntry("ReconciliationScheduleService", "Service Started Successfully", EventLogEntryType.Information);
                //catch (Exception ex)
                //{
                //    //trying to log error in event log if its not full.
                //    try
                //    {
                //        EventLog.WriteEntry("ReconciliationSchedular", ex.Message + " " + ex.StackTrace, EventLogEntryType.Error);
                //    }
                //    catch (Exception innerException)
                //    {
                //    }
                //}


                // Execute a query to get all the required Atm records


                // Added By Hamza to query Reconciliation Batch at Once, & use it where required. -- 05-Nov-2024



                CreateBatch();


                //Insert host/switch/KNET transactions to each BAtch based on Transaction DATE
                //files exist check by iyju,
                //string knetSTFileName = "ST";
                //string knetSTFilePath = recon.HostFilePath.Replace("host.csv", knetSTFileName + DateTime.Now.AddDays(-1).ToString("yyMMdd") + ".txt");

                //string knetDRVSTFileName = "DRV_ST";
                //string knetDRVSTFilePath = recon.HostFilePath.Replace("host.csv", knetDRVSTFileName + DateTime.Now.AddDays(-1).ToString("yyMMdd") + ".txt");

                //if (File.Exists(recon.SwitchFilePath) || File.Exists(recon.HostFilePath))
                //{


                //******************************** START ********************************


                bool isSwitchReadyforParse = false, isHostReadyforParse = false;
                List<int> batchIDforHostParse = new List<int>();
                List<int> batchIDforSwitchParse = new List<int>();

                LogableTask.LogMonoActivityTask("AccessFiles", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, $"Trying to access {recon.SwitchFilePath} and {recon.HostFilePath}");

                bool switchFilesExist = Directory.Exists(recon.SwitchFilePath) &&
                        Directory.EnumerateFiles(recon.SwitchFilePath).Any();

                bool hostFilesExist = Directory.Exists(recon.HostFilePath) &&
                                      Directory.EnumerateFiles(recon.HostFilePath).Any();
                
                LogableTask.LogMonoActivityTask("AccessFiles", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, $"Switch files exist: {switchFilesExist}, Host files exist: {hostFilesExist}");

                if (switchFilesExist || hostFilesExist)
                {

                    if (switchFilesExist)
                    {
                        foreach (string file in Directory.EnumerateFiles(recon.SwitchFilePath))
                        {
                            if (file.EndsWith(".xlsx"))
                            {
                                LogableTask.LogMonoActivityTask("ProcessSwitchFile", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, $"BEGIN - Processing SwitchFile: {file}");

                                var tempBatchIDforSwitchParse = new List<int>();

                                ProcessSwitchFile(file, ref hasONUS_SwitchTransactionONLY, ref noONUS_SwitchTransactionONLY, ref isSwitchReadyforParse, ref tempBatchIDforSwitchParse);

                                if (isSwitchReadyforParse)
                                {
                                    batchIDforSwitchParse = batchIDforSwitchParse.Union(tempBatchIDforSwitchParse).ToList();
                                }

                                LogableTask.LogMonoActivityTask("ProcessSwitchFile", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, $"END - Processing SwitchFile: {file}");
                            }
                        }
                        
                    }

                    if (hostFilesExist)
                    {
                        foreach (string file in Directory.EnumerateFiles(recon.HostFilePath))
                        {
                            if (file.EndsWith(".xlsx"))
                            {
                                LogableTask.LogMonoActivityTask("ProcessHostFile", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, $"BEGIN - Processing HostFile: {file}");

                                var tempBatchIDforHostParse = new List<int>();

                                ProcessHostFile(file, ref isHostReadyforParse, ref tempBatchIDforHostParse);

                                if (isHostReadyforParse)
                                {
                                    batchIDforHostParse = batchIDforHostParse.Union(tempBatchIDforHostParse).ToList();
                                }

                                LogableTask.LogMonoActivityTask("ProcessHostFile", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, $"END - Processing HostFile: {file}");
                            }
                        }
                    }

                    
                        ////rerun all batches on the date of any changed batch-- for including cardless and tellerx
                        //List<int> batchIDforParse = batchIDforSwitchParse.Union(batchIDforHostParse).Union(batchIDforKnetParse).ToList();
                        //List<string> batchDateforParse = trxDateforSwitchParse.Union(trxDateforHostParse).Union(trxDateforKnetParse).ToList();
                        //if (batchDateforParse.Count() > 0)
                        //{
                        //    List<string> batchDates = new List<string>();
                        //    foreach (int _batchID in batchIDforParse)
                        //    {
                        //        ReconciliationTransactions.DeleteReconciliationTransactionss(" reconciliation_batch_id=" + _batchID);
                        //        ReconciledTransactions.DeleteReconciledTransactionss(" batch_id=" + _batchID);

                        //        ReconciliationBatch _batch = ReconciliationBatch.LoadReconciliationBatchByPk(_batchID);
                        //        if (batchIDforSwitchParse.Contains(_batchID) && isSwitchReadyforParse) _batch.IsSwitchParsed = 1;
                        //        if (batchIDforHostParse.Contains(_batchID) && isHostReadyforParse) _batch.IsHostParsed = 1;
                        //        if ((batchIDforKnetParse.Contains(_batchID) && isKnetSTReadyforParse) || (isSwitchReadyforParse && hasONUS_SwitchTransactionONLY)) _batch.IsKnetSTParsed = 1;
                        //        _batch.Status = "Scheduled";
                        //        _batch.Save();
                        //        LogableTask.LogMonoActivityTask("BatchId", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "Rerun-BatchId = " + _batchID);
                        //    }

                        //    foreach (string _batchDate in batchDateforParse)
                        //    {
                        //        RerunAllBatches(_batchDate);
                        //    }
                        //}
                        bool stateChanged = false;
                        //rerun only the changed batches
                        List<int> batchIDforParse = batchIDforSwitchParse.Union(batchIDforHostParse).Union(batchIDforKnetParse).ToList();
                        foreach (int _batchID in batchIDforParse)
                        {

                            ReconciliationBatch _batch = ReconciliationBatch.LoadReconciliationBatchByPk(_batchID);
                            if (_batch.BatchDescription == null)
                            {
                                ReconciliationTransactions.DeleteReconciliationTransactionss(" reconciliation_batch_id=" + _batchID);
                                ReconciledTransactions.DeleteReconciledTransactionss(" batch_id=" + _batchID);
                            }

                            if (batchIDforSwitchParse.Contains(_batchID))
                            {
                                _batch.Status = "Pending Host";
                                stateChanged = true;
                                _batch.IsSwitchParsed = 1;
                            }
                            if ((batchIDforHostParse.Contains(_batchID)) || (isSwitchReadyforParse && noONUS_SwitchTransactionONLY))
                            {
                                _batch.Status = "Scheduled";
                                stateChanged = true;
                                _batch.IsHostParsed = 1;
                            }


                            if (_batch.ReprocessCount == null)
                                _batch.ReprocessCount = 1;
                            else
                                _batch.ReprocessCount++;

                            if (_batch.TransactionStartDate < DateTime.Now.AddDays(-1))
                            {
                                _batch.Status = "Scheduled";
                                LogableTask.LogMonoActivityTask("BatchId", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info,
                                    "BatchId = " + _batchID + ", Transaction Date=" + _batch.TransactionStartDate + " rescheduled");
                            }
                            stateChanged = true;

                            if (stateChanged)
                            {
                                _batch.LastInvokedAt = DateTime.Now;
                                _batch.Save();
                            }
                                
                        }
                        //if (isKnetSTReadyforParse)
                        //{
                        //    //if (!isKnetSTReadyforParse)
                        //    //{//set knetst parsed as 1 when the knet st file is empty
                        //    // // atmData is a Dictionary -- Key = Title && Value == atmID
                        //    //    foreach (var atm in atmData)
                        //    //    {
                        //    //        ReconciliationBatch batchRecon = ReconciliationBatch.LoadReconciliationBatch("transaction_start_date = '"
                        //    //                   + DateTime.Now.AddDays(-1).ToString("yyyy/MM/dd") + " 00:00:00.000'and atm_id=" + atm.Value);
                        //    //        if (batchRecon != null)
                        //    //        {
                        //    //            batchRecon.IsKnetSTParsed = 1;
                        //    //            batchRecon.Status = "Scheduled";
                        //    //            batchRecon.Save();
                        //    //            LogableTask.LogMonoActivityTask("BatchId", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "Rerun-BatchId = " + batchRecon.ReconciliationBatchId);
                        //    //        }
                        //    //    }
                        //    //}
                        //    string query = "update reconciliation_batch set status = 'Scheduled',Is_KnetST_Parsed=1,is_host_parsed=1,is_switch_parsed=1 where transaction_start_date = convert(datetime, '" +
                        //       DateTime.Now.AddDays(-1).ToString("dd/MM/yyyy") + "',103)";
                        //    ConnectionFactory.ExecuteQuery(query);
                        //}
                        //Because all 
                        //for (int i = 0; i < listKNETProcessedDates.Count; i++)
                        //{
                        //    string query = "update reconciliation_batch set status = 'Scheduled' where transaction_start_date = convert(datetime, '" +
                        //        DateTime.ParseExact(listKNETProcessedDates[i], "yyyyMMdd", null).ToString("dd/MM/yyyy") + "',103)";
                        //    ConnectionFactory.ExecuteQuery(query);


                        //    LogableTask.LogMonoActivityTask("BatchId", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "executed:" + query);

                        //}





                        //try
                        //{
                        //    TimeSpan _time = TimeSpan.Parse(recon.ServiceRunTime);
                        //    var date = DateTime.Now;
                        //    var _current = TimeSpan.Parse(date.ToString("HH:mm:ss"));
                        //    var temp = new TimeSpan(24, 0, 0) + _time - _current;
                        //    timer = new Timer(SendEmail, null, temp, new TimeSpan(0, 24, 0, 0));
                        //}
                        //catch (Exception ex)
                        //{
                        //    LogableTask.LogMonoActivityTask("BatchId", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "Send email failed = " + ex.Message);
                        //}
                    



                    // LogableTask.LogMonoActivityTask("BatchId", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "Rerun-BatchId = " + string.Join(", ", batchID));
                    //END *************************

                }
                else
                {
                    LogableTask.LogMonoActivityTask("Switch_HostFileAbsence", MethodBase.GetCurrentMethod(), TraceLevel.Info, "Switch and Host is missing");
                }

                //if (!_isFilesinUSE)
                //{
                //    _isFilesinUSE = true;
                //    RemoveRawDuplicate();
                //    _isFilesinUSE = false;
                //}
            }
            catch (Exception ex)
            {
                LogableTask.LogMonoActivityTask("Exception", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Error, ex);
                if (ex.InnerException != null)
                {
                    LogableTask.LogMonoActivityTask("Exception", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Error," Inner: " + ex.InnerException);
                }
                //EventLog.WriteEntry("CCMSSchedular", "Error in DoWork_Ver2(). detail: " + ex.Message + ex.StackTrace, EventLogEntryType.Error);
            }
            finally
            {
                LogableTask.LogMonoActivityTask("Dowork", MethodBase.GetCurrentMethod(), TraceLevel.Info, "Going to sleep for 5 min");

                timer = new Timer(DoWork, null, new TimeSpan(0, 5, 0), new TimeSpan(0, 0, 0, 0, -1));
            }


        }

        protected static DataSet GetDataSetFromSheet(string filepath, string sheetName)
        {
            OleDbCommand command = null;
            OleDbDataAdapter adapter = null;
            DataSet dsSheet = null;

            try
            {
                if (filepath.Contains("xlsx"))
                    command = new OleDbCommand(@"select * from [" + sheetName + "$]", new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filepath + ";Extended Properties=\"Excel 12.0;IMEX=1;HDR=YES;\""));
                else
                    command = new OleDbCommand(@"select *  from [" + sheetName + "$]", new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filepath + ";Extended Properties=\"Excel 8.0;\""));
                adapter = new OleDbDataAdapter(command);
                dsSheet = new DataSet();
                adapter.Fill(dsSheet);
            }
            catch (Exception ex)
            {
                throw new Exception("Unable to read data from excel sheet. " + ex.Message);
            }
            finally
            {
                if (adapter != null)
                    adapter.Dispose();


            }
            return dsSheet;
        }






        private void ProcessHostFile(string file, ref bool isHostReadyforParse, ref List<int> batchIDforHostParse)
        {
            List<ReconciliationHostData> listReconHostData = new List<ReconciliationHostData>();
            List<ReconciliationBnaHostData> listReconBnaHostData = new List<ReconciliationBnaHostData>();
            ReconciliationDataFeedFiles reconciliationDataFeedFiles = null;
            SqlCommand cmd = null;
            SqlTransaction sqlHostTrxn = null;
            try
            {
                if (File.Exists(file))
                {
                    LogableTask.LogMonoActivityTask("HostFileExists", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "Host File Exists");

                    reconciliationDataFeedFiles = PopulateImportFileInfo("Host", file);
                    if (reconciliationDataFeedFiles.Status == "Scheduled")
                    {
                        DataTable dt = ExcelToDataTable(file, null, 8);

                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            DataRow dr = dt.Rows[i];

                            string batchAtmTitle = "QNB"+dr[15].ToString().Substring(5);

                            if (atmData.ContainsKey(batchAtmTitle))
                            {
                                int batchAtmID = atmData[batchAtmTitle];

                                DateTime trxnDate = DateTime.ParseExact(dr[4].ToString().Trim(), "dd/MM/yy HH:mm:ss", null);

                                string key = trxnDate.ToString("yyyy/MM/dd") + "_" + batchAtmID;
                                dynamic batchRecon = null;
                                if (ReconBatchDict.ContainsKey(key))
                                {
                                    batchRecon = ReconBatchDict[key];
                                }
                                else
                                {

                                    batchRecon = ReconciliationBatch.LoadReconciliationBatch("transaction_start_date = '"
                                   + trxnDate.ToString("yyyy-MM-dd HH:mm:ss") + "'and atm_id=" + batchAtmID);
                                    if (batchRecon != null)
                                    {
                                        InsertIntoReconciliationBatchDictionary(batchRecon.TransactionStartDate, batchRecon.AtmId, batchRecon.ReconciliationBatchId);
                                    }
                                }
                                if (batchRecon == null)
                                {
                                    batchRecon = CreateBatch(trxnDate, batchAtmID);
                                    if (batchRecon != null)
                                        InsertIntoReconciliationBatchDictionary(batchRecon.TransactionStartDate, batchRecon.AtmId, batchRecon.ReconciliationBatchId);

                                }
                                if (batchRecon != null)
                                {
                                    //TERMID	LOCAL_DATE	LOCAL_TIME	Pan	ACCTNUM	TRACE	MSG	PCODE	AMOUNT	CoreBank_Resp	ACQUIRER	REFNUM

                                    ReconciliationHostData hostData = new ReconciliationHostData();
                                    hostData.ReconciliationBatchId = batchRecon.ReconciliationBatchId;

                                    hostData.AtmId = batchAtmTitle;
                                    hostData.TransactionDate = trxnDate.Date.ToString("dd/MM/yyyy hh:mm:ss tt", CultureInfo.InvariantCulture);
                                    hostData.TransactionTime = ((int)trxnDate.TimeOfDay.TotalSeconds).ToString();
                                    hostData.CardNumber = dr[18].ToString();
                                    hostData.CustomerAccountNo = dr[20].ToString();
                                    hostData.TransactionSequence = dr[21].ToString();
                                    hostData.TransactionAmount = decimal.Parse(dr[19].ToString());
                                    hostData.TransactionResponse = ""; //TODO: UPDATE THIS WHEN STATUS COLUMN AVAILABLE IN EXCEL
                                    hostData.TransactionType = dr[13].ToString();//TODO: Need proper transaction_type for transactions from bank.

                                    hostData.IsKnet = false;

                                    //hostData.TransactionDatetime = DateTime.ParseExact(subParts[3].Trim() + subParts[4], "yyyyMMddHH:mm", null);
                                    hostData.DbAtmId = atmData[hostData.AtmId];
                                    //hostData.InstitutionIdentifier = hostData.CardNumber.Substring(0, 6);
                                    //hostData.IsReversal = hostData.TransactionType.ToLower().Contains("correction") | hostData.TransactionType.ToLower().Contains("-r");


                                    // Since the host file provided is of Withdrawal only that is why adding to list without check
                                    listReconHostData.Add(hostData);
                                    if (!batchIDforHostParse.Contains(batchRecon.ReconciliationBatchId))
                                        batchIDforHostParse.Add(batchRecon.ReconciliationBatchId);


                                    if (i % 1000 == 0)
                                        LogableTask.LogMonoActivityTask("", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, " Host :" + hostData.TransactionType + ", " + hostData.AtmId + ", TSN: " + hostData.TransactionSequence);


                                    //if (!trxDateforHostParse.Contains(batchRecon.TransactionStartDate.ToString())) trxDateforHostParse.Add(batchRecon.TransactionStartDate.ToString());
                                }
                                //else
                                //{//create batch for this OLD transaction
                                //    CreateBatch(trxnDate, batchAtmID);
                                //}
                            }
                            else
                            {
                                LogableTask.LogMonoActivityTask("", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Error, $"ATM Does not Exist or ATM is InActive: Title {batchAtmTitle}");
                            }
                            //check if batch atm title matches the CSV data atm title

                            //for Remote-on us //achw for knet cash withdrawal    acho for knet driven atms

                        }
                        //File.WriteAllText(String.Format("{0}\\ReconcileSchedule_hostCounter.txt", appSettings.LogFilePath), i.ToString());





                        //string data1 = Encoding.ASCII.GetString(File.ReadAllBytes(recon.HostFilePath));
                        //string[] parts1 = data1.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
                        //foreach (string part in parts1)
                        //{
                        //    string[] subParts = part.Split(',');
                        //    ReconciliationHostData hostData = new ReconciliationHostData();
                        //    ReconciliationBnaHostData BNAhostData = new ReconciliationBnaHostData();

                        //}

                        if (listReconBnaHostData.Count > 0 || listReconHostData.Count > 0)
                        {
                            int counter = 0;
                            int timeout = 300;
                            while (true)
                            {
                                try
                                {
                                    cmd = ConnectionFactory.GetNewCommand(true);
                                    sqlHostTrxn = cmd.Connection.BeginTransaction();
                                    cmd.Transaction = sqlHostTrxn;
                                    cmd.CommandTimeout = timeout;
                                    //ReconciliationBnaHostData.BulkSave(listReconBnaHostData, sqlHostTrxn);
                                    ReconciliationHostData.BulkSave(listReconHostData, sqlHostTrxn, timeout);
                                    UpdateProcessedInfo(reconciliationDataFeedFiles, sqlHostTrxn);
                                    sqlHostTrxn.Commit(); break;
                                }
                                catch (Exception ex)
                                {
                                    counter++;
                                    LogableTask.LogMonoActivityTask("HostFileMoved", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Error, ex);
                                    if (cmd != null)
                                        if (cmd.Connection != null)
                                            cmd.Connection.Close();
                                    Thread.Sleep(10 * 1000);
                                    timeout += 300;
                                    LogableTask.LogMonoActivityTask("HostFileMoved", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, $"Reattempting to save Host File: {file} =>  after 10 seconds");
                                    if (counter > 9)
                                        throw;
                                }
                            }
                        }


                        // File.Delete(recon.BackupFolderPath + "\\hostfile" + trxnDate.AddDays(1).ToString("ddMMyyyy") + ".csv");
                        File.Move(file, recon.BackupFolderPath + "\\hostfile" + DateTime.Now.ToString("ddMMyyyy_HHmmss") + ".xlsx");
                        LogableTask.LogMonoActivityTask("HosthFileMoved", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, $"Host File: {file} =>  Moved");
                        isHostReadyforParse = true;
                    }
                    else
                    {
                        LogableTask.LogMonoActivityTask("HosthFileMoved", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, $"Host File: {file} =>  Already processed before");
                        if (File.Exists(file))
                        {
                            File.Move(file, recon.BackupFolderPath + "\\hostfile" + DateTime.Now.ToString("ddMMyyyy_HHmmss") + ".xlsx");
                            LogableTask.LogMonoActivityTask("HosthFileMoved", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, $"Host File: {file} =>  Moved");
                            isHostReadyforParse = true;
                        }
                    }


                }
            }
            catch (Exception ex)
            {
                if (sqlHostTrxn != null)
                    sqlHostTrxn.Rollback();
                if (reconciliationDataFeedFiles != null)
                    UpdateErrorInfo(reconciliationDataFeedFiles, ex.Message);
                isHostReadyforParse = false;

                LogableTask.LogMonoActivityTask("Exception", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Error, $"Failed to process Host : {file} => " + ex.ToString());
                EventLog.WriteEntry("CCMSSchedular", "Error in DoWork_Ver2(). Host file schedule detail: " + ex.Message + ex.StackTrace, EventLogEntryType.Error);
                throw;
            }
            finally
            {
                if (cmd != null)
                    if (cmd.Connection != null)
                        cmd.Connection.Close();
            }

            //   return isHostReadyforParse;
        }
        static void test(string filePath)
        {
            //string filePath = @"C:\path\to\your\file.xlsx";
            //DataTable dt = new DataTable();
            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1); // Access the first worksheet
                var range = worksheet.RangeUsed(); // Get the used range of cells

                foreach (var row in range.Rows())
                {
                    foreach (var cell in row.Cells())
                    {
                        Console.WriteLine(cell.Value); // Read cell value
                    }
                }
            }
        }

        private void ProcessSwitchFile(string file, ref bool hasONUS_SwitchTransactionONLY, ref bool noONUS_SwitchTransactionONLY, ref bool isSwitchReadyforParse, ref List<int> batchIDforSwitchParse)
        {
            List<ReconciliationSwitchData> listReconSwitchData = new List<ReconciliationSwitchData>();
            List<ReconciliationBnaSwitchData> listReconBNAswitchData = new List<ReconciliationBnaSwitchData>();
            ReconciliationDataFeedFiles reconciliationDataFeedFiles = null;
            SqlCommand cmd = null;
            SqlTransaction sqlTrxn = null;

            try
            {

                if (File.Exists(file))
                {
                    reconciliationDataFeedFiles = PopulateImportFileInfo("Switch", file);
                    if (reconciliationDataFeedFiles.Status == "Scheduled")
                    {
                        LogableTask.LogMonoActivityTask("SwitchFileExists", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, " Switch File Exists");

                        // measure execution time for both the excel to dt functions

                        //var sw1 = System.Diagnostics.Stopwatch.StartNew();
                        //DataTable dt = GetDataTableFromExcel(recon.SwitchFilePath, "ATM All Data");
                        //sw1.Stop();
                        //Console.WriteLine($"Time taken convert 1: {sw1.ElapsedMilliseconds} ms");

                        //SAAD SOHAIL: EFFICIENT EXCEL TO DT FUNC
                        DataTable dt = ExcelToDataTable(file, "ATM All Data");

                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            DataRow dr = dt.Rows[i];

                            string batchAtmTitle = "QNB"+dr[0].ToString().Trim();

                            if (atmData.ContainsKey(batchAtmTitle))
                            {
                                int batchAtmID = atmData[batchAtmTitle];
                                //DateTime trxnDate = DateTime.ParseExact(dr[1].ToString(), "dd/MM/yyyy hh:mm:ss tt", null);

                                //SAAD SOHAIL: PARSING NUMERIC DATETIME FOR QNB
                                string datePart = dr[2].ToString().PadLeft(6, '0');
                                string timePart = dr[3].ToString().PadLeft(8, '0');

                                DateTime trxnDate = DateTime.ParseExact(
                                    datePart + timePart,
                                    "yyMMddHHmmssff",
                                    CultureInfo.InvariantCulture
                                );


                                string key = trxnDate.ToString("yyyy/MM/dd") + "_" + batchAtmID;
                                dynamic batchRecon = null;
                                if (ReconBatchDict.ContainsKey(key))
                                {
                                    batchRecon = ReconBatchDict[key];
                                }
                                else
                                {
                                    batchRecon = ReconciliationBatch.LoadReconciliationBatch("transaction_start_date = '"
                                       + trxnDate.ToString("yyyy-MM-dd") + "'and atm_id=" + batchAtmID);
                                    if (batchRecon != null)
                                    {
                                        InsertIntoReconciliationBatchDictionary(batchRecon.TransactionStartDate, batchRecon.AtmId, batchRecon.ReconciliationBatchId);
                                    }
                                }
                                if (batchRecon == null)
                                {
                                    batchRecon = CreateBatch(trxnDate, batchAtmID);
                                    if (batchRecon != null)
                                        InsertIntoReconciliationBatchDictionary(batchRecon.TransactionStartDate, batchRecon.AtmId, batchRecon.ReconciliationBatchId);
                                }

                                if (batchRecon != null)
                                {
                                    if (dr[11].ToString().Trim() == "Withdrawal")
                                    {
                                        ReconciliationSwitchData switchData = new ReconciliationSwitchData();

                                        //TERMID	LOCAL_DATE	LOCAL_TIME	PAN	ACCTNUM	TRACE	MSG	PCODE	AMOUNT	RESP	ACQUIRER	REFNUM	FEE	ACCEPTORNAM

                                        switchData.ReconciliationBatchId = batchRecon.ReconciliationBatchId;
                                        switchData.AtmId = batchAtmTitle;

                                        switchData.TransactionDate = trxnDate.Date.ToString("dd/MM/yyyy hh:mm:ss tt", CultureInfo.InvariantCulture);

                                        switchData.TransactionTime = ((int)trxnDate.TimeOfDay.TotalSeconds).ToString();

                                        switchData.CardNumber = dr[1].ToString().Trim();
                                        switchData.CustomerAccountNo = "";
                                        switchData.TransactionSequence = dr[8].ToString().Trim();// RRN for QNB
                                        switchData.TransactionAmount = decimal.Parse(dr[13].ToString());
                                        switchData.TransactionResponse = dr[14].ToString().Trim();
                                        switchData.TransactionCurrency = CurrencyCodes.NumericToAlphaCode.ContainsKey(int.TryParse(dr[15].ToString().Trim(), out int code) ? code : -1) ? CurrencyCodes.NumericToAlphaCode[code] : "";

                                        switchData.TransactionType = dr[11].ToString().Trim();
                                        switchData.TransactionSettlementDate = DateTime.ParseExact(
                                            dr[7].ToString().Trim(),
                                            "yyMMdd",
                                            CultureInfo.InvariantCulture
                                        ).Date;

                                        switchData.CardNetwork = "";
                                        switchData.CardIssuer = "";
                                        switchData.TransactionDatetime = trxnDate;
                                        switchData.InstitutionIdentifier = "";
                                        switchData.IsReversal = false;
                                        switchData.DbAtmId = batchAtmID;
                                        switchData.CardType = dr[18].ToString().Trim();
                                        /*Row["card_number"] = tran.CardNumber;
                                        Row["customer_account_no"] = tran.CustomerAccountNo;
                                        Row["transaction_date"] = tran.TransactionDate;
                                        Row["transaction_time"] = tran.TransactionTime;
                                        Row["transaction_amount"] = tran.TransactionAmount;
                                        Row["transaction_currency"] = tran.TransactionCurrency;
                                        Row["transaction_sequence"] = tran.TransactionSequence;
                                        Row["transaction_type"] = tran.TransactionType;
                                        Row["transaction_response"] = tran.TransactionResponse;
                                        Row["transaction_settlement_date"] = tran.TransactionSettlementDate;
                                        Row["card_network"] = tran.CardNetwork;
                                        Row["card_issuer"] = tran.CardIssuer;
                                        Row["transaction_datetime"] = tran.TransactionDatetime;
                                        Row["institution_identifier"] = tran.InstitutionIdentifier;
                                        Row["is_reversal"] = tran.IsReversal;
                                        Row["db_atm_id"] = tran.DbAtmId;
                                        Row["card_type"] = tran.CardType;
                                         */
                                        //switchData.TransactionDatetime = DateTime.ParseExact(subParts[3].Trim() + subParts[4], "yyyyMMddHH:mm", null);
                                        //switchData.DbAtmId = batchAtmID;
                                        //switchData.InstitutionIdentifier = switchData.CardNumber.Substring(0, 6);
                                        //switchData.IsReversal = switchData.TransactionType.ToLower().Contains("correction") | switchData.TransactionType.ToLower().Contains("-r");
                                        //if (switchData.CardIssuer != Customer_FIT_BankName)
                                        //    hasONUS_SwitchTransactionONLY = false;
                                        //else
                                        //    noONUS_SwitchTransactionONLY = false;

                                        //switchData.Save(sqlTrxn.Connection, sqlTrxn); //save only if new entry


                                        listReconSwitchData.Add(switchData);

                                        if (!batchIDforSwitchParse.Contains(batchRecon.ReconciliationBatchId))
                                            batchIDforSwitchParse.Add(batchRecon.ReconciliationBatchId);
                                    }
                                    else if (dr[11].ToString().Trim() == "Deposit")
                                    {
                                        ReconciliationBnaSwitchData switchBnaData = new ReconciliationBnaSwitchData();

                                        switchBnaData.ReconciliationBatchId = batchRecon.ReconciliationBatchId;
                                        switchBnaData.AtmId = batchAtmTitle;

                                        switchBnaData.TransactionDate = trxnDate.Date.ToString("dd/MM/yyyy hh:mm:ss tt", CultureInfo.InvariantCulture);

                                        switchBnaData.TransactionTime = ((int)trxnDate.TimeOfDay.TotalSeconds).ToString();
                                        switchBnaData.CardNumber = dr[1].ToString().Trim();
                                        switchBnaData.CustomerAccountNo = "";
                                        switchBnaData.TransactionSequence = dr[8].ToString().Trim();// RRN for QNB
                                        switchBnaData.TransactionAmount = decimal.Parse(dr[13].ToString());
                                        switchBnaData.TransactionResponse = dr[14].ToString().Trim();
                                        switchBnaData.TransactionCurrency = CurrencyCodes.NumericToAlphaCode.ContainsKey(int.TryParse(dr[15].ToString().Trim(), out int code) ? code : -1) ? CurrencyCodes.NumericToAlphaCode[code] : "";

                                        switchBnaData.TransactionType = dr[11].ToString().Trim();
                                        switchBnaData.TransactionSettlementDate = DateTime.ParseExact(
                                            dr[7].ToString().Trim(),
                                            "yyMMdd",
                                            CultureInfo.InvariantCulture
                                        ).Date;

                                        switchBnaData.CardNetwork = "";
                                        switchBnaData.CardIssuer = "";
                                        switchBnaData.TransactionDatetime = trxnDate;
                                        switchBnaData.InstitutionIdentifier = "";
                                        switchBnaData.IsReversal = false;
                                        switchBnaData.DbAtmId = batchAtmID;
                                        switchBnaData.CardType = dr[18].ToString().Trim();

                                        listReconBNAswitchData.Add(switchBnaData);

                                        if (!batchIDforSwitchParse.Contains(batchRecon.ReconciliationBatchId))
                                            batchIDforSwitchParse.Add(batchRecon.ReconciliationBatchId);
                                    }


                                    if (i % 1000 == 0)
                                        LogableTask.LogMonoActivityTask("", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, " Switch : " + dr[11].ToString().Trim() + ", " + batchAtmTitle + ", TSN: " + dr[8].ToString().Trim());




                                    //if (!trxDateforSwitchParse.Contains(batchRecon.TransactionStartDate.ToString())) trxDateforSwitchParse.Add(batchRecon.TransactionStartDate.ToString());
                                }
                                //else
                                //{//create batch for this OLD transaction
                                //    CreateBatch(trxnDate, batchAtmID);
                                //}
                            }
                            else
                            {
                                LogableTask.LogMonoActivityTask("", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Error, $"ATM Does not Exist or ATM is InActive: Title {batchAtmTitle}");
                            }
                            //}

                            //  isSwitchHeader = false;
                            //}
                            //}
                        }

                        if (listReconSwitchData.Count > 0 || listReconBNAswitchData.Count > 0)
                        {
                            int counter = 0;
                            int timeout = 300;
                            while (true)
                            {
                                try
                                {
                                    cmd = ConnectionFactory.GetNewCommand(true);
                                    sqlTrxn = cmd.Connection.BeginTransaction();
                                    cmd.Transaction = sqlTrxn;
                                    cmd.CommandTimeout = timeout;
                                    ReconciliationSwitchData.BulkSave(listReconSwitchData, sqlTrxn);
                                    ReconciliationBnaSwitchData.BulkSave(listReconBNAswitchData, sqlTrxn);
                                    UpdateProcessedInfo(reconciliationDataFeedFiles, sqlTrxn);
                                    sqlTrxn.Commit();
                                    break;
                                }
                                catch (Exception ex)
                                {
                                    counter++;
                                    LogableTask.LogMonoActivityTask("SwitchFileMoved", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Error, ex);
                                    if (cmd != null)
                                        if (cmd.Connection != null)
                                            cmd.Connection.Close();
                                    Thread.Sleep(10 * 1000);
                                    timeout += 300;
                                    LogableTask.LogMonoActivityTask("SwitchFileMoved", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, $"Reattempting to save Switch File: {file} =>  after 10 seconds");
                                    if (counter > 9)
                                        throw;
                                }
                            }

                        }
                        //File.Delete(recon.BackupFolderPath + "\\switchfile" + trxnDate.AddDays(1).ToString("ddMMyyyy") + ".csv");
                        File.Move(file, recon.BackupFolderPath + "\\switchfile" + DateTime.Now.ToString("ddMMyyyy_HHmmss") + ".xlsx");
                        LogableTask.LogMonoActivityTask("SwitchFileMoved", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, $"Switch File: {file} =>  Moved");
                        isSwitchReadyforParse = true;
                    }
                    else
                    {
                        LogableTask.LogMonoActivityTask("SwitchFileMoved", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, $"Switch File: {file} => Already processed before");
                        if (File.Exists(file))
                        {
                            File.Move(file, recon.BackupFolderPath + "\\switchfile" + DateTime.Now.ToString("ddMMyyyy_HHmmss") + ".xlsx");

                            LogableTask.LogMonoActivityTask("SwitchFileMoved", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, $"Switch File: {file} =>  Moved");
                            isSwitchReadyforParse = true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (sqlTrxn != null)
                    sqlTrxn.Rollback();
                if (reconciliationDataFeedFiles != null)
                    UpdateErrorInfo(reconciliationDataFeedFiles, ex.Message);

                isSwitchReadyforParse = false;
                LogableTask.LogMonoActivityTask("Exception", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Error, $"Failed to process Switch : {file} => " + ex.ToString());
                //EventLog.WriteEntry("CCMSSchedular", "Error in DoWork_Ver2(). Switch file schedule detail: " + ex.Message + ex.StackTrace, EventLogEntryType.Error);
                throw;
            }
            finally
            {
                if (cmd != null)
                    if (cmd.Connection != null)
                        cmd.Connection.Close();
            }
        }

        private static void UpdateProcessedInfo(ReconciliationDataFeedFiles reconciliationDataFeedFiles, SqlTransaction sqlHostTrxn)
        {
            reconciliationDataFeedFiles.ProcessedDatetime = DateTime.Now;
            reconciliationDataFeedFiles.Status = "Processed";
            reconciliationDataFeedFiles.Save(sqlHostTrxn.Connection, sqlHostTrxn);
        }
        private static void UpdateErrorInfo(ReconciliationDataFeedFiles reconciliationDataFeedFiles, string msg)
        {
            reconciliationDataFeedFiles.FailureReason = msg;
            reconciliationDataFeedFiles.Status = "Scheduled";
            reconciliationDataFeedFiles.RetryCount++;
            reconciliationDataFeedFiles.Save();
        }

        private static ReconciliationDataFeedFiles PopulateImportFileInfo(string fileType, string filePath)
        {
            string query = string.Format("creation_time>=convert(datetime,'{0}',103) and creation_time<=convert(datetime,'{0} 23:59:59',103) and status = 'Scheduled' and file_type='" + fileType + "'"
                        , DateTime.Today.ToString("dd/MM/yyyy"));

            ReconciliationDataFeedFiles reconciliationDataFeedFile = ReconciliationDataFeedFiles.LoadReconciliationDataFeedFiles(query);
            if (reconciliationDataFeedFile == null)
            {
                reconciliationDataFeedFile = new ReconciliationDataFeedFiles();
                reconciliationDataFeedFile.LastInvoked = DateTime.Now;
                reconciliationDataFeedFile.FileType = fileType;
                reconciliationDataFeedFile.CreationTime = DateTime.Now;
                reconciliationDataFeedFile.ProcessedDatetime = null;
                reconciliationDataFeedFile.RetryCount = 0;
                reconciliationDataFeedFile.SourceFilePath = filePath;
                reconciliationDataFeedFile.Status = "Scheduled";
                reconciliationDataFeedFile.Save();
            }
            else
            {
                if (reconciliationDataFeedFile.Status != "Processed")
                {
                    reconciliationDataFeedFile.Status = "Scheduled";
                    reconciliationDataFeedFile.LastInvoked = DateTime.Now;
                    reconciliationDataFeedFile.Save();
                }
            }
            return reconciliationDataFeedFile;
        }

        void RerunAllBatches(string startDate)
        {
            //ReconciliationBatch.ReconciliationBatchReader _reader = ReconciliationBatch.ExecuteReader(" transaction_start_date ='" + date + "' ");
            //while (_reader.Read())
            //{
            //    ReconciliationBatch _batch = ReconciliationBatch.LoadReconciliationBatchByPk(_reader.CurrentReconciliationBatch.ReconciliationBatchId);
            //    _batch.Status = "Scheduled";
            //    _batch.Save();
            //    LogableTask.LogMonoActivityTask("BatchId", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "Rerun-BatchId = " + _reader.CurrentReconciliationBatch.ReconciliationBatchId);

            //}



            //rerun all selected batches
            //"reExecute"
            LogableTask task = null;
            StringBuilder builder = new StringBuilder();
            try
            {
                task.Log(MethodBase.GetCurrentMethod(), TraceLevel.Info, "Batches for the TRX date :" + startDate);

                task = LogableTask.NewTask("GridView_Batch_RowCommand");

                SqlCommand cmd = ConnectionFactory.GetNewCommand(false);
                cmd.CommandTimeout = 500;

                cmd.CommandText = @"select reconciliation_batch_id from reconciliation_batch r where r.transaction_start_date = '" + startDate + "' ;";

                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                DataTable dtTrxn = new DataTable();
                adapter.Fill(dtTrxn);

                if (dtTrxn.Rows.Count > 0)
                {
                    foreach (DataRow row in dtTrxn.Rows)
                    {
                        int batchID = int.Parse(row["reconciliation_batch_id"].ToString());
                        builder.Append("delete reconciliation_transactions where reconciliation_batch_id=" + batchID + ";");

                        builder.Append(@"update reconciliation_gl_account
                                    set opening_balance = opening_balance+(
                                    select isnull(sum(opening_balance-closing_balance),0)
                                    from reconciliation_gl_account_summary
                                    where reconciliation_batch_id = " + batchID + ");");

                        builder.Append("delete reconciliation_gl_account_summary where reconciliation_batch_id=" + batchID + ";");
                        builder.Append("delete reconciled_transactions where batch_id=" + batchID + ";");

                        builder.Append("update reconciliation_batch set status = 'Scheduled',retry_count=10,auto_reconciled_amount=null,manual_reconciled_amount=null,dropped_transaction_amount = null,no_of_records_processed=0,no_of_records_reconciled=0,no_of_records_failed_to_reconciled=0 where reconciliation_batch_id=" + batchID + ";");
                    }
                    ConnectionFactory.ExecuteQuery(builder.ToString());
                    //   Literal_ok.Text = "Batch Rescheduled Successfully";
                    task.Log(MethodBase.GetCurrentMethod(), TraceLevel.Info, "Batches for the TRX date " + startDate + " Rescheduled Successfully");
                }
                //else
                //    Literal_ok.Text = "No Records Found!";

            }
            catch (Exception ex)
            {
                task.Log(MethodBase.GetCurrentMethod(), TraceLevel.Error, ex);
                // Literal_error.Text = "<li>An Error occured while exporting transactions to XML File. " + ex.Message + "</li>";

            }
            finally
            {

                task.EndTask();

            }
        }
        public static void ManageTx(List<ReconciliationBatch> list)
        {
            SqlCommand cmd = null;
            SqlTransaction SqlTrxn = null;

            try
            {
                cmd = ConnectionFactory.GetNewCommand(true);
                SqlTrxn = cmd.Connection.BeginTransaction();
                cmd.Transaction = SqlTrxn;
                ReconciliationBatch.BulkSave(list, SqlTrxn);
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
        //from iyju
        void CreateBatch()
        {
            try
            {
                bool isBatchCreated = false;
                LogableTask.LogMonoActivityTask("Current DateTime", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "Current DateTime: " + DateTime.Now.ToString());
                //File.AppendAllText(path, DateTime.Now +" Hello World from DoWork\n");
                Atm.AtmReader reader = Atm.ExecuteReader("is_active=1");
                StringBuilder sb = new StringBuilder();
                while (reader.Read())
                {
                    sb.Append(reader.CurrentAtm.ATMId + ",");
                }
                reader.Close();
                sb.Remove(sb.Length - 1, 1);
                //LogableTask.LogMonoActivityTask("ATMList", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "ATM: " + sb.ToString());
                string ATM_id = sb.ToString();
                ReconciliationBatch batch = null;
                List<int> batchID = new List<int>();
                if (ATM_id.Contains(","))
                {//multiple ATMS
                    string[] ATMs = ATM_id.Split(',');

                    for (int i = 0; i <= ATMs.Length - 1; i++)
                    {

                        //if (ATMs[i] == "1554903")
                        //{
                        //    break;
                        //}

                        ReconciliationBatch reconciliationBatch = ReconciliationBatch.LoadReconciliationBatch(" transaction_start_date ='" + DateTime.Now.AddDays(-1).ToString("yyyy/MM/dd") + "' and atm_id=" + ATMs[i].ToString());
                        if (reconciliationBatch == null)
                        {//create new batch

                            batch = new ReconciliationBatch();
                            batch.TransactionStartDate = DateTime.Today.AddDays(-1);
                            batch.TransactionEndDate = DateTime.Today.AddTicks(-1);
                            batch.AcceptableDifference = !string.IsNullOrEmpty(recon.Difference) ? Convert.ToDecimal(recon.Difference) : 0;
                            batch.CreationTime = DateTime.Now;
                            batch.CreatedBy = 1;
                            batch.RetryCount = 10;
                            batch.Status = "Pending Host/Switch";// Scheduled";
                            batch.AtmId = int.Parse(ATMs[i].ToString());
                            batch.AcceptableDifferenceType = recon.DifferenceType;
                            batch.IsHostParsed = 0;
                            batch.IsSwitchParsed = 0;
                            batch.IsKnetSTParsed = 0;
                            batch.LastInvokedAt = DateTime.Now;

                            //if (batch.AtmId == 1) // creating remote batch only for aTM with id ==1 and title = REM99999
                            //    batch.BatchDescription = "Remote";

                            batch.Save();
                            batchID.Add(batch.ReconciliationBatchId);

                            isBatchCreated = true;
                        }

                    }
                }
                else
                {//one atm
                    ReconciliationBatch reconciliationBatch = ReconciliationBatch.LoadReconciliationBatch(" transaction_start_date ='" + DateTime.Now.AddDays(-1).ToString("yyyy/MM/dd") + "' and atm_id=" + ATM_id);
                    if (reconciliationBatch == null)
                    {//create new batch

                        batch = new ReconciliationBatch();
                        batch.TransactionStartDate = DateTime.Today.AddDays(-1);
                        batch.TransactionEndDate = DateTime.Today.AddTicks(-1);
                        batch.AcceptableDifference = !string.IsNullOrEmpty(recon.Difference) ? Convert.ToDecimal(recon.Difference) : 0;
                        batch.CreationTime = DateTime.Now;
                        batch.CreatedBy = 1;
                        batch.RetryCount = 10;
                        batch.Status = "Pending Host/Switch";// Scheduled";
                        batch.AtmId = int.Parse(ATM_id);
                        batch.AcceptableDifferenceType = recon.DifferenceType;
                        batch.IsHostParsed = 0;
                        batch.IsSwitchParsed = 0;
                        batch.IsKnetSTParsed = 0;
                        batch.LastInvokedAt = DateTime.Now;

                        //batch.BatchDescription = "Remote";
                        batch.Save();
                        batchID.Add(batch.ReconciliationBatchId);

                        isBatchCreated = true;
                    }

                }

                if (isBatchCreated)
                {
                    LogableTask.DefaultTraceLevel = TraceLevel.Info;
                    LogableTask.LogMonoActivityTask("GetTraceLevel", MethodBase.GetCurrentMethod(), TraceLevel.Info, "Batches Created for all active atms");
                }

            }
            catch (Exception ex)
            {
                LogableTask.LogMonoActivityTask("Exception", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Error, ex.ToString());
                EventLog.WriteEntry("CCMSSchedular", "Error in CreateBatch(). detail: " + ex.Message + ex.StackTrace, EventLogEntryType.Error);
            }
        }

        ReconciliationBatch CreateBatch(DateTime trxnDate, int batchAtmID)
        {
            ReconciliationBatch batch = null;
            if (batchAtmID != 0)
            {
                batch = new ReconciliationBatch();
                batch.TransactionStartDate = trxnDate;
                batch.TransactionEndDate = trxnDate.AddDays(1);
                batch.AcceptableDifference = !string.IsNullOrEmpty(recon.Difference) ? Convert.ToDecimal(recon.Difference) : 0;
                batch.CreationTime = trxnDate.AddDays(1);//DateTime.Now;
                batch.CreatedBy = 1;
                batch.RetryCount = 10;
                batch.Status = "Pending Host/Switch";// Scheduled";
                batch.AtmId = (batchAtmID);
                batch.AcceptableDifferenceType = recon.DifferenceType;
                batch.IsHostParsed = 0;
                batch.IsSwitchParsed = 0;
                batch.IsKnetSTParsed = 0;
                batch.LastInvokedAt = DateTime.Now;
                batch.Save();
            }
            return batch;
        }
        //from iyju
        //void CreateBatch()
        //{
        //    List<ReconciliationBatch> listReconciliationBatch = new List<ReconciliationBatch>();
        //    ReconciliationBatch batch = null;
        //    List<int> batchID = new List<int>();
        //    int i = 0;
        //    foreach (var atm in atmData)
        //    {
        //        string atmTitle = atm.Key;    // The Title of the ATM
        //        int atmId = atm.Value;        // The corresponding ATMId
        //        string key = DateTime.Now.AddDays(-1).ToString("yyyy/MM/dd") + "_" + atmId;
        //        if (!ReconBatchDict.ContainsKey(key))
        //        {
        //            batch = new ReconciliationBatch();
        //            batch.TransactionStartDate = DateTime.Today.AddDays(-1);
        //            batch.TransactionEndDate = DateTime.Today.AddTicks(-1);
        //            batch.AcceptableDifference = !string.IsNullOrEmpty(recon.Difference) ? Convert.ToDecimal(recon.Difference) : 0;
        //            batch.CreationTime = DateTime.Now;
        //            batch.CreatedBy = 1;
        //            batch.RetryCount = 10;
        //            batch.Status = "Pending";// Scheduled";
        //            batch.AtmId = atmId;
        //            batch.AcceptableDifferenceType = recon.DifferenceType;
        //            batch.IsHostParsed = 0;
        //            batch.IsSwitchParsed = 0;
        //            batch.IsKnetSTParsed = 0;

        //            if (i == 0)
        //                batch.BatchDescription = "Remote";

        //            //batch.Save();
        //            listReconciliationBatch.Add(batch);
        //            InsertIntoReconciliationBatchDictionary(batch.TransactionStartDate, batch.AtmId, batch.ReconciliationBatchId);
        //            batchID.Add(batch.ReconciliationBatchId);
        //            i++;
        //        }
        //    }
        //    if (listReconciliationBatch.Count > 0)
        //    {
        //        ManageTx(listReconciliationBatch);
        //        LogableTask.LogMonoActivityTask("", MethodBase.GetCurrentMethod(), TraceLevel.Info, "Batch created with pending state:" + listReconciliationBatch.Count);
        //    }
        //    else
        //        LogableTask.LogMonoActivityTask("", MethodBase.GetCurrentMethod(), TraceLevel.Info, "Batch already exists for today");


        //}

        //void CreateBatch(DateTime trxnDate, int batchAtmID)
        //{
        //    try
        //    {
        //        if (batchAtmID != 0)
        //        {
        //            LogableTask.LogMonoActivityTask("Current DateTime", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "Current DateTime: " + DateTime.Now.ToString());
        //            //one atm
        //            ReconciliationBatch batch = null;
        //            string key = trxnDate.ToString("yyyy/MM/dd") + "_" + batchAtmID;
        //            dynamic batchRecon = null;
        //            if (ReconBatchDict.ContainsKey(key))
        //            {
        //                batchRecon = ReconBatchDict[key];
        //            }
        //            else
        //            {
        //                batchRecon = ReconciliationBatch.LoadReconciliationBatch(" transaction_start_date ='" + trxnDate.ToString("yyyy/MM/dd") + "' and atm_id=" + batchAtmID);
        //                if (batchRecon != null)
        //                {
        //                    InsertIntoReconciliationBatchDictionary(batchRecon.TransactionStartDate, batchRecon.AtmId, batchRecon.ReconciliationBatchId);
        //                }
        //            }
        //            if (batchRecon == null)
        //            {//create new batch

        //                batch = new ReconciliationBatch();
        //                batch.TransactionStartDate = trxnDate;
        //                batch.TransactionEndDate = trxnDate.AddDays(1);
        //                batch.AcceptableDifference = !string.IsNullOrEmpty(recon.Difference) ? Convert.ToDecimal(recon.Difference) : 0;
        //                batch.CreationTime = trxnDate.AddDays(1);
        //                batch.CreatedBy = 1;
        //                batch.RetryCount = 10;
        //                batch.Status = "Pending";// Scheduled";
        //                batch.AtmId = (batchAtmID);
        //                batch.AcceptableDifferenceType = recon.DifferenceType;
        //                batch.IsHostParsed = 0;
        //                batch.IsSwitchParsed = 0;
        //                batch.IsKnetSTParsed = 0;
        //                batch.Save();
        //                InsertIntoReconciliationBatchDictionary(batch.TransactionStartDate, batch.AtmId, batch.ReconciliationBatchId);
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        LogableTask.LogMonoActivityTask("Exception", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Error, ex.ToString());
        //        EventLog.WriteEntry("CCMSSchedular", "Error in CreateBatch(trxnDate,batchAtmID). detail: " + ex.Message + ex.StackTrace, EventLogEntryType.Error);
        //        throw;
        //    }
        //}

        //*************************************************
        //void DoWork(object state)
        //{
        //    try
        //    {
        //        XmlLogWriter.InitXmlLogWriter(String.Format("{0}\\ReconcileSchedule_{1:yyMMMdd}.txt", appSettings.LogFilePath, DateTime.Now));
        //        ThreadPool.SetMaxThreads(5, 5);

        //        string knetTapeFileName = "ST";
        //        string knetTapeFilePath = recon.HostFilePath.Replace("host.csv", knetTapeFileName + DateTime.Now.AddDays(-1).ToString("yyMMdd") + ".txt");

        //        //files exist check by iyju,
        //        if (File.Exists(recon.SwitchFilePath) && File.Exists(recon.HostFilePath) && File.Exists(knetTapeFilePath))
        //        {

        //            bool isValidFiles = true;
        //            DateTime trxnDate = DateTime.MinValue;
        //            ///
        //            string switchdata = Encoding.ASCII.GetString(File.ReadAllBytes(recon.SwitchFilePath));
        //            string[] switchparts = switchdata.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
        //            bool isSwitchHeader = true;
        //            foreach (string part in switchparts)
        //            {
        //                if (!isSwitchHeader)
        //                {
        //                    string[] subParts = part.Split(',');
        //                    if (subParts.Length > 3)
        //                    {
        //                        trxnDate = Convert.ToDateTime(subParts[3].Trim().Substring(0, 4) + "/" + subParts[3].Trim().Substring(4, 2) + "/" + subParts[3].Trim().Substring(6, 2));
        //                        DateTime yesterdayDate = Convert.ToDateTime(DateTime.Now.AddDays(-1).ToString("yyyy/MM/dd") + " 12:00:00AM");
        //                        if (trxnDate < yesterdayDate)
        //                        {

        //                            //Old data check by iyju, on request from GBK
        //                            isValidFiles = false;
        //                            break;
        //                        }
        //                    }
        //                }

        //                isSwitchHeader = false;
        //            }

        //            if (isValidFiles) //avoid checking if switch says old date
        //            {
        //                string hostdata = Encoding.ASCII.GetString(File.ReadAllBytes(recon.HostFilePath));
        //                string[] hostparts = hostdata.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
        //                foreach (string part in hostparts)
        //                {
        //                    string[] subParts = part.Split(',');
        //                    if (subParts.Length > 3)
        //                    {
        //                        trxnDate = Convert.ToDateTime(subParts[3].Trim().Substring(0, 4) + "/" + subParts[3].Trim().Substring(4, 2) + "/" + subParts[3].Trim().Substring(6, 2));
        //                        DateTime yesterdayDate = Convert.ToDateTime(DateTime.Now.AddDays(-1).ToString("yyyy/MM/dd") + " 12:00:00AM");
        //                        if (trxnDate < yesterdayDate)
        //                        {
        //                            //Old data check by iyju, on request from GBK
        //                            isValidFiles = false;
        //                            break;
        //                        }
        //                    }
        //                }
        //            }

        //            if (isValidFiles) //avoid checking if switch says old date
        //            {
        //                string data = Encoding.ASCII.GetString(File.ReadAllBytes(knetTapeFilePath));
        //                string[] parts = data.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
        //                foreach (string part in parts)
        //                {
        //                    if (part.Trim() != "")
        //                    {
        //                        string transactionDate = "20" + part.Substring(8, 6).Trim();
        //                        if (transactionDate.Trim() != DateTime.Now.AddDays(-1).ToString("yyyyMMdd").Trim())
        //                        {
        //                            //Old data check by iyju, on request from GBK
        //                            isValidFiles = false;
        //                            break;
        //                        }
        //                    }
        //                }
        //            }

        //            if (isValidFiles)
        //            {
        //                //check if batch already created even if on same date (today)
        //                //ReconciliationBatch _batch = ReconciliationBatch.LoadReconciliationBatch(" creation_time ='2022/11/21'")
        //            }
        //            if (isValidFiles)
        //            {
        //                try
        //                {
        //                    LogableTask.DefaultTraceLevel = (TraceLevel)Enum.Parse(typeof(TraceLevel), appSettings.ServiceLogLevel);
        //                }
        //                catch
        //                {
        //                    LogableTask.DefaultTraceLevel = TraceLevel.Info;
        //                    LogableTask.LogMonoActivityTask("GetTraceLevel", MethodBase.GetCurrentMethod(), TraceLevel.Error, "Failed to extract trace level from database");
        //                }
        //                LogableTask.LogMonoActivityTask("Current DateTime", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "Current DateTime: " + DateTime.Now.ToString());
        //                //File.AppendAllText(path, DateTime.Now +" Hello World from DoWork\n");
        //                Atm.AtmReader reader = Atm.ExecuteReader("is_active=1");
        //                StringBuilder sb = new StringBuilder();
        //                while (reader.Read())
        //                {
        //                    sb.Append(reader.CurrentAtm.ATMId + ",");
        //                }
        //                reader.Close();
        //                sb.Remove(sb.Length - 1, 1);
        //                LogableTask.LogMonoActivityTask("ATMList", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "ATM: " + sb.ToString());
        //                string ATM_id = sb.ToString();
        //                ReconciliationBatch batch = null;
        //                List<int> batchID = new List<int>();
        //                if (ATM_id.Contains(","))
        //                {
        //                    string[] ATMs = ATM_id.Split(',');

        //                    for (int i = 0; i <= ATMs.Length - 1; i++)
        //                    {
        //                        batch = new ReconciliationBatch();
        //                        batch.TransactionStartDate = DateTime.Today.AddDays(-1);
        //                        batch.TransactionEndDate = DateTime.Today.AddTicks(-1); // yahan se change start karna ha
        //                        batch.AcceptableDifference = !string.IsNullOrEmpty(recon.Difference) ? Convert.ToDecimal(recon.Difference) : 0;
        //                        batch.CreationTime = DateTime.Now;
        //                        batch.CreatedBy = 1;
        //                        batch.RetryCount = 10;
        //                        batch.Status = "Scheduled";
        //                        batch.AtmId = int.Parse(ATMs[i].ToString());
        //                        batch.AcceptableDifferenceType = recon.DifferenceType;
        //                        batch.Save();
        //                        batchID.Add(batch.ReconciliationBatchId);
        //                    }
        //                }
        //                else
        //                {
        //                    batch = new ReconciliationBatch();
        //                    batch.TransactionStartDate = DateTime.Today.AddDays(-1);
        //                    batch.TransactionEndDate = DateTime.Today.AddTicks(-1); // yahan se change start karna ha
        //                    batch.AcceptableDifference = !string.IsNullOrEmpty(recon.Difference) ? Convert.ToDecimal(recon.Difference) : 0;
        //                    batch.CreationTime = DateTime.Now;
        //                    batch.CreatedBy = 1;
        //                    batch.RetryCount = 10;
        //                    batch.Status = "Scheduled";
        //                    batch.AtmId = int.Parse(ATM_id);
        //                    batch.AcceptableDifferenceType = recon.DifferenceType;
        //                    batch.Save();
        //                    batchID.Add(batch.ReconciliationBatchId);
        //                }

        //                if (File.Exists(recon.SwitchFilePath))
        //                {
        //                    LogableTask.LogMonoActivityTask("SwitchFileExists", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "Switch File Exists");
        //                    foreach (int _batchID in batchID)
        //                    {
        //                        string data = Encoding.ASCII.GetString(File.ReadAllBytes(recon.SwitchFilePath));
        //                        //LogableTask.LogMonoActivityTask("", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "data =  " + data);
        //                        string[] parts = data.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
        //                        //LogableTask.LogMonoActivityTask("", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "Parts length =  " + parts.Length);

        //                        //Atm.AtmReader reader1 = Atm.ExecuteReader("ATM_id = 12");
        //                        //StringBuilder sb1 = new StringBuilder();
        //                        //while (reader1.Read())
        //                        //{
        //                        //    sb1.Append(reader1.CurrentAtm.ATMId + ",");
        //                        //}
        //                        //to get the atm title of the iterated batchid
        //                        string batchAtmTitle = "";
        //                        ReconciliationBatch.ReconciliationBatchReader batchRecon = ReconciliationBatch.ExecuteReader("reconciliation_batch_id = " + _batchID);
        //                        while (batchRecon.Read())
        //                        {
        //                            Atm.AtmReader _batchAtm = Atm.ExecuteReader("ATM_id = " + batchRecon.CurrentReconciliationBatch.AtmId);
        //                            while (_batchAtm.Read())
        //                            {
        //                                batchAtmTitle = _batchAtm.CurrentAtm.Title.Trim();
        //                            }
        //                        }


        //                        foreach (string part in parts)
        //                        {

        //                            string[] subParts = part.Split(',');
        //                            if (subParts[4].Length < 5)
        //                            {
        //                                subParts[4] = "0" + subParts[4];
        //                            }
        //                            ReconciliationSwitchData switchData = new ReconciliationSwitchData();
        //                            ReconciliationBnaSwitchData BNAswitchData = new ReconciliationBnaSwitchData();
        //                            //check if batch atm title matches the CSV data atm title
        //                            if (batchAtmTitle == subParts[0].Replace("?", "").Trim())
        //                            {
        //                                if (subParts[8].Contains("Withdrawal") || subParts[8].Contains("Cwd"))
        //                                {
        //                                    switchData.ReconciliationBatchId = _batchID;
        //                                    switchData.AtmId = subParts[0].Replace("?", "").Trim();
        //                                    switchData.CardNumber = subParts[1].Trim();
        //                                    switchData.CustomerAccountNo = subParts[2].Trim();
        //                                    switchData.TransactionDate = subParts[3].Trim();
        //                                    switchData.TransactionTime = subParts[4].Trim();
        //                                    switchData.TransactionAmount = decimal.Parse(subParts[5].Trim());
        //                                    switchData.TransactionCurrency = subParts[6].Trim();
        //                                    switchData.TransactionSequence = subParts[7].Trim();
        //                                    switchData.TransactionType = subParts[8].Trim();
        //                                    switchData.TransactionResponse = subParts[9].Trim();
        //                                    switchData.TransactionSettlementDate = DateTime.ParseExact(subParts[10].Trim(), "yyyyMMdd", null);
        //                                    switchData.CardNetwork = subParts[11].Trim();
        //                                    switchData.CardIssuer = subParts[12].Trim();
        //                                    switchData.Save();
        //                                }
        //                                else if (subParts[8].Contains("Deposit") && !subParts[8].Contains("Cheque"))
        //                                {
        //                                    BNAswitchData.ReconciliationBatchId = _batchID;
        //                                    BNAswitchData.AtmId = subParts[0].Trim();
        //                                    BNAswitchData.CardNumber = subParts[1];
        //                                    BNAswitchData.CustomerAccountNo = subParts[2];
        //                                    BNAswitchData.TransactionDate = subParts[3].Trim();
        //                                    BNAswitchData.TransactionTime = subParts[4];
        //                                    BNAswitchData.TransactionAmount = decimal.Parse(subParts[5]);
        //                                    BNAswitchData.TransactionCurrency = subParts[6];
        //                                    BNAswitchData.TransactionSequence = subParts[7];
        //                                    BNAswitchData.TransactionType = subParts[8];
        //                                    BNAswitchData.TransactionResponse = subParts[9];
        //                                    BNAswitchData.TransactionSettlementDate = DateTime.ParseExact(subParts[10], "yyyyMMdd", null);
        //                                    BNAswitchData.CardNetwork = subParts[11];
        //                                    BNAswitchData.CardIssuer = subParts[12];
        //                                    BNAswitchData.Save();
        //                                }
        //                                else
        //                                {
        //                                    LogableTask.LogMonoActivityTask("TransactionType", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "Transaction Type = " + subParts[8]);
        //                                }
        //                            }

        //                        }
        //                    }
        //                    File.Move(recon.SwitchFilePath, recon.BackupFolderPath + "\\switchfile" + DateTime.Now.ToString("ddMMyyyy") + ".csv");

        //                }

        //                if (File.Exists(recon.HostFilePath))
        //                {
        //                    LogableTask.LogMonoActivityTask("HostFileExists", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "Host File Exists");
        //                    foreach (int _batchID in batchID)
        //                    {
        //                        string data = Encoding.ASCII.GetString(File.ReadAllBytes(recon.HostFilePath));
        //                        string[] parts = data.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);

        //                        //to get the atm title of the iterated batchid
        //                        string batchAtmTitle = "";
        //                        ReconciliationBatch.ReconciliationBatchReader batchRecon = ReconciliationBatch.ExecuteReader("reconciliation_batch_id = " + _batchID);
        //                        while (batchRecon.Read())
        //                        {
        //                            Atm.AtmReader _batchAtm = Atm.ExecuteReader("ATM_id = " + batchRecon.CurrentReconciliationBatch.AtmId);
        //                            while (_batchAtm.Read())
        //                            {
        //                                batchAtmTitle = _batchAtm.CurrentAtm.Title.Trim();
        //                            }
        //                        }

        //                        foreach (string part in parts)
        //                        {
        //                            string[] subParts = part.Split(',');
        //                            ReconciliationHostData hostData = new ReconciliationHostData();
        //                            ReconciliationBnaHostData BNAhostData = new ReconciliationBnaHostData();
        //                            //check if batch atm title matches the CSV data atm title
        //                            if (batchAtmTitle == subParts[0].Trim())
        //                            {
        //                                if (withdrawalTran.ContainsKey(Convert.ToInt32(subParts[8])))
        //                                {
        //                                    hostData.ReconciliationBatchId = _batchID;
        //                                    hostData.AtmId = subParts[0];
        //                                    hostData.CardNumber = subParts[1];
        //                                    hostData.CustomerAccountNo = subParts[2];
        //                                    hostData.TransactionDate = subParts[3];
        //                                    hostData.TransactionTime = subParts[4];
        //                                    hostData.TransactionAmount = decimal.Parse(subParts[5]);
        //                                    hostData.TransactionCurrency = subParts[6];
        //                                    hostData.TransactionSequence = subParts[7];
        //                                    hostData.TransactionType = withdrawalTran[Convert.ToInt32(subParts[8])];
        //                                    hostData.TransactionResponse = subParts[9];
        //                                    hostData.TransactionSettlementDate = DateTime.ParseExact(subParts[10], "yyyyMMdd", null);
        //                                    hostData.CardNetwork = subParts[11];
        //                                    hostData.CardIssuer = subParts[12];
        //                                    hostData.CardAcquirer = Customer_FIT_BankName;// "GBK"; //always  GBK for ON US Withdrawal
        //                                    hostData.IsKnet = false;
        //                                    hostData.Save();
        //                                }
        //                                else if (depositTran.ContainsKey(Convert.ToInt32(subParts[8])))
        //                                {
        //                                    BNAhostData.ReconciliationBatchId = _batchID;
        //                                    BNAhostData.AtmId = subParts[0];
        //                                    BNAhostData.CardNumber = subParts[1];
        //                                    BNAhostData.CustomerAccountNo = subParts[2];
        //                                    BNAhostData.TransactionDate = subParts[3];
        //                                    BNAhostData.TransactionTime = subParts[4];
        //                                    BNAhostData.TransactionAmount = decimal.Parse(subParts[5]);
        //                                    BNAhostData.TransactionCurrency = subParts[6];
        //                                    BNAhostData.TransactionSequence = subParts[7];
        //                                    BNAhostData.TransactionType = depositTran[Convert.ToInt32(subParts[8])];
        //                                    BNAhostData.TransactionResponse = subParts[9];
        //                                    BNAhostData.TransactionSettlementDate = DateTime.ParseExact(subParts[10], "yyyyMMdd", null);
        //                                    BNAhostData.CardNetwork = subParts[11];
        //                                    BNAhostData.CardIssuer = subParts[12];
        //                                    BNAhostData.Save();
        //                                }
        //                            }
        //                        }
        //                    }

        //                    string data1 = Encoding.ASCII.GetString(File.ReadAllBytes(recon.HostFilePath));
        //                    string[] parts1 = data1.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
        //                    foreach (string part in parts1)
        //                    {
        //                        string[] subParts = part.Split(',');
        //                        ReconciliationHostData hostData = new ReconciliationHostData();
        //                        ReconciliationBnaHostData BNAhostData = new ReconciliationBnaHostData();
        //                        //check if batch atm title matches the CSV data atm title
        //                        if (Cardless_Tellerx_TerminalID == subParts[0].Trim())
        //                        {
        //                            if (withdrawalTran.ContainsKey(Convert.ToInt32(subParts[8])))
        //                            {
        //                                hostData.ReconciliationBatchId = Cardless_Tellerx_BatchID;
        //                                hostData.AtmId = subParts[0];
        //                                hostData.CardNumber = subParts[1];
        //                                hostData.CustomerAccountNo = subParts[2];
        //                                hostData.TransactionDate = subParts[3];
        //                                hostData.TransactionTime = subParts[4];
        //                                hostData.TransactionAmount = decimal.Parse(subParts[5]);
        //                                hostData.TransactionCurrency = subParts[6];
        //                                hostData.TransactionSequence = subParts[7];
        //                                hostData.TransactionType = withdrawalTran[Convert.ToInt32(subParts[8])];
        //                                hostData.TransactionResponse = subParts[9];
        //                                hostData.TransactionSettlementDate = DateTime.ParseExact(subParts[10], "yyyyMMdd", null);
        //                                hostData.CardNetwork = subParts[11];
        //                                hostData.CardIssuer = subParts[12];
        //                                hostData.IsKnet = false;
        //                                hostData.Save();
        //                            }
        //                            else if (depositTran.ContainsKey(Convert.ToInt32(subParts[8])))
        //                            {
        //                                BNAhostData.ReconciliationBatchId = Cardless_Tellerx_BatchID;
        //                                BNAhostData.AtmId = subParts[0];
        //                                BNAhostData.CardNumber = subParts[1];
        //                                BNAhostData.CustomerAccountNo = subParts[2];
        //                                BNAhostData.TransactionDate = subParts[3];
        //                                BNAhostData.TransactionTime = subParts[4];
        //                                BNAhostData.TransactionAmount = decimal.Parse(subParts[5]);
        //                                BNAhostData.TransactionCurrency = subParts[6];
        //                                BNAhostData.TransactionSequence = subParts[7];
        //                                BNAhostData.TransactionType = depositTran[Convert.ToInt32(subParts[8])];
        //                                BNAhostData.TransactionResponse = subParts[9];
        //                                BNAhostData.TransactionSettlementDate = DateTime.ParseExact(subParts[10], "yyyyMMdd", null);
        //                                BNAhostData.CardNetwork = subParts[11];
        //                                BNAhostData.CardIssuer = subParts[12];
        //                                BNAhostData.Save();
        //                            }
        //                        }
        //                    }


        //                    File.Move(recon.HostFilePath, recon.BackupFolderPath + "\\hostfile" + DateTime.Now.ToString("ddMMyyyy") + ".csv");
        //                }

        //                if (File.Exists(knetTapeFilePath))
        //                {
        //                    LogableTask.LogMonoActivityTask("STFileExists", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "ST File Exists");
        //                    foreach (int _batchID in batchID)
        //                    {

        //                        //to get the atm title of the iterated batchid
        //                        string batchAtmTitle = "";
        //                        ReconciliationBatch.ReconciliationBatchReader batchRecon = ReconciliationBatch.ExecuteReader("reconciliation_batch_id = " + _batchID);
        //                        while (batchRecon.Read())
        //                        {
        //                            Atm.AtmReader _batchAtm = Atm.ExecuteReader("ATM_id = " + batchRecon.CurrentReconciliationBatch.AtmId);
        //                            while (_batchAtm.Read())
        //                            {
        //                                batchAtmTitle = _batchAtm.CurrentAtm.Title.Trim();
        //                            }
        //                        }
        //                        string data = Encoding.ASCII.GetString(File.ReadAllBytes(knetTapeFilePath));
        //                        string[] parts = data.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
        //                        foreach (string part in parts)
        //                        {
        //                            if (part.Trim() != "")
        //                            {
        //                                string serviceType = part.Substring(0, 4).Trim();
        //                                string cardAcquirer = part.Substring(98, 4).Trim();
        //                                string cardIssuer = part.Substring(144, 4).Trim();
        //                                string transCode = part.Substring(40, 2).Trim();//transCode == "10" WITHDRAWAL
        //                                string terminalID = part.Substring(102, 16).Trim();

        //                                //Customer_FIT_BankName GBK
        //                                if (serviceType == "ATM" && cardAcquirer == Customer_FIT_BankName && cardIssuer != Customer_FIT_BankName && transCode == "10"
        //                                    && terminalID.Trim() == batchAtmTitle.Trim())
        //                                {
        //                                    string cardNo = part.Substring(148, 19).Trim().Replace("X", "*");
        //                                    string transactionDate = "20" + part.Substring(8, 6).Trim();
        //                                    string transactionTime = part.Substring(14, 2).Trim() + ":" + part.Substring(16, 2).Trim();
        //                                    string currency = "KWD";
        //                                    string amount = part.Substring(43, 11).Trim();
        //                                    string seqNo = part.Substring(28, 12).Trim();
        //                                    string transactionResponse = part.Substring(167, 2).Trim();
        //                                    string settlementDate = part.Substring(22, 6).Trim();
        //                                    string authType = part.Substring(118, 4).Trim();
        //                                    // int _batchID = 101;

        //                                    ReconciliationHostData hostData = new ReconciliationHostData();
        //                                    hostData.ReconciliationBatchId = _batchID;
        //                                    hostData.AtmId = terminalID;

        //                                    string starString = "";
        //                                    int starstringLen = cardNo.Trim().Length - 10;
        //                                    for (int sl = 0; sl < starstringLen; sl++)
        //                                    { starString = starString + "*"; }

        //                                    hostData.CardNumber = cardNo.Substring(0, 6) + starString;
        //                                    hostData.CardNumber = hostData.CardNumber + cardNo.Substring(hostData.CardNumber.Length, 4);

        //                                    // hostData.CardNumber = cardNo.Substring(0, 6) + "******" + cardNo.Substring(12, cardNo.Length - 12);//4 cardno can be 15/16 digit
        //                                    // hostData.CustomerAccountNo = subPar;
        //                                    hostData.TransactionDate = transactionDate;
        //                                    hostData.TransactionTime = transactionTime;
        //                                    hostData.TransactionAmount = decimal.Parse(amount) / 1000;
        //                                    hostData.TransactionCurrency = currency;
        //                                    hostData.TransactionSequence = seqNo;
        //                                    if (authType == "0210")
        //                                        hostData.TransactionType = withdrawalTran[10];  //cash withdrawal-KNET
        //                                    else if (authType == "0420") //reversal
        //                                        hostData.TransactionType = withdrawalTran[420];  //withdrawal-KNET-correction

        //                                    hostData.TransactionResponse = transactionResponse;
        //                                    hostData.TransactionSettlementDate = DateTime.ParseExact(settlementDate, "yyMMdd", null);
        //                                    // hostData.CardNetwork = subParts[11];
        //                                    hostData.CardIssuer = cardIssuer;
        //                                    hostData.CardAcquirer = cardAcquirer;
        //                                    hostData.IsKnet = true;
        //                                    hostData.Save();

        //                                    LogableTask.LogMonoActivityTask("", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "KNET TAPE :" + serviceType +
        //                                      ", " + terminalID + ", " + cardAcquirer + ", " + cardIssuer);
        //                                }
        //                            }
        //                        }

        //                    }
        //                    File.Move(knetTapeFilePath, recon.BackupFolderPath + "\\" + knetTapeFileName + "file" + DateTime.Now.ToString("ddMMyyyy") + ".txt");
        //                    //LogableTask.LogMonoActivityTask("Sleep", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "Thread going to sleep");

        //                }

        //                LogableTask.LogMonoActivityTask("BatchId", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "BatchId = " + string.Join(", ", batchID));

        //            }
        //            else
        //            {
        //                LogableTask.DefaultTraceLevel = TraceLevel.Info;
        //                LogableTask.LogMonoActivityTask("Switch_HostOldData", MethodBase.GetCurrentMethod(), TraceLevel.Error, "Switch or Host or Knet ST File has old data");

        //                if (trxnDate != DateTime.MinValue)
        //                {
        //                    //retrieve the batch id
        //                    StringBuilder batchIDs = new StringBuilder();
        //                    ReconciliationBatch.ReconciliationBatchReader _batchreader = ReconciliationBatch.ExecuteReader(" transaction_start_date = '" + trxnDate.ToString("yyyy-MM-dd") + " 00:00:00.000'");
        //                    while (_batchreader.Read())
        //                    {
        //                        batchIDs.Append(_batchreader.CurrentReconciliationBatch.ReconciliationBatchId.ToString() + ",");
        //                    }
        //                    if (batchIDs.ToString().Trim() != "")
        //                    {
        //                        batchIDs.Replace(',', ' ', batchIDs.Length - 1, 1);
        //                        LogableTask.LogMonoActivityTask("OldBatchIDs", MethodBase.GetCurrentMethod(), TraceLevel.Error, "Old Batch IDs :" + batchIDs.ToString());

        //                        //delete old batch id datas
        //                        ReconBatchInfo.DeleteReconBatchInfos(" Reconciliation_Batch_Id in (" + batchIDs.ToString() + ")");
        //                        ReconciliationHostData.DeleteReconciliationHostDatas(" Reconciliation_Batch_Id in (" + batchIDs.ToString() + ")");
        //                        ReconciliationBnaHostData.DeleteReconciliationBnaHostDatas(" Reconciliation_Batch_Id in (" + batchIDs.ToString() + ")");
        //                        ReconciliationSwitchData.DeleteReconciliationSwitchDatas(" Reconciliation_Batch_Id in (" + batchIDs.ToString() + ")");
        //                        ReconciliationBnaSwitchData.DeleteReconciliationBnaSwitchDatas(" Reconciliation_Batch_Id in (" + batchIDs.ToString() + ")");

        //                        //delete tellerx entries (fixed batch id :100)
        //                        ReconciliationHostData.DeleteReconciliationHostDatas(" Reconciliation_Batch_Id=100 and transaction_date='" + trxnDate.ToString("yyyyMMdd") + "'");
        //                        ReconciliationBnaHostData.DeleteReconciliationBnaHostDatas(" Reconciliation_Batch_Id=100 and transaction_date='" + trxnDate.ToString("yyyyMMdd") + "'");

        //                        ReconciliationBatch.DeleteReconciliationBatchs(" Reconciliation_Batch_Id in (" + batchIDs.ToString() + ")");

        //                        LogableTask.LogMonoActivityTask("OldBatchIDRemoved", MethodBase.GetCurrentMethod(), TraceLevel.Error, "Old Batch Removed :" + trxnDate.ToString("yyyyMMdd"));
        //                    }

        //                    //Insert Batch Data--------START ********************************

        //                    try
        //                    {
        //                        LogableTask.DefaultTraceLevel = (TraceLevel)Enum.Parse(typeof(TraceLevel), appSettings.ServiceLogLevel);
        //                    }
        //                    catch
        //                    {
        //                        LogableTask.DefaultTraceLevel = TraceLevel.Info;
        //                        LogableTask.LogMonoActivityTask("GetTraceLevel", MethodBase.GetCurrentMethod(), TraceLevel.Error, "Failed to extract trace level from database");
        //                    }
        //                    LogableTask.LogMonoActivityTask("Current DateTime", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "Rerun-Current DateTime: " + DateTime.Now.ToString());
        //                    //File.AppendAllText(path, DateTime.Now +" Hello World from DoWork\n");
        //                    Atm.AtmReader reader = Atm.ExecuteReader("is_active=1");
        //                    StringBuilder sb = new StringBuilder();
        //                    while (reader.Read())
        //                    {
        //                        sb.Append(reader.CurrentAtm.ATMId + ",");
        //                    }
        //                    reader.Close();
        //                    sb.Remove(sb.Length - 1, 1);
        //                    LogableTask.LogMonoActivityTask("ATMList", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "ATM: " + sb.ToString());
        //                    string ATM_id = sb.ToString();
        //                    ReconciliationBatch batch = null;
        //                    List<int> batchID = new List<int>();
        //                    if (ATM_id.Contains(","))
        //                    {
        //                        string[] ATMs = ATM_id.Split(',');

        //                        for (int i = 0; i <= ATMs.Length - 1; i++)
        //                        {
        //                            batch = new ReconciliationBatch();
        //                            batch.TransactionStartDate = trxnDate;
        //                            batch.TransactionEndDate = trxnDate.AddDays(1);
        //                            batch.AcceptableDifference = !string.IsNullOrEmpty(recon.Difference) ? Convert.ToDecimal(recon.Difference) : 0;
        //                            batch.CreationTime = trxnDate.AddDays(1);
        //                            batch.CreatedBy = 1;
        //                            batch.RetryCount = 10;
        //                            batch.Status = "Scheduled";
        //                            batch.AtmId = int.Parse(ATMs[i].ToString());
        //                            batch.AcceptableDifferenceType = recon.DifferenceType;
        //                            batch.Save();
        //                            batchID.Add(batch.ReconciliationBatchId);
        //                        }
        //                    }
        //                    else
        //                    {
        //                        batch = new ReconciliationBatch();
        //                        batch.TransactionStartDate = trxnDate;
        //                        batch.TransactionEndDate = trxnDate.AddDays(1);
        //                        batch.AcceptableDifference = !string.IsNullOrEmpty(recon.Difference) ? Convert.ToDecimal(recon.Difference) : 0;
        //                        batch.CreationTime = trxnDate;
        //                        batch.CreatedBy = 1;
        //                        batch.RetryCount = 10;
        //                        batch.Status = "Scheduled";
        //                        batch.AtmId = int.Parse(ATM_id);
        //                        batch.AcceptableDifferenceType = recon.DifferenceType;
        //                        batch.Save();
        //                        batchID.Add(batch.ReconciliationBatchId);
        //                    }

        //                    if (File.Exists(recon.SwitchFilePath))
        //                    {
        //                        LogableTask.LogMonoActivityTask("SwitchFileExists", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "Switch File Exists");
        //                        foreach (int _batchID in batchID)
        //                        {
        //                            string data = Encoding.ASCII.GetString(File.ReadAllBytes(recon.SwitchFilePath));
        //                            //LogableTask.LogMonoActivityTask("", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "data =  " + data);
        //                            string[] parts = data.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
        //                            //LogableTask.LogMonoActivityTask("", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "Parts length =  " + parts.Length);

        //                            //Atm.AtmReader reader1 = Atm.ExecuteReader("ATM_id = 12");
        //                            //StringBuilder sb1 = new StringBuilder();
        //                            //while (reader1.Read())
        //                            //{
        //                            //    sb1.Append(reader1.CurrentAtm.ATMId + ",");
        //                            //}
        //                            //to get the atm title of the iterated batchid
        //                            string batchAtmTitle = "";
        //                            ReconciliationBatch.ReconciliationBatchReader batchRecon = ReconciliationBatch.ExecuteReader("reconciliation_batch_id = " + _batchID);
        //                            while (batchRecon.Read())
        //                            {
        //                                Atm.AtmReader _batchAtm = Atm.ExecuteReader("ATM_id = " + batchRecon.CurrentReconciliationBatch.AtmId);
        //                                while (_batchAtm.Read())
        //                                {
        //                                    batchAtmTitle = _batchAtm.CurrentAtm.Title.Trim();
        //                                }
        //                            }


        //                            foreach (string part in parts)
        //                            {

        //                                string[] subParts = part.Split(',');
        //                                if (subParts[4].Length < 5)
        //                                {
        //                                    subParts[4] = "0" + subParts[4];
        //                                }
        //                                ReconciliationSwitchData switchData = new ReconciliationSwitchData();
        //                                ReconciliationBnaSwitchData BNAswitchData = new ReconciliationBnaSwitchData();
        //                                //check if batch atm title matches the CSV data atm title
        //                                if (batchAtmTitle == subParts[0].Replace("?", "").Trim())
        //                                {
        //                                    if (subParts[8].Contains("Withdrawal") || subParts[8].Contains("Cwd"))
        //                                    {
        //                                        switchData.ReconciliationBatchId = _batchID;
        //                                        switchData.AtmId = subParts[0].Replace("?", "").Trim();
        //                                        switchData.CardNumber = subParts[1].Trim();
        //                                        switchData.CustomerAccountNo = subParts[2].Trim();
        //                                        switchData.TransactionDate = subParts[3].Trim();
        //                                        switchData.TransactionTime = subParts[4].Trim();
        //                                        switchData.TransactionAmount = decimal.Parse(subParts[5].Trim());
        //                                        switchData.TransactionCurrency = subParts[6].Trim();
        //                                        switchData.TransactionSequence = subParts[7].Trim();
        //                                        switchData.TransactionType = subParts[8].Trim();
        //                                        switchData.TransactionResponse = subParts[9].Trim();
        //                                        switchData.TransactionSettlementDate = DateTime.ParseExact(subParts[10].Trim(), "yyyyMMdd", null);
        //                                        switchData.CardNetwork = subParts[11].Trim();
        //                                        switchData.CardIssuer = subParts[12].Trim();
        //                                        switchData.Save();
        //                                    }
        //                                    else if (subParts[8].Contains("Deposit") && !subParts[8].Contains("Cheque"))
        //                                    {
        //                                        BNAswitchData.ReconciliationBatchId = _batchID;
        //                                        BNAswitchData.AtmId = subParts[0].Trim();
        //                                        BNAswitchData.CardNumber = subParts[1];
        //                                        BNAswitchData.CustomerAccountNo = subParts[2];
        //                                        BNAswitchData.TransactionDate = subParts[3].Trim();
        //                                        BNAswitchData.TransactionTime = subParts[4];
        //                                        BNAswitchData.TransactionAmount = decimal.Parse(subParts[5]);
        //                                        BNAswitchData.TransactionCurrency = subParts[6];
        //                                        BNAswitchData.TransactionSequence = subParts[7];
        //                                        BNAswitchData.TransactionType = subParts[8];
        //                                        BNAswitchData.TransactionResponse = subParts[9];
        //                                        BNAswitchData.TransactionSettlementDate = DateTime.ParseExact(subParts[10], "yyyyMMdd", null);
        //                                        BNAswitchData.CardNetwork = subParts[11];
        //                                        BNAswitchData.CardIssuer = subParts[12];
        //                                        BNAswitchData.Save();
        //                                    }
        //                                    else
        //                                    {
        //                                        LogableTask.LogMonoActivityTask("TransactionType", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "Transaction Type = " + subParts[8]);
        //                                    }
        //                                }

        //                            }
        //                        }
        //                        File.Delete(recon.BackupFolderPath + "\\switchfile" + trxnDate.AddDays(1).ToString("ddMMyyyy") + ".csv");
        //                        File.Move(recon.SwitchFilePath, recon.BackupFolderPath + "\\switchfile" + trxnDate.AddDays(1).ToString("ddMMyyyy") + ".csv");

        //                    }

        //                    if (File.Exists(recon.HostFilePath))
        //                    {
        //                        LogableTask.LogMonoActivityTask("HostFileExists", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "Host File Exists");
        //                        foreach (int _batchID in batchID)
        //                        {
        //                            string data = Encoding.ASCII.GetString(File.ReadAllBytes(recon.HostFilePath));
        //                            string[] parts = data.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);

        //                            //to get the atm title of the iterated batchid
        //                            string batchAtmTitle = "";
        //                            ReconciliationBatch.ReconciliationBatchReader batchRecon = ReconciliationBatch.ExecuteReader("reconciliation_batch_id = " + _batchID);
        //                            while (batchRecon.Read())
        //                            {
        //                                Atm.AtmReader _batchAtm = Atm.ExecuteReader("ATM_id = " + batchRecon.CurrentReconciliationBatch.AtmId);
        //                                while (_batchAtm.Read())
        //                                {
        //                                    batchAtmTitle = _batchAtm.CurrentAtm.Title.Trim();
        //                                }
        //                            }

        //                            foreach (string part in parts)
        //                            {
        //                                string[] subParts = part.Split(',');
        //                                ReconciliationHostData hostData = new ReconciliationHostData();
        //                                ReconciliationBnaHostData BNAhostData = new ReconciliationBnaHostData();
        //                                //check if batch atm title matches the CSV data atm title
        //                                if (batchAtmTitle == subParts[0].Trim())
        //                                {
        //                                    if (withdrawalTran.ContainsKey(Convert.ToInt32(subParts[8])))
        //                                    {
        //                                        hostData.ReconciliationBatchId = _batchID;
        //                                        hostData.AtmId = subParts[0];
        //                                        hostData.CardNumber = subParts[1];
        //                                        hostData.CustomerAccountNo = subParts[2];
        //                                        hostData.TransactionDate = subParts[3];
        //                                        hostData.TransactionTime = subParts[4];
        //                                        hostData.TransactionAmount = decimal.Parse(subParts[5]);
        //                                        hostData.TransactionCurrency = subParts[6];
        //                                        hostData.TransactionSequence = subParts[7];
        //                                        hostData.TransactionType = withdrawalTran[Convert.ToInt32(subParts[8])];
        //                                        hostData.TransactionResponse = subParts[9];
        //                                        hostData.TransactionSettlementDate = DateTime.ParseExact(subParts[10], "yyyyMMdd", null);
        //                                        hostData.CardNetwork = subParts[11];
        //                                        hostData.CardIssuer = subParts[12];
        //                                        hostData.CardAcquirer = Customer_FIT_BankName;// "GBK"; //always  GBK for ON US Withdrawal
        //                                        hostData.IsKnet = false;
        //                                        hostData.Save();
        //                                    }
        //                                    else if (depositTran.ContainsKey(Convert.ToInt32(subParts[8])))
        //                                    {
        //                                        BNAhostData.ReconciliationBatchId = _batchID;
        //                                        BNAhostData.AtmId = subParts[0];
        //                                        BNAhostData.CardNumber = subParts[1];
        //                                        BNAhostData.CustomerAccountNo = subParts[2];
        //                                        BNAhostData.TransactionDate = subParts[3];
        //                                        BNAhostData.TransactionTime = subParts[4];
        //                                        BNAhostData.TransactionAmount = decimal.Parse(subParts[5]);
        //                                        BNAhostData.TransactionCurrency = subParts[6];
        //                                        BNAhostData.TransactionSequence = subParts[7];
        //                                        BNAhostData.TransactionType = depositTran[Convert.ToInt32(subParts[8])];
        //                                        BNAhostData.TransactionResponse = subParts[9];
        //                                        BNAhostData.TransactionSettlementDate = DateTime.ParseExact(subParts[10], "yyyyMMdd", null);
        //                                        BNAhostData.CardNetwork = subParts[11];
        //                                        BNAhostData.CardIssuer = subParts[12];
        //                                        BNAhostData.Save();
        //                                    }
        //                                }
        //                            }
        //                        }

        //                        string data1 = Encoding.ASCII.GetString(File.ReadAllBytes(recon.HostFilePath));
        //                        string[] parts1 = data1.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
        //                        foreach (string part in parts1)
        //                        {
        //                            string[] subParts = part.Split(',');
        //                            ReconciliationHostData hostData = new ReconciliationHostData();
        //                            ReconciliationBnaHostData BNAhostData = new ReconciliationBnaHostData();
        //                            //check if batch atm title matches the CSV data atm title
        //                            if (Cardless_Tellerx_TerminalID == subParts[0].Trim())
        //                            {
        //                                if (withdrawalTran.ContainsKey(Convert.ToInt32(subParts[8])))
        //                                {
        //                                    hostData.ReconciliationBatchId = Cardless_Tellerx_BatchID;
        //                                    hostData.AtmId = subParts[0];
        //                                    hostData.CardNumber = subParts[1];
        //                                    hostData.CustomerAccountNo = subParts[2];
        //                                    hostData.TransactionDate = subParts[3];
        //                                    hostData.TransactionTime = subParts[4];
        //                                    hostData.TransactionAmount = decimal.Parse(subParts[5]);
        //                                    hostData.TransactionCurrency = subParts[6];
        //                                    hostData.TransactionSequence = subParts[7];
        //                                    hostData.TransactionType = withdrawalTran[Convert.ToInt32(subParts[8])];
        //                                    hostData.TransactionResponse = subParts[9];
        //                                    hostData.TransactionSettlementDate = DateTime.ParseExact(subParts[10], "yyyyMMdd", null);
        //                                    hostData.CardNetwork = subParts[11];
        //                                    hostData.CardIssuer = subParts[12];
        //                                    hostData.IsKnet = false;
        //                                    hostData.Save();
        //                                }
        //                                else if (depositTran.ContainsKey(Convert.ToInt32(subParts[8])))
        //                                {
        //                                    BNAhostData.ReconciliationBatchId = Cardless_Tellerx_BatchID;
        //                                    BNAhostData.AtmId = subParts[0];
        //                                    BNAhostData.CardNumber = subParts[1];
        //                                    BNAhostData.CustomerAccountNo = subParts[2];
        //                                    BNAhostData.TransactionDate = subParts[3];
        //                                    BNAhostData.TransactionTime = subParts[4];
        //                                    BNAhostData.TransactionAmount = decimal.Parse(subParts[5]);
        //                                    BNAhostData.TransactionCurrency = subParts[6];
        //                                    BNAhostData.TransactionSequence = subParts[7];
        //                                    BNAhostData.TransactionType = depositTran[Convert.ToInt32(subParts[8])];
        //                                    BNAhostData.TransactionResponse = subParts[9];
        //                                    BNAhostData.TransactionSettlementDate = DateTime.ParseExact(subParts[10], "yyyyMMdd", null);
        //                                    BNAhostData.CardNetwork = subParts[11];
        //                                    BNAhostData.CardIssuer = subParts[12];
        //                                    BNAhostData.Save();
        //                                }
        //                            }
        //                        }


        //                        File.Delete(recon.BackupFolderPath + "\\hostfile" + trxnDate.AddDays(1).ToString("ddMMyyyy") + ".csv");
        //                        File.Move(recon.HostFilePath, recon.BackupFolderPath + "\\hostfile" + trxnDate.AddDays(1).ToString("ddMMyyyy") + ".csv");
        //                    }

        //                    if (File.Exists(knetTapeFilePath))
        //                    {
        //                        LogableTask.LogMonoActivityTask("STFileExists", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "ST File Exists");
        //                        foreach (int _batchID in batchID)
        //                        {

        //                            //to get the atm title of the iterated batchid
        //                            string batchAtmTitle = "";
        //                            ReconciliationBatch.ReconciliationBatchReader batchRecon = ReconciliationBatch.ExecuteReader("reconciliation_batch_id = " + _batchID);
        //                            while (batchRecon.Read())
        //                            {
        //                                Atm.AtmReader _batchAtm = Atm.ExecuteReader("ATM_id = " + batchRecon.CurrentReconciliationBatch.AtmId);
        //                                while (_batchAtm.Read())
        //                                {
        //                                    batchAtmTitle = _batchAtm.CurrentAtm.Title.Trim();
        //                                }
        //                            }
        //                            string data = Encoding.ASCII.GetString(File.ReadAllBytes(knetTapeFilePath));
        //                            string[] parts = data.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
        //                            foreach (string part in parts)
        //                            {
        //                                if (part.Trim() != "")
        //                                {
        //                                    string serviceType = part.Substring(0, 4).Trim();
        //                                    string cardAcquirer = part.Substring(98, 4).Trim();
        //                                    string cardIssuer = part.Substring(144, 4).Trim();
        //                                    string transCode = part.Substring(40, 2).Trim();//transCode == "10" WITHDRAWAL
        //                                    string terminalID = part.Substring(102, 16).Trim();

        //                                    //Customer_FIT_BankName GBK
        //                                    if (serviceType == "ATM" && cardAcquirer == Customer_FIT_BankName && cardIssuer != Customer_FIT_BankName && transCode == "10"
        //                                        && terminalID.Trim() == batchAtmTitle.Trim())
        //                                    {
        //                                        string cardNo = part.Substring(148, 19).Trim().Replace("X", "*");
        //                                        string transactionDate = "20" + part.Substring(8, 6).Trim();
        //                                        string transactionTime = part.Substring(14, 2).Trim() + ":" + part.Substring(16, 2).Trim();
        //                                        string currency = "KWD";
        //                                        string amount = part.Substring(43, 11).Trim();
        //                                        string seqNo = part.Substring(28, 12).Trim();
        //                                        string transactionResponse = part.Substring(315, 3).Trim();
        //                                        string settlementDate = part.Substring(22, 6).Trim();
        //                                        string authType = part.Substring(118, 4).Trim();
        //                                        // int _batchID = 101;

        //                                        ReconciliationHostData hostData = new ReconciliationHostData();
        //                                        hostData.ReconciliationBatchId = _batchID;
        //                                        hostData.AtmId = terminalID;

        //                                        string starString = "";
        //                                        int starstringLen = cardNo.Trim().Length - 10;
        //                                        for (int sl = 0; sl < starstringLen; sl++)
        //                                        { starString = starString + "*"; }

        //                                        hostData.CardNumber = cardNo.Substring(0, 6) + starString;
        //                                        hostData.CardNumber = hostData.CardNumber + cardNo.Substring(hostData.CardNumber.Length, 4);

        //                                        // hostData.CardNumber = cardNo.Substring(0, 6) + "******" + cardNo.Substring(12, cardNo.Length - 12);//4 cardno can be 15/16 digit
        //                                        // hostData.CustomerAccountNo = subPar;
        //                                        hostData.TransactionDate = transactionDate;
        //                                        hostData.TransactionTime = transactionTime;
        //                                        hostData.TransactionAmount = decimal.Parse(amount) / 1000;
        //                                        hostData.TransactionCurrency = currency;
        //                                        hostData.TransactionSequence = seqNo;
        //                                        if (authType == "0210")
        //                                            hostData.TransactionType = withdrawalTran[10];  //cash withdrawal-KNET
        //                                        else if (authType == "0420") //reversal
        //                                            hostData.TransactionType = withdrawalTran[420];  //withdrawal-KNET-correction
        //                                        hostData.TransactionResponse = transactionResponse;
        //                                        hostData.TransactionSettlementDate = DateTime.ParseExact(settlementDate, "yyMMdd", null);
        //                                        // hostData.CardNetwork = subParts[11];
        //                                        hostData.CardIssuer = cardIssuer;
        //                                        hostData.CardAcquirer = cardAcquirer;
        //                                        hostData.IsKnet = true;
        //                                        hostData.Save();

        //                                        LogableTask.LogMonoActivityTask("", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "KNET TAPE :" + serviceType +
        //                                          ", " + terminalID + ", " + cardAcquirer + ", " + cardIssuer);
        //                                    }
        //                                }
        //                            }

        //                        }
        //                        File.Delete(recon.BackupFolderPath + "\\" + knetTapeFileName + "file" + trxnDate.AddDays(1).ToString("ddMMyyyy") + ".txt");
        //                        File.Move(knetTapeFilePath, recon.BackupFolderPath + "\\" + knetTapeFileName + "file" + trxnDate.AddDays(1).ToString("ddMMyyyy") + ".txt");
        //                        LogableTask.LogMonoActivityTask("Sleep", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "Thread going to sleep");

        //                    }

        //                    LogableTask.LogMonoActivityTask("BatchId", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "Rerun-BatchId = " + string.Join(", ", batchID));
        //                    //Insert Batch Data--------END *************************


        //                }


        //            }
        //        }
        //        else
        //        {
        //            LogableTask.DefaultTraceLevel = TraceLevel.Info;
        //            LogableTask.LogMonoActivityTask("Switch_HostFileAbsence", MethodBase.GetCurrentMethod(), TraceLevel.Error, "Switch or Host or Knet ST File is missing");

        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        LogableTask.LogMonoActivityTask("Exception", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Error, ex.ToString());
        //        EventLog.WriteEntry("CCMSSchedular", "Error in DoWork(). detail: " + ex.Message + ex.StackTrace, EventLogEntryType.Error);
        //    }
        //}
        //*************************************************

        public void OnDebug()
        {
            OnStart(null);
        }
        protected override void OnStart(string[] args)
        {
            timerScheduleThreadForExecution = new Timer(ScheduleThreadForExecution, null, new TimeSpan(0, 0, 10), new TimeSpan(0, 0, 0, 0, -1));
            EventLog.WriteEntry("ReconciliationSchedular", "Thread schedular sent startup request", EventLogEntryType.Information);

        }

        protected override void OnStop()
        {
            try
            {
                //File.AppendAllText(path, DateTime.Now + " Hello World from OnStop\n");
                LogableTask.LogMonoActivityTask("Stopping", MethodBase.GetCurrentMethod(), TraceLevel.Warning, "Service Stopped");
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry("ReconciliationSchedularNotifier", "Error in OnStop(). detail: " + ex.Message + ex.StackTrace, EventLogEntryType.Error);
            }
        }


        // void ScheduleBatch(ReconciliationBatch reconciliationBatch)
        //{
        //    reconciliationBatch.Status = "Scheduled";
        //    reconciliationBatch.Save();
        //}

        void uploadRaw(DataTable dtCAF)
        {
            string knetSTFileName = "ST";//STTokenized
            string knetSTFilePath = recon.HostFilePath.Replace("host.csv", knetSTFileName + DateTime.Now.AddDays(-1).ToString("yyMMdd") + ".txt");
            SqlCommand cmd = null;
            SqlTransaction sqlTrxn = null;


            if (File.Exists(knetSTFilePath))
            {
                //Switch
                try
                {
                    //LogableTask.LogMonoActivityTask("DRV_STFileExists", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "1. STTokenized BEGIN" + _isFilesinUSE);

                    //if (!_isFilesinUSE)
                    //{

                    //KNET ST
                    //try
                    //{
                    //if (File.Exists(knetSTFilePath))
                    //{
                    //LogableTask.LogMonoActivityTask("IntervalSleep", MethodBase.GetCurrentMethod(), TraceLevel.Info, "Sleeping for 60 seconds.");
                    //Thread.Sleep(60000);

                    //LogableTask.LogMonoActivityTask("STFileExists", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "ST Tokenized File Exists__");

                    //string cafFileName = "CAF";
                    //string cafFilePath = recon.HostFilePath.Replace("host.csv", cafFileName + ".txt");
                    //DataTable dtCAF = new DataTable();
                    //dtCAF.Columns.Add("CardNo", typeof(String));
                    //dtCAF.Columns.Add("AccountNo", typeof(String));

                    //if (File.Exists(cafFilePath))
                    //{
                    //    LogableTask.LogMonoActivityTask("CAFFileExists", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "CAF File Exists__");

                    //    string dataCAF = Encoding.ASCII.GetString(File.ReadAllBytes(cafFilePath));
                    //    string[] partsCAF = dataCAF.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
                    //    foreach (string part in partsCAF)
                    //    {
                    //        if (part.Trim() != "")
                    //        {
                    //            if (part.Substring(9, 2).Trim() == "DT") // check if row is Details record
                    //            {
                    //                DataRow dr = dtCAF.NewRow();

                    //                dr[0] = part.Substring(15, 19).Trim(); //card number
                    //                dr[1] = part.Substring(78, 19).Trim(); //accountNumber

                    //                dtCAF.Rows.Add(dr);
                    //            }
                    //        }
                    //    }
                    //}

                    List<ReconciliationHostDataRAW> listReconciliationHostDataRAW = new List<ReconciliationHostDataRAW>();
                    StringBuilder builder;
                    string data = File.ReadAllText(knetSTFilePath);
                    string[] parts = data.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (string part in parts)
                    {
                        if (part.Trim() != "")
                        {
                            string serviceType = part.Substring(0, 4).Trim();
                            string cardAcquirer = part.Substring(98, 4).Trim();
                            string cardIssuer = part.Substring(144, 4).Trim();
                            string transCode = part.Substring(40, 2).Trim();//transCode == "10" WITHDRAWAL
                            string terminalID = part.Substring(102, 16).Trim();

                            string cardNo = part.Substring(148, 19).Trim().Replace("X", "*");
                            string transactionDate = "20" + part.Substring(8, 6).Trim();
                            string transactionTime = part.Substring(14, 2).Trim() + ":" + part.Substring(16, 2).Trim();
                            string currency = "KWD";
                            string amount = part.Substring(43, 11).Trim();
                            string seqNo = part.Substring(28, 12).Trim();
                            string transactionResponse = part.Substring(167, 2).Trim();
                            string settlementDate = part.Substring(22, 6).Trim();
                            string authType = part.Substring(118, 4).Trim();
                            string cardType = part.Substring(169, 4).Trim();
                            string CustomerAccountNo = "";
                            //get account number from CAF file
                            if (dtCAF.Rows.Count > 0)
                            {
                                DataRow[] drArray = dtCAF.Select(" CardNo='" + cardNo.Trim() + "'").ToArray(); ;
                                if (drArray.Length > 0 && drArray[0]["CardNo"].ToString().Trim() != "")
                                {
                                    CustomerAccountNo = drArray[0]["AccountNo"].ToString().Trim();
                                }

                            }

                            string starString = "";
                            int starstringLen = cardNo.Trim().Length - 10;
                            for (int sl = 0; sl < starstringLen; sl++)
                            { starString = starString + "*"; }

                            string CardNumber = cardNo.Substring(0, 6) + starString;
                            CardNumber = CardNumber + cardNo.Substring(CardNumber.Length, 4);

                            // hostData.CardNumber = cardNo.Substring(0, 6) + "******" + cardNo.Substring(12, cardNo.Length - 12);//4 cardno can be 15/16 digit
                            // hostData.CustomerAccountNo = subPar;
                            decimal TransactionAmount = decimal.Parse(amount) / 1000;
                            string TransactionType = "";
                            if (authType == "0210")
                                TransactionType = serviceType;// withdrawalTran["10"];  //cash withdrawal-KNET
                            else if (authType == "0420") //reversal
                                TransactionType = serviceType + "-R";// withdrawalTran["420"];  //withdrawal-KNET-correction
                            string TransactionResponse = Regex.Replace(transactionResponse, "[^0-9]", "10");//replace character with int code
                            DateTime TransactionSettlementDate = DateTime.ParseExact(settlementDate, "yyMMdd", null);

                            bool IsKnet = true;
                            string KNETTranCode = transCode;

                            //builder = new StringBuilder();

                            //builder.Append("insert into reconciliation_host_data_RAW values (" + RemoteOnUs_BatchID + ",'" + terminalID + "', '" + CardNumber + "','" + CustomerAccountNo + "'" +
                            //    ", '" + transactionDate + "', '" + transactionTime + "', " + TransactionAmount + ",'" + currency + "', '" + seqNo + "', '" + TransactionType +
                            //    "', '" + TransactionResponse + "', '" + TransactionSettlementDate + "', '" + cardType + "','" + cardIssuer + "', '" + cardAcquirer + "','" + IsKnet +
                            //    "','" + transCode + "' ) ;");

                            ReconciliationHostDataRAW reconciliationHostDataRAW = new ReconciliationHostDataRAW();
                            reconciliationHostDataRAW.ReconciliationBatchId = RemoteOnUs_BatchID;
                            reconciliationHostDataRAW.AtmId = terminalID;
                            reconciliationHostDataRAW.CardNumber = CardNumber;
                            reconciliationHostDataRAW.CustomerAccountNo = CustomerAccountNo;
                            reconciliationHostDataRAW.TransactionDate = transactionDate;
                            reconciliationHostDataRAW.TransactionTime = transactionTime;
                            reconciliationHostDataRAW.TransactionAmount = TransactionAmount;
                            reconciliationHostDataRAW.TransactionCurrency = currency;
                            reconciliationHostDataRAW.TransactionSequence = seqNo;
                            reconciliationHostDataRAW.TransactionType = TransactionType;
                            reconciliationHostDataRAW.TransactionResponse = TransactionResponse;
                            reconciliationHostDataRAW.TransactionSettlementDate = TransactionSettlementDate;
                            reconciliationHostDataRAW.CardNetwork = cardType;
                            reconciliationHostDataRAW.CardIssuer = cardIssuer;
                            reconciliationHostDataRAW.CardAcquirer = cardAcquirer;
                            reconciliationHostDataRAW.IsKnet = IsKnet;
                            reconciliationHostDataRAW.KnetTranCode = transCode;

                            listReconciliationHostDataRAW.Add(reconciliationHostDataRAW);

                            //ConnectionFactory.ExecuteQuery(builder.ToString(), sqlTrxn);
                        }
                    }
                    cmd = ConnectionFactory.GetNewCommand(true);
                    sqlTrxn = cmd.Connection.BeginTransaction();
                    cmd.Transaction = sqlTrxn;

                    ReconciliationHostDataRAW.BulkSave(listReconciliationHostDataRAW, sqlTrxn);
                    sqlTrxn.Commit();
                    //}

                    LogableTask.LogMonoActivityTask("STFileMoved", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "ST Tokenized File Completed__");
                    // File.Delete(recon.BackupFolderPath + "\\" + knetTapeFileName + "file" + trxnDate.AddDays(1).ToString("ddMMyyyy") + ".txt");
                    //File.Move(knetSTFilePath, recon.BackupFolderPath + "\\" + knetSTFileName + "file" + DateTime.Now.ToString("ddMMyyyy_HHmmss") + ".txt");
                    //LogableTask.LogMonoActivityTask("STFileMoved", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Info, "ST Tokenized File Moved__");

                    //}
                    //catch (Exception ex)
                    //{
                    //    LogableTask.LogMonoActivityTask("Exception", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Error, ex.ToString());
                    //    EventLog.WriteEntry("CCMSSchedular", "Error in DoWork_Ver2(). KNET ST file schedule detail: " + ex.Message + ex.StackTrace, EventLogEntryType.Error);
                    //}


                    // }
                    //ik
                }
                catch (Exception ex)
                {
                    if (sqlTrxn != null)
                        sqlTrxn.Rollback();

                    LogableTask.LogMonoActivityTask("Exception", System.Reflection.MethodBase.GetCurrentMethod(), TraceLevel.Error, ex);

                }
                finally
                {
                    if (cmd != null)
                        if (cmd.Connection != null)
                            cmd.Connection.Close();
                }
            }
        }

        public static DataTable GetDataTableFromExcel(string filePath, string sheetName)
        {
            DataTable dt = new DataTable();

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(sheetName);
                bool firstRow = true;

                foreach (var row in worksheet.RowsUsed())
                {
                    if (firstRow)
                    {
                        // Add columns to DataTable
                        foreach (var cell in row.Cells())
                        {
                            dt.Columns.Add(cell.Value.ToString());
                        }
                        firstRow = false;
                    }
                    else
                    {
                        // Add rows to DataTable
                        dt.Rows.Add(row.Cells().Select(cell => cell.Value.ToString()).ToArray());
                    }
                }
            }

            return dt;

        }
        public static DataTable ExcelToDataTable(string filePath, string sheetName = null, int rowsToSkipTillHeader = 0)
        {
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(stream))
            {
                var conf = new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = true,

                        ReadHeaderRow = rowReader =>
                        {
                            for (int i = 0; i < rowsToSkipTillHeader; i++)
                            {
                                if (!rowReader.Read())
                                    break;
                            }
                        }
                    }
                };

                var dataSet = reader.AsDataSet(conf);

                if (dataSet.Tables.Count == 0)
                    throw new ArgumentException("The Excel file contains no worksheets.");

                // If sheetName is null or empty, use the first sheet
                DataTable source;
                if (string.IsNullOrWhiteSpace(sheetName))
                {
                    source = dataSet.Tables[0];
                }
                else
                {
                    if (!dataSet.Tables.Contains(sheetName))
                        throw new ArgumentException(
                            $"Sheet '{sheetName}' does not exist in the Excel file.");

                    source = dataSet.Tables[sheetName];
                }

                var result = new DataTable();
                foreach (DataColumn col in source.Columns)
                    result.Columns.Add(col.ColumnName, typeof(string));

                foreach (DataRow row in source.Rows)
                {
                    var newRow = result.NewRow();
                    for (int i = 0; i < source.Columns.Count; i++)
                        newRow[i] = row[i]?.ToString() ?? string.Empty;

                    result.Rows.Add(newRow);
                }

                return CleanDataTable(result);
            }
        }

        private static DataTable CleanDataTable(DataTable source)
        {
            var cleanTable = new DataTable();

            // 1) Keep only columns that have a header name.
            var validColumns = new List<DataColumn>();
            foreach (DataColumn col in source.Columns)
            {
                if (!string.IsNullOrWhiteSpace(col.ColumnName))
                {
                    // Treat only default names like "Column0", "Column1", "Column2" as empty
                    bool isDefaultName = Regex.IsMatch(col.ColumnName, @"^Column\d+$", RegexOptions.IgnoreCase);

                    if (!isDefaultName)
                    {
                        cleanTable.Columns.Add(col.ColumnName, typeof(string));
                        validColumns.Add(col);
                    }
                }
            }


            // 2) Copy only non-empty rows
            foreach (DataRow row in source.Rows)
            {
                bool rowHasValue = false;
                var newRow = cleanTable.NewRow();

                for (int i = 0; i < validColumns.Count; i++)
                {
                    var value = row[validColumns[i]]?.ToString() ?? string.Empty;
                    newRow[i] = value.Trim();

                    if (!string.IsNullOrWhiteSpace(value))
                        rowHasValue = true;
                }

                if (rowHasValue)
                    cleanTable.Rows.Add(newRow);
            }

            return cleanTable;
        }

        void InsertIntoReconciliationBatchDictionary(DateTime TransactionStartDate, int AtmId, int ReconciliationBatchId)
        {
            string key = TransactionStartDate.ToString("yyyy/MM/dd") + "_" + AtmId.ToString();
            ReconciliationBatchModel reconciliationBatchModel = new ReconciliationBatchModel();
            reconciliationBatchModel.TransactionStartDate = TransactionStartDate;
            reconciliationBatchModel.AtmId = AtmId;
            reconciliationBatchModel.ReconciliationBatchId = ReconciliationBatchId;
            // Add to the dictionary
            ReconBatchDict[key] = reconciliationBatchModel;
        }
        void ImportHostFile()
        {

        }
    }

    public static class CurrencyCodes
    {
        public static readonly Dictionary<int, string> NumericToAlphaCode = new Dictionary<int, string>
        {
            { 4, "AFN" },
            { 8, "ALL" },
            { 12, "DZD" },
            { 32, "ARS" },
            { 36, "AUD" },
            { 44, "BSD" },
            { 48, "BHD" },
            { 50, "BDT" },
            { 51, "AMD" },
            { 52, "BBD" },
            { 60, "BMD" },
            { 64, "BTN" },
            { 68, "BOB" },
            { 72, "BWP" },
            { 84, "BZD" },
            { 90, "SBD" },
            { 96, "BND" },
            { 104, "MMK" },
            { 108, "BIF" },
            { 116, "KHR" },
            { 124, "CAD" },
            { 132, "CVE" },
            { 136, "KYD" },
            { 144, "LKR" },
            { 152, "CLP" },
            { 156, "CNY" },
            { 170, "COP" },
            { 174, "KMF" },
            { 188, "CRC" },
            { 191, "HRK" },
            { 192, "CUP" },
            { 203, "CZK" },
            { 208, "DKK" },
            { 214, "DOP" },
            { 222, "SVC" },
            { 230, "ETB" },
            { 232, "ERN" },
            { 238, "FKP" },
            { 242, "FJD" },
            { 262, "DJF" },
            { 270, "GMD" },
            { 292, "GIP" },
            { 320, "GTQ" },
            { 324, "GNF" },
            { 328, "GYD" },
            { 332, "HTG" },
            { 340, "HNL" },
            { 344, "HKD" },
            { 348, "HUF" },
            { 352, "ISK" },
            { 356, "INR" },
            { 360, "IDR" },
            { 364, "IRR" },
            { 368, "IQD" },
            { 376, "ILS" },
            { 388, "JMD" },
            { 392, "JPY" },
            { 398, "KZT" },
            { 400, "JOD" },
            { 404, "KES" },
            { 410, "KRW" },
            { 414, "KWD" },
            { 417, "KGS" },
            { 418, "LAK" },
            { 422, "LBP" },
            { 426, "LSL" },
            { 430, "LRD" },
            { 434, "LYD" },
            { 446, "MOP" },
            { 454, "MWK" },
            { 458, "MYR" },
            { 462, "MVR" },
            { 480, "MUR" },
            { 484, "MXN" },
            { 496, "MNT" },
            { 498, "MDL" },
            { 504, "MAD" },
            { 512, "OMR" },
            { 516, "NAD" },
            { 524, "NPR" },
            { 532, "ANG" },
            { 533, "AWG" },
            { 548, "VUV" },
            { 554, "NZD" },
            { 566, "NGN" },
            { 578, "NOK" },
            { 586, "PKR" },
            { 590, "PAB" },
            { 598, "PGK" },
            { 600, "PYG" },
            { 604, "PEN" },
            { 608, "PHP" },
            { 634, "QAR" },
            { 643, "RUB" },
            { 646, "RWF" },
            { 654, "SHP" },
            { 682, "SAR" },
            { 690, "SCR" },
            { 694, "SLL" },
            { 702, "SGD" },
            { 704, "VND" },
            { 706, "SOS" },
            { 710, "ZAR" },
            { 728, "SSP" },
            { 752, "SEK" },
            { 756, "CHF" },
            { 760, "SYP" },
            { 764, "THB" },
            { 776, "TOP" },
            { 780, "TTD" },
            { 784, "AED" },
            { 788, "TND" },
            { 800, "UGX" },
            { 807, "MKD" },
            { 818, "EGP" },
            { 826, "GBP" },
            { 834, "TZS" },
            { 840, "USD" },
            { 858, "UYU" },
            { 860, "UZS" },
            { 882, "WST" },
            { 886, "YER" },
            { 901, "TWD" },
            { 934, "TMT" },
            { 936, "GHS" },
            { 941, "RSD" },
            { 943, "MZN" },
            { 944, "AZN" },
            { 946, "RON" },
            { 949, "TRY" },
            { 960, "XDR" },
            { 967, "ZMW" },
            { 971, "AFN" },
            { 972, "TJS" },
            { 976, "BGN" },
            { 977, "BAM" },
            { 978, "EUR" },
            { 985, "PLN" },
            { 986, "BRL" },
            { 990, "CLF" },
            { 997, "USN" }
        };
    }
}
