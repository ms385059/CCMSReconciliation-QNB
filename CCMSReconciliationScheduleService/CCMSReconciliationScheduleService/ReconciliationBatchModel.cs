using System;

namespace ReconciliationScheduleService
{
    internal class ReconciliationBatchModel
    {
        internal int AtmId { get; set; }
        internal int ReconciliationBatchId { get; set; }
        internal DateTime TransactionStartDate { get; set; }
    }
}
