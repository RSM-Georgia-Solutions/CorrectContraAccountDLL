using System;

namespace CorrectContraAccountLogicDLL
{
    public class CorrectionJournalEntriesParams
    {
        public CorrectionJournalEntriesParams()
        {
            WaitingTimeInMinutes = 60;
        }
        public int MaxLine { get; set; }
        public bool MustSkip { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public int WaitingTimeInMinutes { get; set; }
    }
}
