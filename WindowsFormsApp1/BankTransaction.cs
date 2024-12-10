namespace WindowsFormsApp1
{
    internal class BankTransaction
    {
        public int TranSeq { get; }
        public int AcctSeq { get; }
        public int? Balance100 { get; }
        public int? Bcode { get; }
        public string Cid { get; }
        public string Date { get; }
        public string Description { get; }
        public int? Order { get; }
        public int? Pay100 { get; }
        public int? Save100 { get; }
        public string Time { get; }
        public string Trankind { get; }
        public int TransferSeq { get; }
        public string Where { get; }

        public BankTransaction(
            int tranSeq,
            int acctSeq,
            int? balance100,
            int? bcode,
            string cid,
            string date,
            string description,
            int? order,
            int? pay100,
            int? save100,
            string time,
            string trankind,
            int transferSeq,
            string where
        )
        {
            TranSeq = tranSeq;
            AcctSeq = acctSeq;
            Balance100 = balance100;
            Bcode = bcode;
            Cid = cid;
            Date = date;
            Description = description;
            Order = order;
            Pay100 = pay100;
            Save100 = save100;
            Time = time;
            Trankind = trankind;
            TransferSeq = transferSeq;
            Where = where;
        }
    }
}
