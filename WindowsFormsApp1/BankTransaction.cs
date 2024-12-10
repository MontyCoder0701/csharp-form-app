namespace WindowsFormsApp1
{
    internal class BankTransaction
    {
        public int TranSeq;
        public int AcctSeq;
        public int? Balance100;
        public int? Bcode;
        public string Cid;
        public string Date;
        public string Description;
        public int? Order;
        public int? Pay100;
        public int? Save100;
        public string Time;
        public string Trankind;
        public int TransferSeq;
        public string Where;

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
