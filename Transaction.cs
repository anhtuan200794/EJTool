using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EJTool
{
    public enum TransactionType
    {
        UNKNOWN = -1,
        DEPOSIT = 0,
        DISPENSE,
        BALANCE,
        MINI,
        PINCHANGE,
        FUNDTRANSFER
    }
    public enum CashTransactionType
    {
        UNKNOWN = -1,
        SUCCESS = 0,
        RETRACTED,
        REJECTED
    }
    class Transaction
    {
        public string strDate { get; set; }
        public string strStartTime { get; set; }
        public string strEndTime { get; set; }
        public string strOpCode { get; set; }
        public string strAmount { get; set; }
        public TransactionType eTranType {get;set;}
        public bool bIsTranSuccess { get; set; }
        public string strCardNum { get; set; }
        public string strDipsRequest { get; set; }
        public CashTransactionType eCashTranType { get; set; }
        public int[] arCashDip { get; set;}
        public string strRRN { get; set; }
        public void ResetAllData()
        {
            strDate = "";
            strStartTime = "";
            strEndTime = "";
            strOpCode = "";
            strAmount = "";
            eTranType = TransactionType.UNKNOWN;
            bIsTranSuccess = false;
            strCardNum = "";
            strDipsRequest = "";
            eCashTranType = CashTransactionType.UNKNOWN;
            arCashDip = new int[] { 0,0,0,0};
            strRRN = "";

        }
    }
}
