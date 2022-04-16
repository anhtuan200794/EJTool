using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EJTool
{
    public enum TransactionResult
    {
        FAILED = -1,
        SUCCESS = 0,
        RETRACTED,
        REJECTED,
    }
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
        public string strDispRequest { get; set; }
        public TransactionResult eTranResult { get; set; }
        public int[] arCashDisp { get; set;}
        public int[] arNumOfNote { get; set; }
        public string strRRN { get; set; }
        public string strRC { get; set; }
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
            strDispRequest = "";
            eTranResult = TransactionResult.FAILED;
            arCashDisp = new int[] { 0,0,0,0};
            arNumOfNote = new int[] { 0, 0, 0, 0 };
            strRRN = "";
            strRC = "";
        }
        public Transaction()
        {
            strDate = "";
            strStartTime = "";
            strEndTime = "";
            strOpCode = "";
            strAmount = "";
            eTranType = TransactionType.UNKNOWN;
            bIsTranSuccess = false;
            strCardNum = "";
            strDispRequest = "";
            eTranResult = TransactionResult.FAILED;
            arCashDisp = new int[] { 0, 0, 0, 0 };
            arNumOfNote = new int[] { 0, 0, 0, 0 };
            strRRN = "";
            strRC = "";
        }
    }
}
