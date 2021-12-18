using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace EJTool
{
    public class ProcessEJ
    {
        const string csTxStart = "-> TRANSACTION START";
        const string csTxEnd = "<- TRANSACTION END";
        const string csCardNum = "TRACK 2 DATA:";
        const string csCashRq = "CASH REQUEST:";
        const string csOpCode = "TRANSACTION REQUEST";
        const string csAmount = "Amount:";
        const string csCashTaken = "CASH TAKEN";
        const string csCashRetracted = "CASH RETRACTED";
        const string csRRN = "RRN:[";
        const string csTranType = "*******";
        const string csCashWithdrawal = "CASH WITHDRAWAL";
        const string csBalance = "BALANCE INQ";
        const string csMini = "MINI STATEMENTS";

        public void ProcessEJFiles(string[] EJFiles)
        {
            //string line = "";
            Transaction tran = new Transaction();
            List<Transaction> transactions = new List<Transaction>();
            StreamReader stReader = null;
            foreach (string filePath in EJFiles)
            {
                int index = 0;
                try
                {
                    stReader = new StreamReader(filePath);
                    index = 0;
                    string strDate = Path.GetFileNameWithoutExtension(filePath);

                    foreach (string line in File.ReadLines(filePath, Encoding.UTF8))
                    {
                        //    // process the line
                        //}
                        //while (!stReader.EndOfStream)
                        //{
                        //line = stReader.ReadLine();
                        index++;
                        if (String.IsNullOrEmpty(line))
                        {
                            continue;
                        }
                        else if (line.Contains(csTxStart))
                        {
                            tran = new Transaction();
                            tran.ResetAllData();
                            tran.strStartTime = line.Substring(0, 8);
                            continue;
                        }
                        else if (line.Contains(csTxEnd))
                        {
                            tran.strEndTime = line.Substring(0, 8);
                            if (tran.eTranType == TransactionType.DISPENSE)
                            {
                                if (tran.eCashTranType == CashTransactionType.SUCCESS)
                                {
                                    tran.bIsTranSuccess = true;
                                }
                            }
                            tran.strDate = strDate;
                            transactions.Add(tran);
                            continue;
                        }
                        else if (line.Contains(csCardNum))
                        {
                            tran.strCardNum = line.Substring(23);
                            continue;
                        }
                        else if (line.Contains(csCashRq))
                        {
                            tran.strDipsRequest = line.Substring(23);
                            if (tran.strDipsRequest.Length < 8)
                            {
                                tran.strDipsRequest = tran.strDipsRequest.PadRight(8, '0');
                            }
                            Int32.TryParse(tran.strDipsRequest.Substring(0, 2), out tran.arCashDip[0]);
                            Int32.TryParse(tran.strDipsRequest.Substring(2, 2), out tran.arCashDip[1]);
                            Int32.TryParse(tran.strDipsRequest.Substring(4, 2), out tran.arCashDip[2]);
                            Int32.TryParse(tran.strDipsRequest.Substring(6, 2), out tran.arCashDip[3]);
                            continue;
                        }
                        else if (line.Contains(csOpCode))
                        {
                            tran.strOpCode = line.Substring(29);
                            continue;
                        }
                        else if (line.Contains(csAmount))
                        {
                            tran.strAmount = line.Substring(8);
                            continue;
                        }
                        else if (line.Contains(csCashRetracted))
                        {
                            tran.eCashTranType = CashTransactionType.RETRACTED;
                            continue;
                        }
                        else if (line.Contains(csCashTaken))
                        {
                            tran.eCashTranType = CashTransactionType.SUCCESS;
                            continue;
                        }
                        else if (line.Contains(csRRN))
                        {
                            tran.strRRN = line.Substring(5, 12);
                        }
                        else if (line.Contains(csTranType))
                        {
                            if (line.Contains(csCashWithdrawal))
                            {
                                tran.eTranType = TransactionType.DISPENSE;
                            }
                            else if (line.Contains(csMini))
                            {
                                tran.eTranType = TransactionType.MINI;
                            }
                            else if (line.Contains(csBalance))
                            {
                                tran.eTranType = TransactionType.BALANCE;
                            }
                            continue;
                        }
                    }
                    Console.WriteLine("Read {0} files done!", filePath);
                }
                catch (Exception ex)
                {
                    throw new Exception("Current File: " + filePath, ex);
                }
            }
            Console.WriteLine("Read all files done!");
            Microsoft.Office.Interop.Excel.Application oXL = null;
            Microsoft.Office.Interop.Excel._Workbook oWB = null;
            Microsoft.Office.Interop.Excel._Worksheet oSheet = null;
            object misvalue = System.Reflection.Missing.Value;
            try
            {
                oXL = new Microsoft.Office.Interop.Excel.Application();
                string path = EJFiles[0];
                string fpath = Path.GetDirectoryName(path);
                //oWB = oXL.Workbooks.Open(Path.GetDirectoryName(path) + "\\Report.xlsx");

                oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
                oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

                //oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;
                oSheet.Cells[1, 1] = "STT"; // Row Col
                oSheet.Cells[1, 2] = "DATE";
                oSheet.Columns[2].NumberFormat = "@";
                oSheet.Cells[1, 3] = "START";
                oSheet.Cells[1, 4] = "END";
                oSheet.Cells[1, 5] = "CARD NO";
                oSheet.Columns[5].NumberFormat = "@";
                oSheet.Cells[1, 6] = "RRN";
                oSheet.Columns[6].NumberFormat = "@";
                oSheet.Cells[1, 7] = "OPKEY";
                oSheet.Cells[1, 8] = "AMOUNT";
                oSheet.Cells[1, 9] = "DISP STRING";
                oSheet.Columns[9].NumberFormat = "@";
                oSheet.Cells[1, 10] = "C1";
                oSheet.Cells[1, 11] = "C2";
                oSheet.Cells[1, 12] = "C3";
                oSheet.Cells[1, 13] = "C4";
                oSheet.Cells[1, 14] = "TRAN TYPE";
                oSheet.Cells[1, 15] = "RESULT";

                int index = 2;
                foreach (Transaction currentTran in transactions)
                {
                    oSheet.Cells[index, 1] = index;
                    oSheet.Cells[index, 2] = currentTran.strDate;
                    oSheet.Cells[index, 3] = currentTran.strStartTime;
                    oSheet.Cells[index, 4] = currentTran.strEndTime;
                    oSheet.Cells[index, 5] = currentTran.strCardNum;
                    oSheet.Cells[index, 6] = currentTran.strRRN;
                    oSheet.Cells[index, 7] = currentTran.strOpCode;
                    oSheet.Cells[index, 8] = currentTran.strAmount;
                    oSheet.Cells[index, 9] = currentTran.strDipsRequest;
                    oSheet.Cells[index, 10] = currentTran.arCashDip[0];
                    oSheet.Cells[index, 11] = currentTran.arCashDip[1];
                    oSheet.Cells[index, 12] = currentTran.arCashDip[2];
                    oSheet.Cells[index, 13] = currentTran.arCashDip[3];
                    oSheet.Cells[index, 14] = currentTran.eTranType.ToString();
                    //oSheet.Cells[index, 1] = "STT"; // Row Col
                    //oSheet.Cells[index, 2] = "DATE";
                    //oSheet.Cells[index, 3] = "START";
                    //oSheet.Cells[index, 4] = "END";
                    //oSheet.Cells[index, 5] = "CARD NO";
                    //oSheet.Cells[index, 6] = "RRN";
                    //oSheet.Cells[index, 7] = "OPKEY";
                    //oSheet.Cells[index, 8] = "AMOUNT";
                    //oSheet.Cells[index, 9] = "DISP STRING";
                    //oSheet.Cells[index, 10] = "C1";
                    //oSheet.Cells[index, 11] = "C2";
                    //oSheet.Cells[index, 12] = "C3";
                    //oSheet.Cells[index, 13] = "C4";
                    //oSheet.Cells[index, 14] = "TRAN TYPE";
                    //oSheet.Cells[index, 15] = "RESULT";
                    //oSheet.Cells[index, 15] = currentTran.bIsTranSuccess;
                    if (currentTran.bIsTranSuccess)
                    {
                        oSheet.Cells[index, 15] = "SUCCESS";
                    }
                    else
                    {
                        oSheet.Cells[index, 15] = "UNKNOWN";
                    }
                    index++;
                }

                oWB.SaveAs(fpath + "\\Report.xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                // oWB.Save();
            }
            catch (Exception ex)
            {
            }
            finally
            {
                if (oWB != null)
                    oWB.Close();
            }
        }
    }
}
