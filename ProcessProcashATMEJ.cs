using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace EJTool
{
    public class ProcessProcashATMEJ
    {
        const string csTxStart = "-> TRANSACTION START";
        const string csTxEnd = "<- TRANSACTION END";
        const string csCardNum = "TRACK 2 DATA:";
        const string csPinEnter = "PIN ENTERED";
        const string csCashRq = "CASH REQUEST:";
        const string csOpCode = "TRANSACTION REQUEST";
        const string csAmount = "AMOUNT";
        const string csCashTaken = "CASH TAKEN";
        const string csCashRetracted = "CASH RETRACTED";
        const string csPGRRN = "RRN:[";
        const string csSHBRRN = "C.A.:";
        const string csVBARRN = "TRAN.ID";
        const string csCashWithdrawal = "CASH WITHDRAWAL";
        const string csBalance = "AVAILABLE BALANCE";
        const string csMini = "STATEMENT REQUEST";
        const string csVBAAmount = "WITHDRAWAL";
        const string csVBAResCode = "RESPONSE CODE";
        const string csVBABalance = "BALANCE INQUIRY";
        const string csVBACurrency = "00 VND";
        const string csVBAPinChange = "PIN CHANGE";
        const string csSHBTranType = "TRAN.:";
        int nTXStep = 0; // TX start:1 Pin Enter:2 Cash taken:3 TX end:0
        public void ProcessProcashATMEJFiles(string[] EJFiles)
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
                    string line;
                    while ((line = stReader.ReadLine()) != null)
                    {
                        index++;
                        Console.WriteLine(index + ":" + line);
                        //if (index == 1203)
                        //{
                        //    Console.WriteLine(index + ":" + line);
                        //}
                        if (String.IsNullOrEmpty(line))
                        {
                            continue;
                        }
                        #region TX Start
                        else if (line.Contains(csTxStart))
                        {
                            tran = new Transaction();
                            tran.strStartTime = line.Substring(0, 8);
                            nTXStep = 1;
                            continue;
                        }
                        #endregion

                        #region TX end
                        else if (line.Contains(csTxEnd))
                        {
                            tran.strEndTime = line.Substring(0, 8);
                            tran.strDate = strDate;
                            transactions.Add(tran);
                            nTXStep = 0;
                            continue;
                        }
                        else if (line.Contains(csPinEnter))
                        {
                            if (3 == nTXStep)
                            {
                                tran.strEndTime = line.Substring(0, 8);
                                tran.strDate = strDate;
                                transactions.Add(tran);
                                string ctrCardnumberTemp = tran.strCardNum;
                                tran = new Transaction();
                                tran.strStartTime = line.Substring(0, 8);
                                tran.strCardNum = ctrCardnumberTemp;
                                nTXStep = 1;

                                continue;
                            }
                            else
                            {
                                nTXStep = 2;
                            }
                            
                        }
                        #endregion

                        #region Get CardNum
                        else if (line.Contains(csCardNum))
                        {
                            tran.strCardNum = line.Substring(23);
                            continue;
                        }
                        #endregion

                        #region Get Disp String
                        else if (line.Contains(csCashRq))
                        {
                            tran.strDispRequest = line.Substring(23);
                            if (tran.strDispRequest.Length < 8)
                            {
                                tran.strDispRequest = tran.strDispRequest.PadRight(8, '0');
                            }
                            Int32.TryParse(tran.strDispRequest.Substring(0, 2), out tran.arCashDisp[0]);
                            Int32.TryParse(tran.strDispRequest.Substring(2, 2), out tran.arCashDisp[1]);
                            Int32.TryParse(tran.strDispRequest.Substring(4, 2), out tran.arCashDisp[2]);
                            Int32.TryParse(tran.strDispRequest.Substring(6, 2), out tran.arCashDisp[3]);
                            tran.eTranType = TransactionType.DISPENSE;
                            continue;
                        }
                        #endregion

                        #region Get OPP Code
                        else if (line.Contains(csOpCode))
                        {
                            tran.strOpCode = line.Substring(29);
                            continue;
                        }
                        #endregion

                        #region Get Amount
                        else if (line.Contains(csAmount))
                        {
                            tran.strAmount = line.Substring(8);
                            List<string> list = new List<string>();
                            list = line.Split(' ').ToList();
                            tran.strAmount = list.Find(p => p.Contains("000"));
                            continue;
                        }
                        #endregion

                        #region Get TX result
                        else if (line.Contains(csCashRetracted))
                        {
                            tran.eTranResult = TransactionResult.RETRACTED;
                            nTXStep = 3;
                            continue;
                        }
                        else if (line.Contains(csCashTaken))
                        {
                            tran.eTranResult = TransactionResult.CASHTAKEN;
                            nTXStep = 3;
                            continue;
                        }
                        else if (line.Contains(csVBAResCode)) //VBA
                        {
                            string temp = line.Substring(csVBAResCode.Length);
                            Int32.TryParse(temp, out int ResCode);
                            if (ResCode == 1)
                                tran.eTranResult = TransactionResult.SUCCESS;
                            else
                                tran.eTranResult = TransactionResult.REJECTED;
                        }
                        #endregion

                        #region Get RRN
                        else if (line.Contains(csPGRRN)) // PG
                        {
                            tran.strRRN = line.Substring(5, 12);
                        }
                        else if (line.Contains(csVBARRN)) // VBA
                        {
                            tran.strRRN = line.Substring(csVBARRN.Length + 1);
                        }
                        else if (line.Contains(csSHBRRN)) // SHB
                        {
                            tran.strRRN = line.Substring(line.IndexOf(csSHBRRN) + 6);
                        }
                        #endregion

                        #region Get TX Type
                        // DN200-400 Deposit TX will be set by the "BANKNOTES DETECTED"
                        else if (line.Contains(csSHBTranType)) // PG
                        {
                            if (line.Contains(csCashWithdrawal))
                            {
                                tran.eTranType = TransactionType.DISPENSE;
                                continue;
                            }
                            else if (line.Contains(csMini))
                            {
                                tran.eTranType = TransactionType.MINI;
                                continue;
                            }
                            else if (line.Contains(csBalance))
                            {
                                tran.eTranType = TransactionType.BALANCE;
                                continue;
                            }
                            continue;
                        }
                        else if (line.Contains(csVBAPinChange))
                        {
                            tran.eTranType = TransactionType.PINCHANGE;
                            continue;
                        }
                        else if (line.Contains(csVBABalance))
                        {
                            tran.eTranType = TransactionType.BALANCE;
                            continue;
                        }
                        #endregion
                    }
                    Console.WriteLine("Read {0} files done!", filePath);
                }
                catch (Exception ex)
                {
                    throw new Exception("Current File: " + filePath, ex);
                }
            }
            Console.WriteLine("Read all files done!");



            #region write to excell file
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

                //oSheet.Activate();
                oSheet.Application.ActiveWindow.SplitRow = 1;
                oSheet.Application.ActiveWindow.FreezePanes = true;

                Excel.Range formatRange;
                formatRange = oSheet.get_Range("a1");
                formatRange.EntireRow.Font.Bold = true;

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
                    oSheet.Cells[index, 1] = index-1;
                    oSheet.Cells[index, 2] = currentTran.strDate;
                    oSheet.Cells[index, 3] = currentTran.strStartTime;
                    oSheet.Cells[index, 4] = currentTran.strEndTime;
                    oSheet.Cells[index, 5] = currentTran.strCardNum;
                    oSheet.Cells[index, 6] = currentTran.strRRN;
                    oSheet.Cells[index, 7] = currentTran.strOpCode;
                    oSheet.Cells[index, 8] = currentTran.strAmount;
                    oSheet.Cells[index, 9] = currentTran.strDispRequest;
                    oSheet.Cells[index, 10] = currentTran.arCashDisp[0];
                    oSheet.Cells[index, 11] = currentTran.arCashDisp[1];
                    oSheet.Cells[index, 12] = currentTran.arCashDisp[2];
                    oSheet.Cells[index, 13] = currentTran.arCashDisp[3];
                    oSheet.Cells[index, 14] = currentTran.eTranType.ToString();
                    oSheet.Cells[index, 15] = currentTran.eTranResult.ToString();
                    index++;
                }
                //oWB.SaveAs(fpath + "Report.xls");
                oWB.SaveAs(fpath + "\\Report.xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                // oWB.Save();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception :" + ex);
            }
            finally
            {
                if (oWB != null)
                    oWB.Close();
            }
            #endregion
        }
    }
}
