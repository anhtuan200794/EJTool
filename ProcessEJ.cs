using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

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
        const string csCashIn = "CASH IN OK";
        const string csRRN = "RRN:[";
        const string csVBARRN = "TRAN.ID";
        const string csTranType = "*******";
        const string csCashWithdrawal = "CASH WITHDRAWAL";
        const string csBalance = "BALANCE INQ";
        const string csMini = "MINI STATEMENTS";
        const string csVBAAmount = "WITHDRAWAL";
        const string csVBAResCode = "RESPONSE CODE";
        const string csVBABalance = "BALANCE INQUIRY";
        const string csVBACurrency = "00 VND";
        const string csVBAPinChange = "PIN CHANGE";
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
                    string line;
                    while ((line = stReader.ReadLine()) != null)
                    {

                        //}
                        //foreach (string line in File.ReadLines(filePath, Encoding.UTF8))
                        //{
                        //    // process the line
                        //}
                        //while (!stReader.EndOfStream)
                        //{
                        //line = stReader.ReadLine();
                        index++;
                        //Console.WriteLine(index + ":" + line);

                        if (String.IsNullOrEmpty(line))
                        {
                            continue;
                        }
                        #region TX Start
                        else if (line.Contains(csTxStart))
                        {
                            tran = new Transaction();
                            tran.ResetAllData();
                            tran.strStartTime = line.Substring(0, 8);
                            continue;
                        }
                        #endregion

                        #region TX end
                        else if (line.Contains(csTxEnd))
                        {
                            tran.strEndTime = line.Substring(0, 8);
                            //if (tran.eTranType == TransactionType.DISPENSE)
                            //{
                            //    if (tran.eTranResult == TransactionResult.SUCCESS)
                            //    {
                            //        tran.bIsTranSuccess = true;
                            //    }
                            //}
                            tran.strDate = strDate;
                            transactions.Add(tran);
                            continue;
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
                            continue;
                        }
                        else if (line.Contains(csVBAAmount) && line.Contains(csVBACurrency))// VBA
                        {
                            tran.eTranType = TransactionType.DISPENSE;
                            tran.strAmount = line.Substring(csVBAAmount.Length, line.Length - line.IndexOf('.') + 1);
                            continue;
                        }
                        #endregion

                        #region Get TX result
                        else if (line.Contains(csCashRetracted))
                        {
                            tran.eTranResult = TransactionResult.RETRACTED;
                            continue;
                        }
                        else if (line.Contains(csCashTaken))
                        {
                            tran.eTranResult = TransactionResult.SUCCESS;
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
                        else if (line.Contains(csRRN)) // PG
                        {
                            tran.strRRN = line.Substring(5, 12);
                        }
                        else if (line.Contains(csVBARRN)) // VBA
                        {
                            tran.strRRN = line.Substring(csVBARRN.Length + 1);
                        }
                        #endregion

                        #region Get TX Type
                        // DN200-400 Deposit TX will be set by the "BANKNOTES DETECTED"
                        else if (line.Contains(csTranType)) // PG
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
                        #region Get number note for Deposit TX
                        if (line.Contains(csCashIn))
                        {
                            tran.eTranType = TransactionType.DEPOSIT;
                            // Skip 2 next line to get the number of note for each denomination.
                            line = stReader.ReadLine();
                            line = stReader.ReadLine();
                            index += 2;

                        GetNumOfNote:
                            line = stReader.ReadLine();
                            index++;
                            string[] temp = line.Split(' ');
                            // Eg: 17:13:36 VND 100000 * 23
                            if (!temp[1].Contains("VND"))
                                continue;
                            Int32.TryParse(temp[4], out int num);
                            if (temp[2] == "50000")
                            {
                                tran.arNumOfNote[0] = tran.arNumOfNote[0] + num;
                            }
                            else if (temp[2] == "100000")
                            {
                                tran.arNumOfNote[1] = tran.arNumOfNote[1] + num;
                            }
                            else if (temp[2] == "200000")
                            {
                                tran.arNumOfNote[2] = tran.arNumOfNote[2] + num;
                            }
                            else if (temp[2] == "500000")
                            {
                                tran.arNumOfNote[3] = tran.arNumOfNote[3] + num;
                            }
                            goto GetNumOfNote;
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
                oSheet.Cells[1, 14] = "50K";
                oSheet.Cells[1, 15] = "100K";
                oSheet.Cells[1, 16] = "200K";
                oSheet.Cells[1, 17] = "500K";
                oSheet.Cells[1, 18] = "TRAN TYPE";
                oSheet.Cells[1, 19] = "RESULT";

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
                    oSheet.Cells[index, 9] = currentTran.strDispRequest;
                    oSheet.Cells[index, 10] = currentTran.arCashDisp[0];
                    oSheet.Cells[index, 11] = currentTran.arCashDisp[1];
                    oSheet.Cells[index, 12] = currentTran.arCashDisp[2];
                    oSheet.Cells[index, 13] = currentTran.arCashDisp[3];
                    oSheet.Cells[index, 14] = currentTran.arNumOfNote[0];
                    oSheet.Cells[index, 15] = currentTran.arNumOfNote[1];
                    oSheet.Cells[index, 16] = currentTran.arNumOfNote[2];
                    oSheet.Cells[index, 17] = currentTran.arNumOfNote[3];
                    oSheet.Cells[index, 18] = currentTran.eTranType.ToString();
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
                    oSheet.Cells[index, 19] = currentTran.eTranResult.ToString();
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
