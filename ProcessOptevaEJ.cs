using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Windows.Forms;


namespace EJTool
{
    class ProcessOptevaEJ
    {
        const string csTxStart = "Card Inserted (";
        const string csTxEnd = "Card Ejected";
        const string csCardNum = "TRACK 2 DATA:";
        const string csCashRq = "CASH REQUEST:";
        const string csOpCode = "Transaction Request";
        const string csAmount = "AMOUNT         :";
        const string csCashTaken = "CASH TAKEN";
        const string csCashRetracted = "CASH RETRACTED";
        const string csCashIn = "CASH IN OK";
        const string csBABRRN = "RRN:";
        const string csTranType = "*******";
        const string csBABCashWithdrawal = "CASH_WITHDRAWAL";
        const string csBABBalance = "BALANCE_INQ";
        const string csMini = "MINI STATEMENTS";
        const string csBABRc = "RC:";
        const string csDispenseType1 = "Dispense Type1";
        const string csDispenseType2 = "Dispense Type2";
        const string csDispenseType3 = "Dispense Type3";
        const string csDispenseType4 = "Dispense Type4";

        public void ProcessOptevaEJFile(string[] EJFiles)
        {
            List<Transaction> transactions = new List<Transaction>();
            #region Read file to get the data
            Transaction tran = new Transaction();
            StreamReader stReader = null;
            foreach (string filePath in EJFiles)
            {
                int index = 0;
                try
                {
                    stReader = new StreamReader(filePath);
                    index = 0;
                    string strDate = Path.GetFileNameWithoutExtension(filePath).Substring(15, 8);
                    string line;
                    while ((line = stReader.ReadLine()) != null)
                    {
                        index++;
                        //Console.WriteLine(index + ":" + line);
                        if (index == 1203)
                        {
                            Console.WriteLine(index + ":" + line);
                        }
                        if (String.IsNullOrEmpty(line))
                        {
                            continue;
                        }
                        #region TX Start and get card number
                        else if (line.Contains(csTxStart))
                        {
                            if(tran.strStartTime != "") // end the previous transaction
                            {
                                transactions.Add(tran);
                            }

                            tran = new Transaction();
                            tran.strStartTime = line.Substring(25, 8);
                            tran.strCardNum = line.Substring(48);
                            continue;
                        }
                        #endregion

                        #region TX end
                        else if (line.Contains(csTxEnd))
                        {
                            tran.strEndTime = line.Substring(25, 8);
                            tran.strDate = strDate;
                            continue;
                        }
                        #endregion

                        #region Get transaction type
                        else if (line.Contains(csBABCashWithdrawal))
                        {
                            tran.eTranType = TransactionType.DISPENSE;
                            continue;
                        }
                        else if (line.Contains(csBABBalance))
                        {
                            tran.eTranType = TransactionType.BALANCE;
                            continue;
                        }
                        #endregion
                        #region Get Dispense 
                        else if(line.Contains(csDispenseType1))
                        {
                            string temp = line.Substring(51, 2);
                            Int32.TryParse(temp, out int num);
                            tran.arCashDisp[0] = num;
                        }
                        else if (line.Contains(csDispenseType2))
                        {
                            string temp = line.Substring(51, 2);
                            Int32.TryParse(temp, out int num);
                            tran.arCashDisp[1] = num;
                        }
                        else if (line.Contains(csDispenseType3))
                        {
                            string temp = line.Substring(51, 2);
                            Int32.TryParse(temp, out int num);
                            tran.arCashDisp[2] = num;
                        }
                        else if (line.Contains(csDispenseType4))
                        {
                            string temp = line.Substring(51, 2);
                            Int32.TryParse(temp, out int num);
                            tran.arCashDisp[3] = num;
                        }

                        #endregion
                        #region Get RC
                        else if (line.Contains(csBABRc))
                        {
                            tran.strRC = line.Substring(42, 2);
                            continue;
                        }

                        #endregion
                        #region Get OPP Code
                        else if (line.Contains(csOpCode))
                        {
                            tran.strOpCode = line.Substring(55, 8);
                            continue;
                        }
                        #endregion

                        #region Get Amount
                        else if (line.Contains(csAmount))
                        {
                            tran.strAmount = line.Substring(53, 12);
                            continue;
                        }
                        #endregion

                        #region Get TX result
                        //else if (line.Contains(csCashRetracted))
                        //{
                        //    tran.eTranResult = TransactionResult.RETRACTED;
                        //    continue;
                        //}
                        //else if (line.Contains(csCashTaken))
                        //{
                        //    tran.eTranResult = TransactionResult.SUCCESS;
                        //    continue;
                        //}
                        #endregion

                        #region Get RRN
                        else if (line.Contains(csBABRRN)) // PG
                        {
                            tran.strRRN = line.Substring(43, 12);
                        }
                        #endregion
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Vui lòng chọn thư mục chỉ có file EJ");
                    throw new Exception("Current File: " + filePath, ex);
                }
                Console.WriteLine("Read all files done!");
            }
            #endregion
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
                oSheet.Cells[1, 6] = "RRN";
                oSheet.Columns[6].NumberFormat = "@";
                oSheet.Cells[1, 7] = "OPKEY";
                oSheet.Cells[1, 8] = "AMOUNT";
                oSheet.Cells[1, 9] = "C1";
                oSheet.Cells[1, 10] = "C2";
                oSheet.Cells[1, 11] = "C3";
                oSheet.Cells[1, 12] = "C4";
                oSheet.Cells[1, 13] = "TRAN TYPE";
                oSheet.Cells[1, 14] = "RC";

                int ColIndex = 2;
                foreach (Transaction currentTran in transactions)
                {
                    oSheet.Cells[ColIndex, 1] = ColIndex-1;
                    oSheet.Cells[ColIndex, 2] = currentTran.strDate;
                    oSheet.Cells[ColIndex, 3] = currentTran.strStartTime;
                    oSheet.Cells[ColIndex, 4] = currentTran.strEndTime;
                    oSheet.Cells[ColIndex, 5] = currentTran.strCardNum;
                    oSheet.Cells[ColIndex, 6] = currentTran.strRRN;
                    oSheet.Cells[ColIndex, 7] = currentTran.strOpCode;
                    oSheet.Cells[ColIndex, 8] = currentTran.strAmount;
                    oSheet.Cells[ColIndex, 9] = currentTran.arCashDisp[0];
                    oSheet.Cells[ColIndex, 10] = currentTran.arCashDisp[1];
                    oSheet.Cells[ColIndex, 11] = currentTran.arCashDisp[2];
                    oSheet.Cells[ColIndex, 12] = currentTran.arCashDisp[3];
                    oSheet.Cells[ColIndex, 13] = currentTran.eTranType.ToString();
                    oSheet.Cells[ColIndex, 14] = currentTran.strRC;
                    ColIndex++;
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
