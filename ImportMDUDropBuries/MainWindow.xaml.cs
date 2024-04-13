/* Title:       Import MDU Drop Buries - Main Window
 * Date:        9-7-17
 * Author:      Terry Holmes */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using NewEventLogDLL;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using CustomersDLL;
using DataValidationDLL;
using DateSearchDLL;
using DropBuryMDUDLL;
using KeyWordDLL;
using WorkOrderDLL;
using WorkTypeDLL;

namespace ImportMDUDropBuries
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //setting up the classes
        WPFMessagesClass TheMessagesClass = new WPFMessagesClass();
        EventLogClass TheEventLogClass = new EventLogClass();
        CustomersClass TheCustomersClass = new CustomersClass();
        DataValidationClass TheDataValidationClass = new DataValidationClass();
        DateSearchClass TheDateSearchClass = new DateSearchClass();
        DropBuryMDUClass TheDropBuryMDUClass = new DropBuryMDUClass();
        KeyWordClass TheKeyWordClass = new KeyWordClass();
        WorkOrderClass TheWorkOrderClass = new WorkOrderClass();
        WorkTypeClass TheWorkTypeClass = new WorkTypeClass();

        //created Data Sets
        ImportedJobsDataSet TheImportedJobsDataSet = new ImportedJobsDataSet();
        DataReadyDataSet TheMDUDataReadyDataSet = new DataReadyDataSet();
        DataReadyDataSet TheDropDataReadyDataSet = new DataReadyDataSet();
        FindCustomerAddressDateMatchDataSet TheFindCustomerAddressDateMatchDataSet = new FindCustomerAddressDateMatchDataSet();
        FindCustomerByPhoneNumberDataSet TheFindCustomerByPhoneNumberDataSet = new FindCustomerByPhoneNumberDataSet();
        FindCustomersByAddressIDDataSet TheFindCustomersByAddressIDDataSet = new FindCustomersByAddressIDDataSet();
        FindActiveCustomerByAccountNumberDataSet TheFindActiveCustomerByAccountNumberDataSet = new FindActiveCustomerByAccountNumberDataSet();
        FindCustomerByAccountNumberDataSet TheFindCustomerByAccountNumberDataSet = new FindCustomerByAccountNumberDataSet();
        FindWorkOrderByWorkOrderNumberDataSet TheFindWorkOrderByWorkOrderNumberDataSet = new FindWorkOrderByWorkOrderNumberDataSet();
        FindAddressByAddressesDataSet TheFindAddressByAddressesDataSet = new FindAddressByAddressesDataSet();
        
        //stored procedures
        FindWorkZoneByZoneNameDataSet TheFindWorkZoneByZoneNameDataSet = new FindWorkZoneByZoneNameDataSet();
        CustomersDataSet TheCustomersDataSet = new CustomersDataSet();
        
        string gstrErrorMessage;
        int gintPhoneCounter;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            TheMessagesClass.CloseTheProgram();
        }

        private void Grid_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DragMove();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            //this will load up the control
            bool blnFatalError = false;

            PleaseWait PleaseWait = new PleaseWait();
            PleaseWait.Show();

            TheCustomersDataSet = TheCustomersClass.GetCustomersInfo();

            gintPhoneCounter = TheCustomersDataSet.customers.Rows.Count;

            //work goes here
            blnFatalError = LoadMDUGrid();
            if (blnFatalError == false)
                blnFatalError = LoadDropBuryGrid();

            PleaseWait.Close();

            if(blnFatalError == true)
            {
                TheMessagesClass.ErrorMessage(gstrErrorMessage);

                btnProcess.IsEnabled = false;
            }
        }
        private bool LoadDropBuryGrid()
        {
            bool blnFatalError = false;
            Excel.Application xlDropOrder;
            Excel.Workbook xlDropBook;
            Excel.Worksheet xlDropSheet;
            Excel.Range range;

            string strInformation = "";
            int intRowCount;
            int intColumnCount;
            int intRowRange;
            int intColumnRange = 0;
            DateTime datScheduledDate;
            int intCounter;
            int intNumberOfRecords;
            int intRecordsReturned;
            int intZoneID;
            string strZipCode;
            string strZone;

            try
            {
                TheImportedJobsDataSet.importedjobs.Rows.Clear();

                xlDropOrder = new Excel.Application();
                xlDropBook = xlDropOrder.Workbooks.Open(@"c:\users\tholmes\desktop\dropbury.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlDropSheet = (Excel.Worksheet)xlDropOrder.Worksheets.get_Item(1);

                range = xlDropSheet.UsedRange;
                intRowRange = range.Rows.Count - 1;
                intColumnRange = range.Columns.Count;


                for (intRowCount = 1; intRowCount <= intRowRange; intRowCount++)
                {
                    ImportedJobsDataSet.importedjobsRow NewImportedRow = TheImportedJobsDataSet.importedjobs.NewimportedjobsRow();

                    datScheduledDate = DateTime.Now;

                    for (intColumnCount = 1; intColumnCount <= intColumnRange; intColumnCount++)
                    {
                        strInformation = Convert.ToString((range.Cells[intRowCount, intColumnCount] as Excel.Range).Value2);

                        if (intColumnCount == 1)
                            NewImportedRow.AccountID = strInformation.ToUpper();
                        else if (intColumnCount == 2)
                            NewImportedRow.WorkOrderID = strInformation.ToUpper();
                        else if (intColumnCount == 3)
                            NewImportedRow.JobStatus = strInformation.ToUpper();
                        else if (intColumnCount == 4)
                        {
                            datScheduledDate = DateTime.FromOADate(Convert.ToDouble(strInformation));

                            NewImportedRow.ScheduledDate = datScheduledDate;
                        }

                        else if (intColumnCount == 5)
                            NewImportedRow.Pool = strInformation.ToUpper();
                        else if (intColumnCount == 6)
                        {
                            if(strInformation == null)
                            {
                                strInformation = "NOT PROVIDED";
                            }

                            NewImportedRow.FirstName = strInformation.ToUpper();
                        }
                            
                        else if (intColumnCount == 7)
                            NewImportedRow.LastName = strInformation.ToUpper();
                        else if (intColumnCount == 8)
                        {
                            if (strInformation == null)
                            {
                                strInformation = Convert.ToString(gintPhoneCounter);
                                gintPhoneCounter++;
                            }
                            else if (strInformation == "(999) 999-9999")
                            {
                                strInformation = Convert.ToString(gintPhoneCounter);
                                gintPhoneCounter++;
                            }

                            NewImportedRow.PhoneNumber = strInformation.ToUpper();
                        }
                            
                        else if (intColumnCount == 9)
                            NewImportedRow.StreetAddress = strInformation.ToUpper();
                        else if (intColumnCount == 10)
                            NewImportedRow.City = strInformation.ToUpper();
                        else if (intColumnCount == 11)
                            NewImportedRow.State = strInformation.ToUpper();
                        else if (intColumnCount == 12)
                            NewImportedRow.ZipCode = strInformation.ToUpper();
                    }

                    TheImportedJobsDataSet.importedjobs.Rows.Add(NewImportedRow);
                }

                xlDropBook.Close(true, null, null);
                xlDropOrder.Quit();

                Marshal.ReleaseComObject(xlDropSheet);
                Marshal.ReleaseComObject(xlDropBook);
                Marshal.ReleaseComObject(xlDropOrder);

                intNumberOfRecords = TheImportedJobsDataSet.importedjobs.Rows.Count - 1;

                for (intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    if (TheImportedJobsDataSet.importedjobs[intCounter].JobStatus == "ASSIGNED")
                    {
                        strZipCode = TheImportedJobsDataSet.importedjobs[intCounter].ZipCode;

                        if (strZipCode.Length > 5)
                        {
                            strZipCode = strZipCode.Substring(0, 5);
                        }

                        strZone = TheImportedJobsDataSet.importedjobs[intCounter].Pool;

                        strZone = strZone.Substring(2);

                        if (strZone == "MAPLE")
                        {
                            strZone = "MAPLE HTS";
                        }

                        TheFindWorkZoneByZoneNameDataSet = TheCustomersClass.FindWorkZoneByZoneName(strZone);

                        intRecordsReturned = TheFindWorkZoneByZoneNameDataSet.FindWorkZoneByZoneName.Rows.Count;

                        if (intRecordsReturned == 0)
                        {
                            strZone = "UNKNOWN";
                            intZoneID = 1010;
                        }
                        else
                        {
                            intZoneID = TheFindWorkZoneByZoneNameDataSet.FindWorkZoneByZoneName[0].ZoneID;
                        }

                        DataReadyDataSet.insertdataRow NewJobWork = TheDropDataReadyDataSet.insertdata.NewinsertdataRow();

                        NewJobWork.StreetAddress = TheImportedJobsDataSet.importedjobs[intCounter].StreetAddress;
                        NewJobWork.ZoneID = intZoneID;
                        NewJobWork.WorkZone = strZone;
                        NewJobWork.City = TheImportedJobsDataSet.importedjobs[intCounter].City;
                        NewJobWork.State = TheImportedJobsDataSet.importedjobs[intCounter].State;
                        NewJobWork.ZipCode = strZipCode;
                        NewJobWork.PhoneNumber = TheImportedJobsDataSet.importedjobs[intCounter].PhoneNumber;
                        NewJobWork.AccountNumber = TheImportedJobsDataSet.importedjobs[intCounter].AccountID;
                        NewJobWork.WorkOrderNumber = TheImportedJobsDataSet.importedjobs[intCounter].WorkOrderID;
                        NewJobWork.WorkTypeID = 1002;
                        NewJobWork.DateEntered = DateTime.Now;
                        NewJobWork.DateScheduled = TheImportedJobsDataSet.importedjobs[intCounter].ScheduledDate;
                        NewJobWork.DateReceived = DateTime.Now;
                        NewJobWork.StatusDate = DateTime.Now;
                        NewJobWork.StatusID = 1001;
                        NewJobWork.FirstName = TheImportedJobsDataSet.importedjobs[intCounter].FirstName;
                        NewJobWork.LastName = TheImportedJobsDataSet.importedjobs[intCounter].LastName;

                        TheDropDataReadyDataSet.insertdata.Rows.Add(NewJobWork);
                    }
                }

                dgrDropBury.ItemsSource = TheDropDataReadyDataSet.insertdata;
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Import MDU Drop Buries // Load Drop Bury Grid " + Ex.Message);

                gstrErrorMessage = Ex.ToString();

                blnFatalError = true;
            }

            return blnFatalError;
        }
        private bool LoadMDUGrid()
        {
            bool blnFatalError = false;
            Excel.Application xlMDUOrder;
            Excel.Workbook xlMDUBook;
            Excel.Worksheet xlMDUSheet;
            Excel.Range range;

            string strInformation = "";
            int intRowCount;
            int intColumnCount;
            int intRowRange;
            int intColumnRange = 0;
            DateTime datScheduledDate;
            int intCounter;
            int intNumberOfRecords;
            string strZipCode;
            string strZone;
            int intRecordsReturned;
            int intZoneID = 0;
            
            try
            {
                TheImportedJobsDataSet.importedjobs.Rows.Clear();

                xlMDUOrder = new Excel.Application();
                xlMDUBook = xlMDUOrder.Workbooks.Open(@"c:\users\tholmes\desktop\mdu.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlMDUSheet = (Excel.Worksheet)xlMDUOrder.Worksheets.get_Item(1);

                range = xlMDUSheet.UsedRange;
                intRowRange = range.Rows.Count - 1;
                intColumnRange = range.Columns.Count;


                for (intRowCount = 1; intRowCount <= intRowRange; intRowCount++)
                {
                    ImportedJobsDataSet.importedjobsRow NewImportedRow = TheImportedJobsDataSet.importedjobs.NewimportedjobsRow();

                    datScheduledDate = DateTime.Now;

                    for (intColumnCount = 1; intColumnCount <= intColumnRange; intColumnCount++)
                    {
                        strInformation = Convert.ToString((range.Cells[intRowCount, intColumnCount] as Excel.Range).Value2);

                        if (intColumnCount == 1)
                            NewImportedRow.AccountID = strInformation.ToUpper();
                        else if (intColumnCount == 2)
                            NewImportedRow.WorkOrderID = strInformation.ToUpper();
                        else if (intColumnCount == 3)
                            NewImportedRow.JobStatus = strInformation.ToUpper();
                        else if (intColumnCount == 4)
                        {
                            datScheduledDate = DateTime.FromOADate(Convert.ToDouble(strInformation));

                            NewImportedRow.ScheduledDate = datScheduledDate;
                        }

                        else if (intColumnCount == 5)
                            NewImportedRow.Pool = strInformation.ToUpper();
                        else if (intColumnCount == 6)
                            NewImportedRow.FirstName = strInformation.ToUpper();
                        else if (intColumnCount == 7)
                            NewImportedRow.LastName = strInformation.ToUpper();
                        else if (intColumnCount == 8)
                        {
                            if (strInformation == null)
                            {
                                strInformation = Convert.ToString(gintPhoneCounter);
                                gintPhoneCounter++;
                            }
                            else if(strInformation == "(999) 999-9999")
                            {
                                strInformation = Convert.ToString(gintPhoneCounter);
                                gintPhoneCounter++;
                            }
                            
                            NewImportedRow.PhoneNumber = strInformation.ToUpper();
                            
                        }
                        else if (intColumnCount == 9)
                            NewImportedRow.StreetAddress = strInformation.ToUpper();
                        else if (intColumnCount == 10)
                            NewImportedRow.City = strInformation.ToUpper();
                        else if (intColumnCount == 11)
                            NewImportedRow.State = strInformation.ToUpper();
                        else if (intColumnCount == 12)
                            NewImportedRow.ZipCode = strInformation.ToUpper();
                    }

                    TheImportedJobsDataSet.importedjobs.Rows.Add(NewImportedRow);
                }

                xlMDUBook.Close(true, null, null);
                xlMDUOrder.Quit();

                Marshal.ReleaseComObject(xlMDUSheet);
                Marshal.ReleaseComObject(xlMDUBook);
                Marshal.ReleaseComObject(xlMDUOrder);

                intNumberOfRecords = TheImportedJobsDataSet.importedjobs.Rows.Count - 1;

                for(intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    if(TheImportedJobsDataSet.importedjobs[intCounter].JobStatus == "ASSIGNED")
                    {
                        strZipCode = TheImportedJobsDataSet.importedjobs[intCounter].ZipCode;

                        if(strZipCode.Length > 5)
                        {
                            strZipCode = strZipCode.Substring(0, 5);
                        }

                        strZone = TheImportedJobsDataSet.importedjobs[intCounter].Pool;

                        strZone = strZone.Substring(2);

                        if(strZone == "MAPLE")
                        {
                            strZone = "MAPLE HTS";
                        }

                        TheFindWorkZoneByZoneNameDataSet = TheCustomersClass.FindWorkZoneByZoneName(strZone);

                        intRecordsReturned = TheFindWorkZoneByZoneNameDataSet.FindWorkZoneByZoneName.Rows.Count;

                        if(intRecordsReturned == 0)
                        {
                            strZone = "UNKNOWN";
                            intZoneID = 1010;
                        }
                        else
                        {
                            intZoneID = TheFindWorkZoneByZoneNameDataSet.FindWorkZoneByZoneName[0].ZoneID;
                        }

                        DataReadyDataSet.insertdataRow NewJobWork = TheMDUDataReadyDataSet.insertdata.NewinsertdataRow();

                        NewJobWork.StreetAddress = TheImportedJobsDataSet.importedjobs[intCounter].StreetAddress;
                        NewJobWork.ZoneID = intZoneID;
                        NewJobWork.WorkZone = strZone;
                        NewJobWork.City = TheImportedJobsDataSet.importedjobs[intCounter].City;
                        NewJobWork.State = TheImportedJobsDataSet.importedjobs[intCounter].State;
                        NewJobWork.ZipCode = strZipCode;
                        NewJobWork.PhoneNumber = TheImportedJobsDataSet.importedjobs[intCounter].PhoneNumber;
                        NewJobWork.AccountNumber = TheImportedJobsDataSet.importedjobs[intCounter].AccountID;
                        NewJobWork.WorkOrderNumber = TheImportedJobsDataSet.importedjobs[intCounter].WorkOrderID;
                        NewJobWork.WorkTypeID = 1001;
                        NewJobWork.DateEntered = DateTime.Now;
                        NewJobWork.DateScheduled = TheImportedJobsDataSet.importedjobs[intCounter].ScheduledDate;
                        NewJobWork.DateReceived = DateTime.Now;
                        NewJobWork.StatusDate = DateTime.Now;
                        NewJobWork.StatusID = 1001;
                        NewJobWork.FirstName = TheImportedJobsDataSet.importedjobs[intCounter].FirstName;
                        NewJobWork.LastName = TheImportedJobsDataSet.importedjobs[intCounter].LastName;

                        TheMDUDataReadyDataSet.insertdata.Rows.Add(NewJobWork);
                    }
                }

                dgrMDU.ItemsSource = TheMDUDataReadyDataSet.insertdata;
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Import MDU Drop Buries // Load MDU Grid " + Ex.Message);

                gstrErrorMessage = Ex.ToString();

                blnFatalError = true;
            }

            return blnFatalError;
        }

        private void btnProcess_Click(object sender, RoutedEventArgs e)
        {
            //this will insert the jobs into the three tables
            PleaseWait PleaseWait = new PleaseWait();
            PleaseWait.Show();

            InsertMDUJobs();

            InsertDropJobs();

            PleaseWait.Close();
        }
        private void InsertDropJobs()
        {
            int intCounter;
            int intNumberOfRecords;
            string strAddress;
            string strCity;
            string strZip;
            string strPhoneNumber;
            string strAccountNumber;
            int intRecordsReturned;
            bool blnAddressFound;
            bool blnCustomerFound;
            string strWorkOrderNumber;
            bool blnWorkOrderFound;
            int intAddressCounter;
            bool blnFatalError;
            DateTime datTransactionDate;
            int intAddressID = 0;
            int intCustomerID = 0;
            string strFirstName;
            string strLastName;
            string strState;

            try
            {
                intNumberOfRecords = TheDropDataReadyDataSet.insertdata.Rows.Count - 1;

                for (intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    blnAddressFound = false;
                    blnCustomerFound = false;
                    blnWorkOrderFound = false;
                    datTransactionDate = DateTime.Now;

                    //loading the variabvles
                    strAccountNumber = TheDropDataReadyDataSet.insertdata[intCounter].AccountNumber;
                    strAddress = TheDropDataReadyDataSet.insertdata[intCounter].StreetAddress;
                    strCity = TheDropDataReadyDataSet.insertdata[intCounter].City;
                    strZip = TheDropDataReadyDataSet.insertdata[intCounter].ZipCode;
                    strPhoneNumber = TheDropDataReadyDataSet.insertdata[intCounter].PhoneNumber;
                    strWorkOrderNumber = TheDropDataReadyDataSet.insertdata[intCounter].WorkOrderNumber;
                    strFirstName = TheDropDataReadyDataSet.insertdata[intCounter].FirstName;
                    strLastName = TheDropDataReadyDataSet.insertdata[intCounter].LastName;
                    strState = TheDropDataReadyDataSet.insertdata[intCounter].State;

                    TheFindAddressByAddressesDataSet = TheCustomersClass.FindAddressesByAddress(strAddress);

                    intRecordsReturned = TheFindAddressByAddressesDataSet.FindAddressesByAddress.Rows.Count;

                    if (intRecordsReturned > 0)
                    {
                        intRecordsReturned -= 1;

                        for (intAddressCounter = 0; intAddressCounter <= intRecordsReturned; intAddressCounter++)
                        {
                            if (strZip == TheFindAddressByAddressesDataSet.FindAddressesByAddress[0].ZipCode)
                            {
                                blnAddressFound = true;
                                intAddressID = TheFindAddressByAddressesDataSet.FindAddressesByAddress[0].AddressID;
                            }
                        }

                    }

                    TheFindCustomerByAccountNumberDataSet = TheCustomersClass.FindCustomerByAccountNumber(strAccountNumber);

                    intRecordsReturned = TheFindCustomerByAccountNumberDataSet.FindCustomerByAccountNumber.Rows.Count;

                    if (intRecordsReturned > 0)
                    {
                        blnCustomerFound = true;
                    }
                    else
                    {
                        TheFindCustomerByPhoneNumberDataSet = TheCustomersClass.FindCustomerByPhoneNumber(strPhoneNumber);

                        intRecordsReturned = TheFindCustomerByPhoneNumberDataSet.FindCustomerByPhoneNumber.Rows.Count;

                        if (intRecordsReturned > 0)
                        {
                            blnCustomerFound = true;
                        }
                    }

                    TheFindWorkOrderByWorkOrderNumberDataSet = TheWorkOrderClass.FindWorkOrderByWorkOrderNumber(strWorkOrderNumber);

                    intRecordsReturned = TheFindWorkOrderByWorkOrderNumberDataSet.FindWorkOrderByWorkOrderNumber.Rows.Count;

                    if (intRecordsReturned > 0)
                    {
                        blnWorkOrderFound = true;
                    }

                    if (blnAddressFound == false)
                    {
                        blnFatalError = TheCustomersClass.InsertCustomerAddress(strAddress, strCity, strState, TheDropDataReadyDataSet.insertdata[intCounter].ZoneID, strZip, datTransactionDate);

                        if (blnFatalError == true)
                        {
                            throw new Exception();
                        }
                    }
                    if (blnCustomerFound == false)
                    {
                        if (blnAddressFound == false)
                        {
                            TheFindCustomerAddressDateMatchDataSet = TheCustomersClass.FindCustomerAddressDateMatch(datTransactionDate);

                            intAddressID = TheFindCustomerAddressDateMatchDataSet.FindCustomerAddressesDateMatch[0].AddressID;
                        }

                        blnFatalError = TheCustomersClass.InsertCustomer(intAddressID, strPhoneNumber, strAccountNumber, strFirstName, strLastName);

                        if (blnFatalError == true)
                        {
                            throw new Exception();
                        }
                    }
                    if (blnWorkOrderFound == false)
                    {
                        TheFindActiveCustomerByAccountNumberDataSet = TheCustomersClass.FindActiveCustomerByAccountNumber(strAccountNumber);

                        intCustomerID = TheFindActiveCustomerByAccountNumberDataSet.FindActiveCustomerByAccountNumber[0].CustomerID;

                        blnFatalError = TheWorkOrderClass.InsertWorkOrder(strWorkOrderNumber, TheDropDataReadyDataSet.insertdata[intCounter].WorkTypeID, intCustomerID, intAddressID, TheDropDataReadyDataSet.insertdata[intCounter].DateScheduled, TheDropDataReadyDataSet.insertdata[intCounter].DateReceived, TheDropDataReadyDataSet.insertdata[intCounter].StatusID);

                        if (blnFatalError == true)
                        {
                            throw new Exception();
                        }

                        TheFindWorkOrderByWorkOrderNumberDataSet = TheWorkOrderClass.FindWorkOrderByWorkOrderNumber(strWorkOrderNumber);

                        blnFatalError = TheWorkOrderClass.InsertWorkOrderUpdate(TheFindWorkOrderByWorkOrderNumberDataSet.FindWorkOrderByWorkOrderNumber[0].WorkOrderID, 20007, "IMPORTED FROM SPREADSHEET");
                    }
                }
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Import MDU Drop Buries // Insert MDU Jobs " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }

        }
        private void InsertMDUJobs()
        {
            int intCounter;
            int intNumberOfRecords;
            string strAddress;
            string strCity;
            string strZip;
            string strPhoneNumber;
            string strAccountNumber;
            int intRecordsReturned;
            bool blnAddressFound;
            bool blnCustomerFound;
            string strWorkOrderNumber;
            bool blnWorkOrderFound;
            int intAddressCounter;
            bool blnFatalError;
            DateTime datTransactionDate;
            int intAddressID = 0;
            int intCustomerID = 0;
            string strFirstName;
            string strLastName;
            string strState;

            try
            {
                intNumberOfRecords = TheMDUDataReadyDataSet.insertdata.Rows.Count - 1;

                for(intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    blnAddressFound = false;
                    blnCustomerFound = false;
                    blnWorkOrderFound = false;
                    datTransactionDate = DateTime.Now;

                    //loading the variabvles
                    strAccountNumber = TheMDUDataReadyDataSet.insertdata[intCounter].AccountNumber;
                    strAddress = TheMDUDataReadyDataSet.insertdata[intCounter].StreetAddress;
                    strCity = TheMDUDataReadyDataSet.insertdata[intCounter].City;
                    strZip = TheMDUDataReadyDataSet.insertdata[intCounter].ZipCode;
                    strPhoneNumber = TheMDUDataReadyDataSet.insertdata[intCounter].PhoneNumber;
                    strWorkOrderNumber = TheMDUDataReadyDataSet.insertdata[intCounter].WorkOrderNumber;
                    strFirstName = TheMDUDataReadyDataSet.insertdata[intCounter].FirstName;
                    strLastName = TheMDUDataReadyDataSet.insertdata[intCounter].LastName;
                    strState = TheMDUDataReadyDataSet.insertdata[intCounter].State;

                    TheFindAddressByAddressesDataSet = TheCustomersClass.FindAddressesByAddress(strAddress);

                    intRecordsReturned = TheFindAddressByAddressesDataSet.FindAddressesByAddress.Rows.Count;

                    if(intRecordsReturned > 0)
                    {
                        intRecordsReturned -= 1;

                        for(intAddressCounter = 0; intAddressCounter <= intRecordsReturned; intAddressCounter++)
                        {
                            if(strZip == TheFindAddressByAddressesDataSet.FindAddressesByAddress[0].ZipCode)
                            {
                                blnAddressFound = true;
                                intAddressID = TheFindAddressByAddressesDataSet.FindAddressesByAddress[0].AddressID;
                            }
                        }

                    }

                    TheFindCustomerByAccountNumberDataSet = TheCustomersClass.FindCustomerByAccountNumber(strAccountNumber);

                    intRecordsReturned = TheFindCustomerByAccountNumberDataSet.FindCustomerByAccountNumber.Rows.Count;

                    if(intRecordsReturned > 0)
                    {
                        blnCustomerFound = true;
                    }
                    else
                    {
                        TheFindCustomerByPhoneNumberDataSet = TheCustomersClass.FindCustomerByPhoneNumber(strPhoneNumber);

                        intRecordsReturned = TheFindCustomerByPhoneNumberDataSet.FindCustomerByPhoneNumber.Rows.Count;

                        if(intRecordsReturned > 0)
                        {
                            blnCustomerFound = true;
                            strAccountNumber = TheFindCustomerByPhoneNumberDataSet.FindCustomerByPhoneNumber[0].AccountNumber;
                        }                       
                    }

                    TheFindWorkOrderByWorkOrderNumberDataSet = TheWorkOrderClass.FindWorkOrderByWorkOrderNumber(strWorkOrderNumber);

                    intRecordsReturned = TheFindWorkOrderByWorkOrderNumberDataSet.FindWorkOrderByWorkOrderNumber.Rows.Count;

                    if(intRecordsReturned > 0)
                    {
                        blnWorkOrderFound = true;
                    }

                    if(blnAddressFound == false)
                    {
                        blnFatalError = TheCustomersClass.InsertCustomerAddress(strAddress, strCity, strState, TheMDUDataReadyDataSet.insertdata[intCounter].ZoneID, strZip, datTransactionDate);

                        if(blnFatalError == true)
                        {
                            throw new Exception();
                        }
                    }
                    if(blnCustomerFound == false)
                    {
                        if(blnAddressFound == false)
                        {
                            TheFindCustomerAddressDateMatchDataSet = TheCustomersClass.FindCustomerAddressDateMatch(datTransactionDate);

                            intAddressID = TheFindCustomerAddressDateMatchDataSet.FindCustomerAddressesDateMatch[0].AddressID;
                        }

                        blnFatalError = TheCustomersClass.InsertCustomer(intAddressID, strPhoneNumber, strAccountNumber, strFirstName, strLastName);

                        if(blnFatalError == true)
                        {
                            throw new Exception();
                        }
                    }
                    if(blnWorkOrderFound == false)
                    {
                        TheFindActiveCustomerByAccountNumberDataSet = TheCustomersClass.FindActiveCustomerByAccountNumber(strAccountNumber);

                        intCustomerID = TheFindActiveCustomerByAccountNumberDataSet.FindActiveCustomerByAccountNumber[0].CustomerID;

                        blnFatalError = TheWorkOrderClass.InsertWorkOrder(strWorkOrderNumber, TheMDUDataReadyDataSet.insertdata[intCounter].WorkTypeID, intCustomerID, intAddressID, TheMDUDataReadyDataSet.insertdata[intCounter].DateScheduled, TheMDUDataReadyDataSet.insertdata[intCounter].DateReceived, TheMDUDataReadyDataSet.insertdata[intCounter].StatusID);

                        if(blnFatalError == true)
                        {
                            throw new Exception();
                        }

                        TheFindWorkOrderByWorkOrderNumberDataSet = TheWorkOrderClass.FindWorkOrderByWorkOrderNumber(strWorkOrderNumber);

                        blnFatalError = TheWorkOrderClass.InsertWorkOrderUpdate(TheFindWorkOrderByWorkOrderNumberDataSet.FindWorkOrderByWorkOrderNumber[0].WorkOrderID, 20007, "IMPORTED FROM SPREADSHEET");
                    }
                }
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Import MDU Drop Buries // Insert MDU Jobs " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        
        }
    }
}
