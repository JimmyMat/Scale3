

/***************************************************************


System:     BioWindow -- Data Acquisition, Analysis, and Report Generator


File:       CScale3MainWin.CPP


Procedure:  Implements all of the methods contained in the CScale3MainWin
            class.


Programmer: James D. Matthews


Date:       Oct. 1, 2002



Copyright (C) 2002, Modular Instruments, Inc.

=================================================================

                        Revision Block
                        --------------

   ECO No.     Date     Init.    Procedure and Description of Change
   --------    -------- -----    -----------------------------------


*****************************************************************/

/*****************************************************************
June 12th thru 19th, 2015

So I want to test this code in the real world. I don't have a MT scale or any old style
RS232 instrument to hook up to the Statech device. I do have some Parrllax Basic Stamp twos
and associated training boards from the 1990's. All I have to do is find one.
I found a populated Board of Education, serial cable, and power supply in the attic. Never throw
away anything! I downloaded the latest editor from the Parrallax website and installed it on
the Win8.1 Asus netbook that the Startech serial hub is connected to. (This netbook also acts as my
remote debugging target.) I start the BS2 editor up, power up the Board of Education, (BE), and
connect the BE to the Startech serial hub via the straight thru DB9 serial cable. The Parrallax
editor identifies it. So I'm in business and can program the BE to act as a Mettler Toledo lab scale.

I programmed the BE to wait until it receives the "SI\r\n" string. When it gets the string it sends back
numbers followed by a \r\n combo. Just like the MT lab scale. See my notes above the function that uses
CreateFile for issues and resolutions. But as of June 21st, 2015, I'm satisfied that this code works and
that I should continue with the BioWindow port.

My Test platform consists of a Basic Stamp 2, (BS2), pluged into a Board of Education, (BE). These devices
date from around 1998. The BE is used as a programming conection between the computer and the BS2. It uses
an RS232 serial connection to do this. All BS2 modules have a dedicated serial I/O pin, pin 16. The editor
uses this pin for programming the BS2, but the user is free to use it for RS232 serial communication within
their program. Using pin 16 eliminates the need for external wiring/components to complete an RS232 connection
to one of the 0 thru 15 pins that can also be used but are not internally setup for the RS232 standard. There
is a downside to using the BS2's dedicated serial pin. The downside is that it echos everything it receives. So
the subsequent read buffer will contain the command string in addition to the manufactured data. This echo
doesn't happen in real life but since all I expect to receive from the lab instrument is a numerical string,
I can just accept characters from 0 to 9. Alternatively, I could just filter out the command string or just
ignore the first x characters in the read buffer, where x equals the sizeof the write buffer.


*****************************************************************/


/***************** Controls in Scale3MainWin *********************

ID_OPENING1_SHUTDOWN
ID_OPENING1_PORTTOOLS_CREATEFILE
ID_OPENING1_PORTTOOLS_GETCOMMPROPS
ID_OPENING1_PORTTOOLS_GETCOMMSTATE
ID_OPENING1_PORTTOOLS_SETCOMMSTATE
ID_OPENING1_PORTTOOLS_GETCOMMTIMEOUTS
ID_OPENING1_PORTTOOLS_SETCOMMTIMEOUTS
ID_OPENING1_PORTTOOLS_DOITALL
ID_OPENING1_PORTTOOLS_CLOSEHANDLE
ID_OPENING1_SCALETOOLS_RDWRT_CURRENT
ID_OPENING1_SCALETOOLS_RDWRT_CONFIGURED

m_ComPortLBTitle           IDC_COM_PORT_LB_TITLE
m_ComPortLB                IDC_COM_PORT_LB

m_BaudRateCBTitle          IDC_BAUD_RATE_CB_TITLE
m_BaudRateCB               IDC_BAUD_RATE_CB

m_DataBitsCBTitle          IDC_DATA_BITS_CB_TITLE
m_DataBitsCB               IDC_DATA_BITS_CB

m_StopBitsCBTitle          IDC_STOP_BITS_CB_TITLE
m_StopBitsCB               IDC_STOP_BITS_CB

m_ParityCBTitle            IDC_PARITY_CB_TITLE
m_ParityCB                 IDC_PARITY_CB

m_StatusText1              IDC_STATUS_TEXT_1
m_StatusText2              IDC_STATUS_TEXT_2

m_StructEFTitle            IDC_STRUCT_EF_TITLE
m_StructEF                 IDC_STRUCT_EF

m_ConfiguredLBTitle        IDC_CONFIG_LB_TITLE
m_ConfiguredLB             IDC_CONFIG_LB

m_StartThreadsPB           IDC_STRT_THREAD_PB
m_UseThreadsPB             IDC_USE_THREAD_PB
m_TimerUseThreadsPB        IDC_TIMER_USE_THREAD_PB
m_EndThreadsPB             IDC_END_THREAD_PB

m_ThreadTimerEBTitle       IDC_THREAD_TIMER_EB_TITLE
m_ThreadTimerEB            IDC_THREAD_TIMER_EB

*****************************************************************/

#include <assert.h>
#include "stdafx.h"
#include "Scale3MainWin.hpp"

ULONG ulBaudRates[13] =
{
   110, 300, 600, 1200, 2400, 4800, 9600, 19200, 38400, 57600, 115200, 128000, 256000
};

short sDataBits[4] =
{
   5, 6, 7, 8
};

float fStopBits[3] =
{
   1, 1.5, 2
};

//CString szParity[5] =
//{
   //"NONE", "ODD", "EVEN", "MARK", "SPACE"
//};
CString szParity2[5] =
{
   _T("N"),
   _T("O"),
   _T("E"),
   _T("M"),
   _T("S")
};
CString szParity[5] =
{
   _T("NOPARITY"),
   _T("ODDPARITY"),
   _T("EVENPARITY"),
   _T("MARKPARITY"),
   _T("SPACEPARITY")
};

/*************************************************
This is a major weakness or point of contention in the age of UniCode.

Assuming that the device you're communicating with is using/expecting ANSI instead of
UniCode. For now it seems safe to use ANSI. So I'm using unsigned char in conjunction
with the expectation that the device is
1) sending and receiving a byte at a time
2) each received or sent byte is an ANSI representation of a United States character set
3) all values from 0x00 to 0xff are available.

I don't like this since the code must be modified if it tries to communicate with a
UTF-8, UTF-16 UniCode compliante device.
Also I'm using the /J compilier flag to force char type to be unsigned char as opposed
to the default signed char.
Also by using char type, string and char functions, like strcpy and strlen, are
available for use.

*************************************************/

char szCurrentCmd[] = {"SI\r\n"};

CString COMMPROPMembers[18] =
{
  _T("WORD wPacketLength"),
  _T("WORD wPacketVersion"),
  _T("DWORD dwServiceMask"),
  _T("DWORD dwReserved1"),
  _T("DWORD dwMaxTxQueue"),
  _T("DWORD dwMaxRxQueue"),
  _T("DWORD dwMaxBaud"),
  _T("DWORD dwProvSubType"),
  _T("DWORD dwProvCapabilities"),
  _T("DWORD dwSettableParams"),
  _T("DWORD dwSettableBaud"),
  _T("WORD wSettableData"),
  _T("WORD wSettableStopParity"),
  _T("DWORD dwCurrentTxQueue"),
  _T("DWORD dwCurrentRxQueue"),
  _T("DWORD dwProvSpec1"),
  _T("DWORD dwProvSpec2"),
  _T("WCHAR wcProvChar[1]")
};

CString DCBMembers[28] =
{
  _T("DWORD DCBlength"),
  _T("DWORD BaudRate"),
  _T("DWORD fBinary"),
  _T("DWORD fParity"),
  _T("DWORD fOutxCtsFlow"),
  _T("DWORD fOutxDsrFlow"),
  _T("DWORD fDtrControl"),
  _T("DWORD fDsrSensitivity"),
  _T("DWORD fTXContinueOnXoff"),
  _T("DWORD fOutX"),
  _T("DWORD fInX"),
  _T("DWORD fErrorChar"),
  _T("DWORD fNull"),
  _T("DWORD fRtsControl"),
  _T("DWORD fAbortOnError"),
  _T("DWORD fDummy2"),
  _T("WORD wReserved"),
  _T("WORD XonLim"),
  _T("WORD XoffLim"),
  _T("BYTE ByteSize"),
  _T("BYTE Parity"),
  _T("BYTE StopBits"),
  _T("char XonChar"),
  _T("char XoffChar"),
  _T("char ErrorChar"),
  _T("char EofChar"),
  _T("char EvtChar"),
  _T("WORD wReserved1")
};

CString COMMTIMEOUTSMembers[5] =
{
  _T("DWORD ReadIntervalTimeout"),
  _T("DWORD ReadTotalTimeoutMultiplier"),
  _T("DWORD ReadTotalTimeoutConstant"),
  _T("DWORD WriteTotalTimeoutMultiplier"),
  _T("DWORD WriteTotalTimeoutConstant")
};

BEGIN_MESSAGE_MAP (CScale3MainWin, CFrameWnd)
   ON_WM_CREATE ()
   ON_WM_CLOSE ()
   ON_WM_SIZE ()
   ON_WM_PAINT ()
   ON_WM_TIMER ()
   ON_UPDATE_COMMAND_UI (ID_OPENING1_SHUTDOWN, OnUpdateOpening1ShutDown)
   ON_UPDATE_COMMAND_UI (ID_OPENING1_PORTTOOLS_CREATEFILE, OnUpdateOpening1CreateFile)
   ON_UPDATE_COMMAND_UI (ID_OPENING1_PORTTOOLS_GETCOMMPROPS, OnUpdateOpening1GetCommProps)
   ON_UPDATE_COMMAND_UI (ID_OPENING1_PORTTOOLS_GETCOMMSTATE, OnUpdateOpening1GetCommState)
   ON_UPDATE_COMMAND_UI (ID_OPENING1_PORTTOOLS_SETCOMMSTATE, OnUpdateOpening1SetCommState)
   ON_UPDATE_COMMAND_UI (ID_OPENING1_PORTTOOLS_GETCOMMTIMEOUTS, OnUpdateOpening1GetCommTimeouts)
   ON_UPDATE_COMMAND_UI (ID_OPENING1_PORTTOOLS_SETCOMMTIMEOUTS, OnUpdateOpening1SetCommTimeouts)
   ON_UPDATE_COMMAND_UI (ID_OPENING1_PORTTOOLS_DOITALL, OnUpdateOpening1DoAllOfTheAbove)
   ON_UPDATE_COMMAND_UI (ID_OPENING1_PORTTOOLS_CLOSEHANDLE, OnUpdateOpening1CloseHandle)
   ON_UPDATE_COMMAND_UI (ID_OPENING1_SCALETOOLS_RDWRT_CURRENT, OnUpdateOpening1RdWrtCurrent)
   ON_UPDATE_COMMAND_UI (ID_OPENING1_SCALETOOLS_RDWRT_CONFIGURED, OnUpdateOpening1RdWrtConfigured)
   ON_COMMAND (ID_OPENING1_SHUTDOWN, OnOpening1ShutDown)
   ON_COMMAND (ID_OPENING1_PORTTOOLS_CREATEFILE, OnOpening1CreateFile)
   ON_COMMAND (ID_OPENING1_PORTTOOLS_GETCOMMPROPS, OnOpening1GetCommProps)
   ON_COMMAND (ID_OPENING1_PORTTOOLS_GETCOMMSTATE, OnOpening1GetCommState)
   ON_COMMAND (ID_OPENING1_PORTTOOLS_SETCOMMSTATE, OnOpening1SetCommState)
   ON_COMMAND (ID_OPENING1_PORTTOOLS_GETCOMMTIMEOUTS, OnOpening1GetCommTimeouts)
   ON_COMMAND (ID_OPENING1_PORTTOOLS_SETCOMMTIMEOUTS, OnOpening1SetCommTimeouts)
   ON_COMMAND (ID_OPENING1_PORTTOOLS_DOITALL, OnOpening1DoAllOfTheAbove)
   ON_COMMAND (ID_OPENING1_PORTTOOLS_CLOSEHANDLE, OnOpening1CloseHandle)
   ON_COMMAND (ID_OPENING1_SCALETOOLS_RDWRT_CURRENT, OnOpening1RdWrtCurrent)
   ON_COMMAND (ID_OPENING1_SCALETOOLS_RDWRT_CONFIGURED, OnOpening1RdWrtConfigured)
   ON_LBN_SELCHANGE (IDC_COM_PORT_LB, OnSelchangePortList)
   ON_CBN_CLOSEUP (IDC_BAUD_RATE_CB, OnCloseUpBaudRate)
   ON_CBN_CLOSEUP (IDC_DATA_BITS_CB, OnCloseUpDataBits)
   ON_CBN_CLOSEUP (IDC_STOP_BITS_CB, OnCloseUpStopBits)
   ON_CBN_CLOSEUP (IDC_PARITY_CB, OnCloseUpParity)
   ON_BN_CLICKED (IDC_STRT_THREAD_PB, OnClickedStartThreads)
   ON_BN_CLICKED (IDC_USE_THREAD_PB, OnClickedUseThreads)
   ON_BN_CLICKED (IDC_TIMER_USE_THREAD_PB, OnClickedTimerUseThreads)
   ON_BN_CLICKED (IDC_END_THREAD_PB, OnClickedEndThreads)
   ON_BN_CLICKED (IDC_TEST_PB, OnClickedTest)
   ON_MESSAGE (WM_USER_PORT1_THREAD_FINISHED_DATA, OnPort1ThreadFinishedData)
   ON_MESSAGE (WM_USER_PORT2_THREAD_FINISHED_DATA, OnPort2ThreadFinishedData)
   ON_MESSAGE (WM_USER_PORT3_THREAD_FINISHED_DATA, OnPort3ThreadFinishedData)
   ON_MESSAGE (WM_USER_PORT4_THREAD_FINISHED_DATA, OnPort4ThreadFinishedData)
   ON_MESSAGE (WM_USER_PORT1_THREAD_FINISHED_ERRORS, OnPort1ThreadFinishedErrors)
   ON_MESSAGE (WM_USER_PORT2_THREAD_FINISHED_ERRORS, OnPort2ThreadFinishedErrors)
   ON_MESSAGE (WM_USER_PORT3_THREAD_FINISHED_ERRORS, OnPort3ThreadFinishedErrors)
   ON_MESSAGE (WM_USER_PORT4_THREAD_FINISHED_ERRORS, OnPort4ThreadFinishedErrors)
END_MESSAGE_MAP ()


CScale3MainWin::CScale3MainWin ()
{

   int                     i;

   for (i = 0; i < MAX_NUM_COMPORT; i++)
   {
      SecureZeroMemory (&tp[i], sizeof (THREADPARAMS));
      SecureZeroMemory (&tr[i], sizeof (THREADRETURNS));
      InitializeCriticalSectionAndSpinCount (&(tp[i].CritSec), 0x4);
   }
   m_sThread1Running = m_sThread2Running = m_sThread3Running = m_sThread4Running = 0;

   // Initialize some stuff, MI2 Icon Stuff, and more
   m_MI2Icon1 = (HICON)NULL;
   m_MI2Icon2 = (HICON)NULL;

   m_sCurrentPortIndex = 0;

   // Set them to the known defaults for the MT Scales
   m_sCurrentBaudRateIndex = 6;
   m_sCurrentParityIndex = 2;
   m_sCurrentDataBitsIndex = 2;
   m_sCurrentStopBitsIndex = 0;

   // Set them to the known defaults for the BS2 Simulated Scales
#ifdef BS2PIN16
   m_sCurrentBaudRateIndex = 6;
   m_sCurrentParityIndex = 0;
   m_sCurrentDataBitsIndex = 3;
   m_sCurrentStopBitsIndex = 0;
#endif


   SecureZeroMemory (&m_oCurrentOverlappedStruct, sizeof (OVERLAPPED));
   SecureZeroMemory (&m_CurrentCommTimeouts, sizeof (COMMTIMEOUTS));
   SecureZeroMemory (&m_dcbCurrentDCB, sizeof (DCB));
   m_dcbCurrentDCB.DCBlength = sizeof (DCB);
   // Fill in DCB: 2400 bps, 7 data bits, even parity, and 1 stop bit.
   m_dcbCurrentDCB.BaudRate = CBR_9600;
   m_dcbCurrentDCB.ByteSize = 7;
   m_dcbCurrentDCB.Parity = EVENPARITY;
   m_dcbCurrentDCB.StopBits = ONESTOPBIT;

   m_hCurrentHandle = (HANDLE)NULL;
   m_sHandleFlag = 0;
   m_sGotCommState = 0;
   m_sDCBBuiltFlag = 0;
   m_sGotTimeouts = 0;
   m_sCommStateSet = 0;
   m_sTimeoutsSet = 0;
   m_sPortsConfigured = 0;
   m_sShowConfig = 1;
   m_sShowStruct = 1;

   m_iThreadTimerValue = DEFAULTTHREADTIMERVALUE;
   m_ThreadTimer = 0;

   SecureZeroMemory (&m_CurrentCommProp, sizeof (COMMPROP));
   //m_CurrentCommProp.dwProvSpec1 = COMMPROP_INITIALIZED;

   try
   {
      strMyClass = AfxRegisterWndClass(
         CS_DBLCLKS | CS_HREDRAW | CS_VREDRAW | CS_BYTEALIGNWINDOW | CS_PARENTDC,
         ::LoadCursor(NULL, IDC_ARROW),
         (HBRUSH)::GetStockObject(LTGRAY_BRUSH),
         ::LoadIcon(NULL, IDI_ERROR));
   }
   catch (CResourceException* pEx)
   {
      AfxMessageBox(_T("Couldn't register class! (Already registered?)"));
      pEx->Delete();
   }

   Create (strMyClass, _T("MI2's Mettler Toledo SR16000 Scale Specific Proof of Concept Main Window"));

}  // End of constructor CScale3MainWin


CScale3MainWin::~CScale3MainWin ()
{

   COMSTAT              ComStat;
   DWORD                ComError = 0;
   BOOL                 bResult1;
   int                  i;

   if (m_sThread1Running | m_sThread2Running | m_sThread3Running | m_sThread4Running)
      EndThreads ();

   if (m_hCurrentHandle && !m_sPortsConfigured)
   {
      SecureZeroMemory (&ComStat, sizeof (COMSTAT));
      bResult1 = ClearCommError (
                     m_hCurrentHandle,
                     &ComError,
                     &ComStat);
      CloseHandle (m_hCurrentHandle);
   }

   for (i = 0; i < MAX_NUM_COMPORT; i++)
   {
      if (m_ComPorts[i].hCom)
      {
         SecureZeroMemory (&ComStat, sizeof (COMSTAT));
         bResult1 = ClearCommError (
                        m_ComPorts[i].hCom,
                        &ComError,
                        &ComStat);
         CloseHandle (m_ComPorts[i].hCom);
      }
      DeleteCriticalSection (&(tp[i].CritSec));
   }

}  // End of ~CScale3MainWin


int CScale3MainWin::OnCreate (LPCREATESTRUCT lpcs)
{

   // Error Stuff
   DWORD                   MyError = 0;
   BOOL                    bResult1 = 0;

   TCHAR                   szTemp1[16], szTemp2[STANDARD_STRING];

   int                     iX, iY;
   CRect                   rect;

   HWND                    DTHwnd;
   RECT                    DTRect;

   if (CFrameWnd::OnCreate (lpcs) == -1)
      return -1;

   iX = AreTherSerialPorts();
   if (!iX)
   {
      AfxMessageBox(_T("There are no SerialPorts. Terminating program."));
      PostMessage(WM_CLOSE, 0, 0);
      return 0;
   }
   iY = AreTherStartechSerialPorts();
   if (!iY)
   {
      AfxMessageBox(_T("There are no Startech SerialPorts. Terminating program."));
      PostMessage(WM_CLOSE, 0, 0);
      return 0;
   }
   else
   {
      _stprintf_s(
         szTemp2, 
         STANDARD_STRING,
         _T("%d%s"),
         iY,
         _T(" Startech Serial Ports exist."));
      AfxMessageBox(szTemp2, 0, 0);

   }
   // Take care of the MI2 Icon. I just want to make this guy available for use
   // latter on, as in the Paint routine. This is a dark two tone icon.
   m_MI2Icon1 = LoadIcon (AfxGetInstanceHandle (), MAKEINTRESOURCE (IDI_MI2_1));
   if (m_MI2Icon1 == (HICON)NULL)
      DisplaySystemError (&MyError);

   // Take care of the second MI2 Icon. After this the icon should appear everywhere.
   // This is the Icon I want to appear everywhere, so this one gets used in the
   // SetIcon call. Note that m_MI2Icon1 doesn't get used in this fashion.
   // m_MI2Icon2 is a more colorful three tone icon that is easy to see in the
   // title bar.
   m_MI2Icon2 = LoadIcon (AfxGetInstanceHandle (), MAKEINTRESOURCE (IDI_MI2_2));
   if (m_MI2Icon2 == (HICON)NULL)
      DisplaySystemError (&MyError);
   else
      SetIcon (m_MI2Icon2, TRUE);

   CMenu *pSystemMenu = GetSystemMenu (FALSE);
   pSystemMenu->DeleteMenu (SC_CLOSE, MF_BYCOMMAND);

   // Initialize Sizing Stuff
   GetClientRect(&m_rect);
   m_iClientWidth = m_rect.Width ();
   m_iClientHeight = m_rect.Height ();


   // Load opening menus. Since they're members of the CGenericWindow class
   // their destructors will be called when CGenericWindow goes into destruction.
   // You should see no resource leaks.
   m_menuOpeningMenu1.LoadMenu (IDR_OPENING_MENU1);
   SetMenu (&m_menuOpeningMenu1);


   m_fontMain.CreatePointFont (80, _T("MS Sans Serif"));

   CClientDC dc (this);
   CFont *pOldFont = dc.SelectObject (&m_fontMain);
   TEXTMETRIC tm;
   dc.GetTextMetrics (&tm);
   m_cxChar = tm.tmAveCharWidth;
   m_cyChar = tm.tmHeight + tm.tmExternalLeading;
   dc.SelectObject (pOldFont);

   // Com-Port listbox
   // m_ComPortLBTitle           IDC_COM_PORT_LB_TITLE
   // m_ComPortLB                IDC_COM_PORT_LB
   rect.SetRect (m_cxChar * 2, m_cyChar, m_cxChar * 40, m_cyChar * 2);
   m_ComPortLBTitle.Create (
      _T("Available Com Ports"),
      WS_CHILD | WS_VISIBLE | SS_LEFT,
      rect, this, IDC_COM_PORT_LB_TITLE);

   rect.SetRect (m_cxChar * 2, m_cyChar * 2, m_cxChar * 40, m_cyChar * 18);
   m_ComPortLB.Create (
      WS_CHILD | WS_HSCROLL | WS_VISIBLE | LBS_NOTIFY,
      rect,
      this,
      IDC_COM_PORT_LB);

   m_ComPortLBTitle.SetFont (&m_fontMain, false);
   m_ComPortLB.SetFont (&m_fontMain, false);

   FillComPortLB ();

   // Baud Rate combobox
   // m_BaudRateCBTitle          IDC_BAUD_RATE_CB_TITLE
   // m_BaudRateCB               IDC_BAUD_RATE_CB
   rect.SetRect (m_cxChar * 45, m_cyChar, m_cxChar * 60, m_cyChar * 2);
   bResult1 = m_BaudRateCBTitle.Create (
      _T("BAUD Rates"),
      WS_CHILD | WS_VISIBLE | SS_LEFT,
      rect, this, IDC_BAUD_RATE_CB_TITLE);

   rect.SetRect (m_cxChar * 45, m_cyChar * 2, m_cxChar * 60, m_cyChar * 9);
   bResult1 = m_BaudRateCB.Create (
      WS_CHILD | WS_VISIBLE | WS_VSCROLL | CBS_DROPDOWNLIST,
      rect,
      this,
      IDC_BAUD_RATE_CB);

   m_BaudRateCBTitle.SetFont (&m_fontMain, false);
   m_BaudRateCB.SetFont (&m_fontMain, false);

   // Parity combobox
   // m_ParityCBTitle            IDC_PARITY_CB_TITLE
   // m_ParityCB                 IDC_PARITY_CB
   rect.SetRect (m_cxChar * 60, m_cyChar, m_cxChar * 80, m_cyChar * 2);
   bResult1 = m_ParityCBTitle.Create (
      _T("Parity"),
      WS_CHILD | WS_VISIBLE | SS_LEFT,
      rect, this, IDC_PARITY_CB_TITLE);

   rect.SetRect (m_cxChar * 60, m_cyChar * 2, m_cxChar * 80, m_cyChar * 9);
   bResult1 = m_ParityCB.Create (
      WS_CHILD | WS_VISIBLE | WS_VSCROLL | CBS_DROPDOWNLIST,
      rect,
      this,
      IDC_PARITY_CB);

   m_ParityCBTitle.SetFont (&m_fontMain, false);
   m_ParityCB.SetFont (&m_fontMain, false);

   // Data Bits combobox
   // m_DataBitsCBTitle          IDC_DATA_BITS_CB_TITLE
   // m_DataBitsCB               IDC_DATA_BITS_CB
   rect.SetRect (m_cxChar * 80, m_cyChar, m_cxChar * 90, m_cyChar * 2);
   bResult1 = m_DataBitsCBTitle.Create (
      _T("Data Bits"),
      WS_CHILD | WS_VISIBLE | SS_LEFT,
      rect, this, IDC_DATA_BITS_CB_TITLE);

   rect.SetRect (m_cxChar * 80, m_cyChar * 2, m_cxChar * 90, m_cyChar * 9);
   bResult1 = m_DataBitsCB.Create (
      WS_CHILD | WS_VISIBLE | WS_VSCROLL | CBS_DROPDOWNLIST,
      rect,
      this,
      IDC_DATA_BITS_CB);

   m_DataBitsCBTitle.SetFont (&m_fontMain, false);
   m_DataBitsCB.SetFont (&m_fontMain, false);

   // Stop Bits combobox
   // m_StopBitsCBTitle          IDC_STOP_BITS_CB_TITLE
   // m_StopBitsCB               IDC_STOP_BITS_CB
   rect.SetRect (m_cxChar * 90, m_cyChar, m_cxChar * 100, m_cyChar * 2);
   bResult1 = m_StopBitsCBTitle.Create (
      _T("Stop Bits"),
      WS_CHILD | WS_VISIBLE | SS_LEFT,
      rect, this, IDC_STOP_BITS_CB_TITLE);

   rect.SetRect (m_cxChar * 90, m_cyChar * 2, m_cxChar * 100, m_cyChar * 9);
   bResult1 = m_StopBitsCB.Create (
      WS_CHILD | WS_VISIBLE | WS_VSCROLL | CBS_DROPDOWNLIST,
      rect,
      this,
      IDC_STOP_BITS_CB);

   m_StopBitsCBTitle.SetFont (&m_fontMain, false);
   m_StopBitsCB.SetFont (&m_fontMain, false);

   SetComboBoxes ();

   // Status Text Fields
   // m_StatusText1              IDC_STATUS_TEXT_1
   rect.SetRect (m_cxChar * 2, m_cyChar * 26, m_cxChar * 123, m_cyChar * 27);
   bResult1 = m_StatusText1.Create (
      NULL,
      WS_CHILD | WS_VISIBLE | SS_LEFT,
      rect, this, IDC_STATUS_TEXT_1);

   rect.SetRect (m_cxChar * 2, m_cyChar * 27, m_cxChar * 123, m_cyChar * 28);
   bResult1 = m_StatusText2.Create (
      NULL,
      WS_CHILD | WS_VISIBLE | SS_LEFT,
      rect, this, IDC_STATUS_TEXT_2);

   m_StatusText1.SetFont (&m_fontMain, false);
   m_StatusText2.SetFont (&m_fontMain, false);

   // Structure Members Display
   // m_StructLBTitle          IDC_STRUCT_LB_TITLE
   // m_StructLB               IDC_STRUCT_LB
   rect.SetRect (m_cxChar * 45, m_cyChar * 5, m_cxChar * 90, m_cyChar * 6);
   bResult1 = m_StructLBTitle.Create (
      _T(""),
      WS_CHILD | WS_VISIBLE | SS_LEFT,
      rect, this, IDC_STRUCT_LB_TITLE);

   rect.SetRect (m_cxChar * 45, m_cyChar * 6, m_cxChar * 90, m_cyChar * 19);
   bResult1 = m_StructLB.Create (
      WS_CHILD | WS_VISIBLE | WS_HSCROLL | WS_VSCROLL | LBS_NOSEL,
      rect,
      this,
      IDC_STRUCT_LB);

   m_StructLBTitle.SetFont (&m_fontMain, false);
   m_StructLBTitle.ShowWindow (SW_HIDE);
   m_StructLB.SetFont (&m_fontMain, false);
   m_StructLB.ShowWindow (SW_HIDE);

   // Configured Ports Data Display
   // m_ConfiguredLBTitle          IDC_CONFIG_LB_TITLE
   // m_ConfiguredLB               IDC_CONFIG_LB
   rect.SetRect (m_cxChar * 2, m_cyChar * 20, m_cxChar * 100, m_cyChar * 21);
   bResult1 = m_ConfiguredLBTitle.Create (
      _T("Data received from configured ports"),
      WS_CHILD | WS_VISIBLE | SS_LEFT,
      rect, this, IDC_CONFIG_LB_TITLE);

   rect.SetRect (m_cxChar * 2, m_cyChar * 21, m_cxChar * 100, m_cyChar * 26);
   bResult1 = m_ConfiguredLB.Create (
      WS_CHILD | WS_VISIBLE | WS_HSCROLL | WS_VSCROLL | LBS_NOSEL,
      rect,
      this,
      IDC_CONFIG_LB);

   m_ConfiguredLBTitle.SetFont (&m_fontMain, false);
   m_ConfiguredLBTitle.ShowWindow (SW_HIDE);
   m_ConfiguredLB.SetFont (&m_fontMain, false);
   m_ConfiguredLB.ShowWindow (SW_HIDE);


   // m_TestPB                   IDC_TEST_PB
   //rect.SetRect (m_cxChar * 105, m_cyChar * 2, m_cxChar * 120, m_cyChar * 5);
   //bResult1 = m_TestPB.Create (
      //"Test",
      //WS_CHILD | WS_VISIBLE | BS_MULTILINE | BS_CENTER | BS_PUSHBUTTON,
      //rect,
      //this,
      //IDC_TEST_PB);


   // Configured Ports Threads Start, Use, Use w/ Timers, and End
   // m_StartThreadsPB           IDC_STRT_THREAD_PB
   // m_UseThreadsPB             IDC_USE_THREAD_PB
   // m_TimerUseThreadsPB        IDC_TIMER_USE_THREAD_PB
   // m_EndThreadsPB             IDC_END_THREAD_PB

   rect.SetRect (m_cxChar * 105, m_cyChar * 2, m_cxChar * 120, m_cyChar * 5);
   bResult1 = m_StartThreadsPB.Create (
      _T("Start Threads"),
      WS_CHILD | WS_VISIBLE | BS_MULTILINE | BS_CENTER | BS_PUSHBUTTON,
      rect,
      this,
      IDC_STRT_THREAD_PB);
   rect.SetRect (m_cxChar * 105, m_cyChar * 6, m_cxChar * 120, m_cyChar * 9);
   bResult1 = m_UseThreadsPB.Create (
      _T("Use Threads"),
      WS_CHILD | WS_VISIBLE | BS_MULTILINE | BS_CENTER | BS_PUSHBUTTON,
      rect,
      this,
      IDC_USE_THREAD_PB);
   rect.SetRect (m_cxChar * 105, m_cyChar * 10, m_cxChar * 120, m_cyChar * 13);
   bResult1 = m_TimerUseThreadsPB.Create (
      _T("Use Threads w/ Timers"),
      WS_CHILD | WS_VISIBLE | BS_MULTILINE | BS_CENTER | BS_PUSHBUTTON,
      rect,
      this,
      IDC_TIMER_USE_THREAD_PB);
   rect.SetRect (m_cxChar * 105, m_cyChar * 14, m_cxChar * 120, m_cyChar * 17);
   bResult1 = m_EndThreadsPB.Create (
      _T("End Threads"),
      WS_CHILD | WS_VISIBLE | BS_MULTILINE | BS_CENTER | BS_PUSHBUTTON,
      rect,
      this,
      IDC_END_THREAD_PB);

   m_StartThreadsPB.SetFont (&m_fontMain, false);
   m_StartThreadsPB.ShowWindow (SW_HIDE);
   m_UseThreadsPB.SetFont (&m_fontMain, false);
   m_UseThreadsPB.ShowWindow (SW_HIDE);
   m_TimerUseThreadsPB.SetFont (&m_fontMain, false);
   m_TimerUseThreadsPB.ShowWindow (SW_HIDE);
   m_EndThreadsPB.SetFont (&m_fontMain, false);
   m_EndThreadsPB.ShowWindow (SW_HIDE);

   // m_ThreadTimerEBTitle       IDC_THREAD_TIMER_EB_TITLE
   // m_ThreadTimerEB            IDC_THREAD_TIMER_EB
   rect.SetRect (m_cxChar * 80, m_cyChar * 10, m_cxChar * 100, m_cyChar * 11);
   bResult1 = m_ThreadTimerEBTitle.Create (
      _T("Enter Timer in mS"),
      WS_CHILD | WS_VISIBLE | SS_LEFT,
      rect,
      this,
      IDC_THREAD_TIMER_EB_TITLE);
   rect.SetRect (m_cxChar * 80, m_cyChar * 11, m_cxChar * 100, m_cyChar * 12);
   bResult1 = m_ThreadTimerEB.Create (
      WS_CHILD | WS_VISIBLE | ES_LEFT | ES_NUMBER,
      rect,
      this,
      IDC_THREAD_TIMER_EB);

   m_ThreadTimerEBTitle.SetFont (&m_fontMain, false);
   m_ThreadTimerEBTitle.ShowWindow (SW_HIDE);
   m_ThreadTimerEB.SetFont (&m_fontMain, false);
   m_ThreadTimerEB.ShowWindow (SW_HIDE);
   m_ThreadTimerEB.LimitText (5);
   _stprintf_s(szTemp1, 16, _T("%d"), DEFAULTTHREADTIMERVALUE);
   m_ThreadTimerEB.SetWindowText (szTemp1);

   m_StatusText1.SetWindowText(_T("Current Device Path :"));
   m_StatusText2.SetWindowText (m_ComPorts[m_sCurrentPortIndex].szDevicePath);

   DTHwnd = ::GetDesktopWindow ();
   ::GetWindowRect (DTHwnd, &DTRect);
   iX = (int)(DTRect.right * 0.25);
   iY = (int)(DTRect.bottom * 0.6);
   //SetWindowPos (
      //&wndTop,
      //iX,
      //0,
      //2 * iX,
      //iY,
      //SWP_SHOWWINDOW | SWP_NOZORDER);

   CRect MyRect (0, 0, m_cxChar * 125, m_cyChar * 30);
   CalcWindowRect (&MyRect);
   SetWindowPos (&wndTop, iX, 0, MyRect.Width (), MyRect.Height (), SWP_SHOWWINDOW | SWP_NOZORDER);

   return 1;

}  // End of OnCreate

void CScale3MainWin::OnClose (void)
{

   CFrameWnd::OnClose ();

}  // End of OnClose


///////////////////// MFC/WIN-32 Menu-Update Messages ///////////////////////////

void CScale3MainWin::OnUpdateOpening1ShutDown (CCmdUI *pCmdUI)
{

   bool bTemp1;

   bTemp1 = true;

   pCmdUI->Enable (bTemp1);

}  // OnUpdateOpening1ShutDown

void CScale3MainWin::OnUpdateOpening1CreateFile (CCmdUI *pCmdUI)
{

   bool bTemp1;

   if (!m_sHandleFlag)
      bTemp1 = true;
   else
      bTemp1 = false;

   pCmdUI->Enable (bTemp1);

}  // OnUpdateOpening1CreateFile

void CScale3MainWin::OnUpdateOpening1GetCommProps (CCmdUI *pCmdUI)
{

   bool bTemp1;

   if (!m_sHandleFlag)
      bTemp1 = false;
   else
      bTemp1 = true;

   pCmdUI->Enable (bTemp1);

}  // OnUpdateOpening1GetCommProps

void CScale3MainWin::OnUpdateOpening1GetCommState (CCmdUI *pCmdUI)
{

   bool bTemp1;

   if (!m_sHandleFlag)
      bTemp1 = false;
   else
      bTemp1 = true;

   pCmdUI->Enable (bTemp1);

}  // OnUpdateOpening1GetCommState

void CScale3MainWin::OnUpdateOpening1SetCommState (CCmdUI *pCmdUI)
{

   bool bTemp1;

   if (!m_sHandleFlag || !m_sGotCommState)
      bTemp1 = false;
   else
      bTemp1 = true;

   pCmdUI->Enable (bTemp1);

}  // OnUpdateOpening1SetCommState

void CScale3MainWin::OnUpdateOpening1GetCommTimeouts (CCmdUI *pCmdUI)
{

   bool bTemp1;

   if (!m_sHandleFlag || !m_sDCBBuiltFlag)
      bTemp1 = false;
   else
      bTemp1 = true;

   pCmdUI->Enable (bTemp1);

}  // OnUpdateOpening1GetCommTimeouts

void CScale3MainWin::OnUpdateOpening1SetCommTimeouts (CCmdUI *pCmdUI)
{

   bool bTemp1;

   if (!m_sHandleFlag || !m_sGotTimeouts)
      bTemp1 = false;
   else
      bTemp1 = true;

   pCmdUI->Enable (bTemp1);

}  // OnUpdateOpening1SetCommTimeouts

void CScale3MainWin::OnUpdateOpening1DoAllOfTheAbove (CCmdUI *pCmdUI)
{

   bool bTemp1;

   if (!m_sHandleFlag)
      bTemp1 = true;
   else
      bTemp1 = false;

   pCmdUI->Enable (bTemp1);

}  // OnUpdateOpening1DoAllOfTheAbove

void CScale3MainWin::OnUpdateOpening1CloseHandle (CCmdUI *pCmdUI)
{

   bool bTemp1;

   if (!m_sHandleFlag)
      bTemp1 = false;
   else
      bTemp1 = true;

   pCmdUI->Enable (bTemp1);

}  // OnUpdateOpening1CloseHandle

void CScale3MainWin::OnUpdateOpening1RdWrtCurrent (CCmdUI *pCmdUI)
{

   bool bTemp1;

   if (!m_sHandleFlag || !m_sCommStateSet || !m_sTimeoutsSet)
      bTemp1 = false;
   else
      bTemp1 = true;

   pCmdUI->Enable (bTemp1);

}  // OnUpdateOpening1RdWrtCurrent


void CScale3MainWin::OnUpdateOpening1RdWrtConfigured (CCmdUI *pCmdUI)
{

   bool bTemp1;

   if (!m_sPortsConfigured)
      bTemp1 = false;
   else
      bTemp1 = true;

   pCmdUI->Enable (bTemp1);

}  // OnUpdateOpening1RdWrtConfigured 


/////////////////////////// User Messages //////////////////////////


long CScale3MainWin::OnPort1ThreadFinishedData (WPARAM wParam, LPARAM lParam)
{

   TCHAR                      szTemp1[STANDARD_STRING];

   EnterCriticalSection (&(tp[0].CritSec));
   _stprintf_s(szTemp1, STANDARD_STRING, _T("%s - %S"), m_ComPorts[0].szName, (char *)tr[0].Data);
   LeaveCriticalSection(&(tp[0].CritSec));
   m_ConfiguredLB.AddString (szTemp1);
   return (0);

}  // End of OnPort1ThreadFinished


long CScale3MainWin::OnPort2ThreadFinishedData (WPARAM wParam, LPARAM lParam)
{

   TCHAR                      szTemp1[STANDARD_STRING];

   EnterCriticalSection (&(tp[1].CritSec));
   _stprintf_s(szTemp1, STANDARD_STRING, _T("%s - %S"), m_ComPorts[1].szName, (char *)tr[1].Data);
   LeaveCriticalSection(&(tp[1].CritSec));
   m_ConfiguredLB.AddString (szTemp1);
   return (0);

}  // End of OnPort1ThreadFinished


long CScale3MainWin::OnPort3ThreadFinishedData (WPARAM wParam, LPARAM lParam)
{

   TCHAR                       szTemp1[STANDARD_STRING];

   EnterCriticalSection (&(tp[2].CritSec));
   _stprintf_s(szTemp1, STANDARD_STRING, _T("%s - %S"), m_ComPorts[2].szName, (char *)tr[2].Data);
   LeaveCriticalSection (&(tp[2].CritSec));
   m_ConfiguredLB.AddString (szTemp1);
   return (0);

}  // End of OnPort1ThreadFinished


long CScale3MainWin::OnPort4ThreadFinishedData (WPARAM wParam, LPARAM lParam)
{

   TCHAR                      szTemp1[STANDARD_STRING];

   EnterCriticalSection (&(tp[3].CritSec));
   _stprintf_s(szTemp1, STANDARD_STRING, _T("%s - %S"), m_ComPorts[3].szName, (char *)tr[3].Data);
   LeaveCriticalSection (&(tp[3].CritSec));
   m_ConfiguredLB.AddString (szTemp1);
   return (0);

}  // End of OnPort1ThreadFinished


long CScale3MainWin::OnPort1ThreadFinishedErrors (WPARAM wParam, LPARAM lParam)
{
   EnterCriticalSection (&(tp[0].CritSec));
   m_StatusText1.SetWindowText (tr[0].ErrorMsg1);
   m_StatusText2.SetWindowText (tr[0].ErrorMsg2);
   LeaveCriticalSection (&(tp[0].CritSec));
   return (0);
}  // End of OnPort1ThreadFinished


long CScale3MainWin::OnPort2ThreadFinishedErrors (WPARAM wParam, LPARAM lParam)
{
   EnterCriticalSection (&(tp[1].CritSec));
   m_StatusText1.SetWindowText (tr[1].ErrorMsg1);
   m_StatusText2.SetWindowText (tr[1].ErrorMsg2);
   LeaveCriticalSection (&(tp[1].CritSec));
   return (0);
}  // End of OnPort2ThreadFinished


long CScale3MainWin::OnPort3ThreadFinishedErrors (WPARAM wParam, LPARAM lParam)
{
   EnterCriticalSection (&(tp[2].CritSec));
   m_StatusText1.SetWindowText (tr[2].ErrorMsg1);
   m_StatusText2.SetWindowText (tr[2].ErrorMsg2);
   LeaveCriticalSection (&(tp[2].CritSec));
   return (0);
}  // End of OnPort3ThreadFinished


long CScale3MainWin::OnPort4ThreadFinishedErrors (WPARAM wParam, LPARAM lParam)
{
   EnterCriticalSection (&(tp[3].CritSec));
   m_StatusText1.SetWindowText (tr[3].ErrorMsg1);
   m_StatusText2.SetWindowText (tr[3].ErrorMsg2);
   LeaveCriticalSection (&(tp[3].CritSec));
   return (0);
}  // End of OnPort4ThreadFinished


///////////////////// Menu Commands ///////////////////////

void CScale3MainWin::OnOpening1ShutDown (void)
{

   PostMessage (WM_CLOSE, 0, 0);

}  // End of OnOpening1ShutDown

/************************************************************************************************
June 12th, 2015

Specifying FILE_FLAG_OVERLAPPED means you will be performing asychronous I/O. Using a value of zero
instead means you will be doing synchronous I/O. Synchronous I/O is serialized I/O. In other words,
when you make any sort of I/O call to the handle returned by CreateFile, the thread making that call
stops at that I/O call and waits for the I/O call to return either sucessfully or unsucessfully.

Asyncronous I/O means that when you make any sort of I/O call to the handle returned by CreateFile,
that I/O call returns immediately with the message that Overlapped I/O is in progress. You would
mistakenly think this is an error condition but it is not. Asynchronous I/O makes use of the OVERLAPPED
structure. You create this structure and include it as a parameter in calls to I/O Functions.
It is up to you to use the Event you created in the OVERLAPPED structure. The kernel issues this event
when the I/O function that includes the OVERLAPPED structure as a parameter has
completed. This means that instead of your thread stopping at each I/O call and waiting
for the I/O call to return, the thread can proceed with other tasks.

You should use synchronous I/O when you expect the I/O call to have a speedy response, say if you're
trying to communicate with a device connected to the computer via a serial port. Use asynchronous I/O
when you know that a response to an I/O call could take up too large a chunk of processor time, say to
a device that performs backup duties.

While investigating the port from XP to Win8.1, I found that the use of asynchronous I/O under Win8.1
is much more kernel intensive. Under XP I didn't have to check for the completion of overlapped I/O.
I/O calls completed exactly the same if I specified asynchronous or synchronous I/O in the CreateFile.
Not so under Win8.1. I have to use the WinAPI GetOverlappedResult after each I/O call that includes the
OVERLAPPED structure as a parameter when I specify asynchronous I/O.
I have not investigated making use of the event I created in the OVERLAPPED structure
and that the kernel is supposed to issue when the I/O is done. Since I'm trying to communicate with a
piece of lab equipment, I'm going to assume that its I/O response will be quick enough not to represent
a burnousume time to the process. Thus I'll specify synchronous I/O in my call to CreateFile.

It is interesting to note that asynchronous I/O such as read and writes infact seem to read and write
at the time of the call. However, they return zero bytes read or written. It is only the subsequent call
to GetOverlappedResult that returns the number of bytes read or written.

It is also interesting to note that the synchronous I/O calls can be made with or without an OVERLAPPED
structure as a parameter. Also I have run sucessfully when using synchronous I/O and making the call to
GetOverlappedResult. I have not investigated what happens on a synchronous I/O call that doesn't include
an OVERLAPPED structure followed by a call to GetOverLappedResult. I do know it seems pointless to make an
API call to GetOverlappedResult if I'm using synchronous I/O so I'm using the precompilier via the #define
ASYNC variable. Define ASYNC if you wish to compile an EXE that uses asynchronous I/O.

************************************************************************************************/
void CScale3MainWin::OnOpening1CreateFile (void)
{

   DWORD                MyError = 0;
   TCHAR                szTemp1[STANDARD_STRING];

#ifdef ASYNC
   m_hCurrentHandle = CreateFile (
                           m_ComPorts[m_sCurrentPortIndex].szDevicePath,
                           GENERIC_READ | GENERIC_WRITE,
                           0,       // must be opened with exclusive-access
                           NULL,    // default security attributes
                           OPEN_EXISTING, // must use OPEN_EXISTING
                           FILE_FLAG_OVERLAPPED,  // Asynchronous I/O
                           NULL);   // hTemplate must be NULL for comm devices
#else
   m_hCurrentHandle = CreateFile(
                           m_ComPorts[m_sCurrentPortIndex].szDevicePath,
                           GENERIC_READ | GENERIC_WRITE,
                           0,       // must be opened with exclusive-access
                           NULL,    // default security attributes
                           OPEN_EXISTING, // must use OPEN_EXISTING
                           0,       // Synchronous I/O
                           NULL);   // hTemplate must be NULL for comm devices
#endif

   if (m_hCurrentHandle == INVALID_HANDLE_VALUE) 
   {
      DisplaySystemError (&MyError);
      _stprintf_s(szTemp1, STANDARD_STRING, _T("%s%d%s"),
         _T("CreateFile failed with error number "), MyError,
         _T(", the returned File Handle = INVALID_HANDLE_VALUE"));
      m_StatusText1.SetWindowText (szTemp1);
      m_hCurrentHandle = (HANDLE)NULL;
      m_sHandleFlag = 0;
      return;
   }

   _stprintf_s(szTemp1, STANDARD_STRING, _T("%s"), _T("CreateFile succeeded."));
   m_StatusText1.SetWindowText (szTemp1);
   m_StatusText2.SetWindowText(_T(""));
   m_sHandleFlag = 1;

}  // End of OnOpening1CreateFile


// See notes at end of this file on the COMMPROP structure.
void CScale3MainWin::OnOpening1GetCommProps (void)
{

   DWORD                MyError = 0;
   BOOL                 bResult1;
   TCHAR                szTemp1[STANDARD_STRING];

   bResult1 = GetCommProperties (m_hCurrentHandle, &m_CurrentCommProp);

   if (!bResult1)
   {
      DisplaySystemError (&MyError);
      _stprintf_s(szTemp1, STANDARD_STRING, _T("%s%d"), _T("GetCommProperties failed with error number"), MyError);
      m_StatusText1.SetWindowText (szTemp1);
   }
   else
   {
      _stprintf_s(szTemp1, STANDARD_STRING, _T("%s"), _T("GetCommProperties succeeded."));
      m_StatusText1.SetWindowText (szTemp1);
      m_StatusText2.SetWindowText(_T(""));
      OutputProperties ();
   }

}  // End of OnOpening1GetCommProps


void CScale3MainWin::OnOpening1GetCommState (void)
{

   DWORD                MyError = 0;
   BOOL                 bResult1;
   TCHAR                szTemp1[STANDARD_STRING];

   if (!m_sDCBBuiltFlag)
   {
      SecureZeroMemory (&m_dcbCurrentDCB, sizeof (DCB));
      m_dcbCurrentDCB.DCBlength = sizeof (DCB);
   }
   bResult1 = GetCommState (m_hCurrentHandle, &m_dcbCurrentDCB);

   if (!bResult1)
   {
      DisplaySystemError (&MyError);
      _stprintf_s(szTemp1, STANDARD_STRING, _T("%s%d"), _T("GetCommState failed with error number "), MyError);
      m_StatusText1.SetWindowText (szTemp1);
   }
   else
   {
      _stprintf_s(szTemp1, STANDARD_STRING, _T("%s"), _T("GetCommState succeeded."));
      m_StatusText1.SetWindowText (szTemp1);
      m_StatusText2.SetWindowText(_T(""));
      m_sGotCommState = 1;
      OutputCommState ();
   }

}  // End of OnOpening1GetCommState


void CScale3MainWin::OnOpening1SetCommState (void)
{

   // First do BuildCommDCB, if it fails don't do anything else and return
   // Second do SetCommState, if it fails don't do anything else and return
   // Third do SetCommMask, if it fails don't do anything else and return
   // If SetCommMask succeedes, the important flag, m_sCommStateSet, is set
   //   and the member OVERLAPPED struct has its event created and is initialized.

   DWORD                MyError = 0;
   int                  iTemp1;
   BOOL                 bResult1;
   TCHAR                szTemp1[STANDARD_STRING];

   //SecureZeroMemory (&m_dcbCurrentDCB, sizeof (DCB));
   //m_dcbCurrentDCB.DCBlength = sizeof (DCB);

   iTemp1 = (int)fStopBits[m_sCurrentStopBitsIndex];
   //_stprintf_s (szTemp1, STANDARD_STRING, _T("baud=%d parity=%s data=%d stop=%d"),
      //ulBaudRates[m_sCurrentBaudRateIndex],
      //szParity[m_sCurrentParityIndex],
      //sDataBits[m_sCurrentDataBitsIndex],
      //iTemp1);
   _stprintf_s(szTemp1, STANDARD_STRING, _T("%d,%s,%d,%d"),
      ulBaudRates[m_sCurrentBaudRateIndex],
      (LPCTSTR)szParity2[m_sCurrentParityIndex],
      sDataBits[m_sCurrentDataBitsIndex],
      iTemp1);
   bResult1 = BuildCommDCB (szTemp1, &m_dcbCurrentDCB);
   //bResult1 = BuildCommDCB ("2400,e,7,1", &m_dcbCurrentDCB);
   //m_dcbCurrentDCB.BaudRate = ulBaudRates[m_sCurrentBaudRateIndex];
   //m_dcbCurrentDCB.ByteSize = (BYTE)sDataBits[m_sCurrentDataBitsIndex];
   //m_dcbCurrentDCB.Parity = (BYTE)m_sCurrentParityIndex;
   //m_dcbCurrentDCB.StopBits = (BYTE)m_sCurrentStopBitsIndex;
   //bResult1 = true;


   if (!bResult1)
   {
      DisplaySystemError (&MyError);
      _stprintf_s(szTemp1, STANDARD_STRING, _T("%s%d"), _T("BuildCommDCB failed with error number "), MyError);
      m_StatusText1.SetWindowText (szTemp1);
      m_sDCBBuiltFlag = 0;
   }  // End of BuildCommDCB failed
   else
   {
      _stprintf_s(szTemp1, STANDARD_STRING, _T("%s"), _T("BuildCommDCB succeeded."));
      m_StatusText1.SetWindowText (szTemp1);
      m_StatusText2.SetWindowText(_T(""));
      m_sDCBBuiltFlag = 1;
      m_dcbCurrentDCB.EvtChar = '\n';

      bResult1 = SetCommState (m_hCurrentHandle, &m_dcbCurrentDCB);

      if (!bResult1)
      {
         DisplaySystemError (&MyError);
         _stprintf_s(szTemp1, STANDARD_STRING, _T("%s%d"),
                     _T("BuildCommDCB succeeded but SetCommState failed with error number "), MyError);
         m_StatusText1.SetWindowText (szTemp1);
         return;
      }  // End of SetCommState failure
      else
      {
         bResult1 = SetCommMask (m_hCurrentHandle, EV_RXFLAG);// | EV_ERR | EV_BREAK);
         if (!bResult1)
         {
            DisplaySystemError (&MyError);
            _stprintf_s(szTemp1, STANDARD_STRING, _T("%s%d"),
                        _T("BuildCommDCB & SetCommState succeeded but SetCommMask failed with error number "), MyError);
            m_StatusText1.SetWindowText (szTemp1);
            return;
         }  // End of SetCommMask failed

         m_oCurrentOverlappedStruct.hEvent = CreateEvent (
                                                NULL,
                                                true,
                                                false,
                                                NULL);
         m_oCurrentOverlappedStruct.Internal = 0;  // STATUS_PENDING after an I/O. Set back to zero on completion?
         m_oCurrentOverlappedStruct.InternalHigh = 0;
         m_oCurrentOverlappedStruct.Offset = 0;
         m_oCurrentOverlappedStruct.OffsetHigh = 0;

         assert (m_oCurrentOverlappedStruct.hEvent);

         _stprintf_s(szTemp1, STANDARD_STRING, _T("%s"),
                     _T("BuildCommDCB, SetCommState, SetCommMask, and the OVERLAPPED struct intialization succeeded."));
         m_StatusText1.SetWindowText (szTemp1);
         m_StatusText2.SetWindowText(_T(""));
         m_sCommStateSet = 1;

      }  // End of SetCommState success

   }  // End of BuildCommDCB success

}  // End of OnOpening1SetCommState


void CScale3MainWin::OnOpening1GetCommTimeouts (void)
{

   DWORD                MyError = 0;
   BOOL                 bResult1;
   TCHAR                szTemp1[STANDARD_STRING];

   bResult1 = GetCommTimeouts (m_hCurrentHandle, &m_CurrentCommTimeouts);

   if (!bResult1)
   {
      DisplaySystemError (&MyError);
      _stprintf_s(szTemp1, STANDARD_STRING, _T("%s%d"), _T("GetCommTimeouts failed with error number "), MyError);
      m_StatusText1.SetWindowText (szTemp1);
   }
   else
   {
      _stprintf_s(szTemp1, STANDARD_STRING, _T("%s"), _T("GetCommTimeouts succeeded."));
      m_StatusText1.SetWindowText (szTemp1);
      m_StatusText2.SetWindowText(_T(""));
      m_sGotTimeouts = 1;
      OutputCommTimeouts ();
   }

}  // End of OnOpening1GetCommTimeouts


void CScale3MainWin::OnOpening1SetCommTimeouts (void)
{

   DWORD                MyError = 0;
   BOOL                 bResult1;
   TCHAR                szTemp1[STANDARD_STRING];

   m_CurrentCommTimeouts.ReadIntervalTimeout = MAXDWORD;
   m_CurrentCommTimeouts.ReadTotalTimeoutMultiplier = 0;
   m_CurrentCommTimeouts.ReadTotalTimeoutConstant = 0;
   m_CurrentCommTimeouts.WriteTotalTimeoutMultiplier = 10;
   m_CurrentCommTimeouts.WriteTotalTimeoutConstant = 1000;
   bResult1 = SetCommTimeouts (m_hCurrentHandle, &m_CurrentCommTimeouts);

   if (!bResult1)
   {
      DisplaySystemError (&MyError);
      _stprintf_s(szTemp1, STANDARD_STRING, _T("%s%d"), _T("SetCommTimeouts failed with error number "), MyError);
      m_StatusText1.SetWindowText (szTemp1);
   }
   else
   {
      _stprintf_s(szTemp1, STANDARD_STRING, _T("%s"), _T("SetCommTimeouts succeeded."));
      m_StatusText1.SetWindowText (szTemp1);
      m_StatusText2.SetWindowText(_T(""));
      m_sTimeoutsSet = 1;
   }

}  // End of OnOpening1SetCommTimeouts


void CScale3MainWin::OnOpening1DoAllOfTheAbove (void)
{
   COMSTAT              ComStat;
   DWORD                ComError = 0;
   int                  iResult1 = 0;
   TCHAR                szTemp1[STANDARD_STRING];

   iResult1 = DoAllOfTheAboveStep1 ();    // CreateFile
   if (!iResult1)
      return;

   iResult1 = DoAllOfTheAboveStep2 ();    // BuildCommDCB
   if (!iResult1)
   {
      SecureZeroMemory (&ComStat, sizeof (COMSTAT));
      ClearCommError (
         m_hCurrentHandle,
         &ComError,
         &ComStat);
      m_hCurrentHandle = (HANDLE)NULL;
      m_sHandleFlag = 0;
      return;
   }

   iResult1 = DoAllOfTheAboveStep3 ();    // SetCommState
   if (!iResult1)
   {
      m_sDCBBuiltFlag = 0;
      SecureZeroMemory (&ComStat, sizeof (COMSTAT));
      ClearCommError (
         m_hCurrentHandle,
         &ComError,
         &ComStat);
      m_hCurrentHandle = (HANDLE)NULL;
      m_sHandleFlag = 0;
      return;
   }

   iResult1 = DoAllOfTheAboveStep4 ();    // SetCommMask
   if (!iResult1)
   {
      m_sDCBBuiltFlag = 0;
      SecureZeroMemory (&ComStat, sizeof (COMSTAT));
      ClearCommError (
         m_hCurrentHandle,
         &ComError,
         &ComStat);
      m_hCurrentHandle = (HANDLE)NULL;
      m_sHandleFlag = 0;
      return;
   }

   iResult1 = DoAllOfTheAboveStep5 ();    // Build OVERLAP structure
   if (!iResult1)
   {
      m_sDCBBuiltFlag = 0;
      SecureZeroMemory (&ComStat, sizeof (COMSTAT));
      ClearCommError (
         m_hCurrentHandle,
         &ComError,
         &ComStat);
      m_hCurrentHandle = (HANDLE)NULL;
      m_sHandleFlag = 0;
      return;
   }

   iResult1 = DoAllOfTheAboveStep6 ();    // SetCommTimeouts
   if (!iResult1)
   {
      m_sCommStateSet = 0;
      m_sDCBBuiltFlag = 0;
      SecureZeroMemory (&ComStat, sizeof (COMSTAT));
      ClearCommError (
         m_hCurrentHandle,
         &ComError,
         &ComStat);
      m_hCurrentHandle = (HANDLE)NULL;
      m_sHandleFlag = 0;
      return;
   }

   _stprintf_s(szTemp1, STANDARD_STRING, _T("%s"),
               _T("Com Port opened, intitialized, and ready for Write/Read sequence."));
   m_StatusText1.SetWindowText (szTemp1);
   m_StatusText2.SetWindowText(_T(""));

}  // End of OnOpening1DoAllOfTheAbove


void CScale3MainWin::OnOpening1CloseHandle (void)
{

   COMSTAT              ComStat;
   DWORD                ComError = 0, MyError = 0;
   int                  i;
   BOOL                 bResult1;
   short                sUnConfigedPorts = 0;
   TCHAR                szTemp1[512], szTemp2[128];

   if (m_hCurrentHandle)
   {

      szTemp2[0] = '\0';
      SecureZeroMemory (&ComStat, sizeof (COMSTAT));
      bResult1 = ClearCommError (
                     m_hCurrentHandle,
                     &ComError,
                     &ComStat);
      if (bResult1 && ComError)
      {
         _stprintf_s(szTemp2, 128, _T(" ClearCommError returned com port error = %x"), ComError);
      }

      bResult1 = CloseHandle (m_hCurrentHandle);
      if (!bResult1)
      {
         DisplaySystemError (&MyError);
         _stprintf_s(szTemp1, 512, _T("%s%d%s"),
            _T("CloseHandle failed with error "), MyError,
            _T(" File Handle = INVALID_HANDLE_VALUE"));
         _tcsncat_s (szTemp1, 512, szTemp2, _tcsnlen (szTemp2, 128));
         m_StatusText1.SetWindowText (szTemp1);
      }  // End of CloseHandle failure
      else
      {
         _stprintf_s(szTemp1, 512, _T("%s"), _T("CloseHandle succeeded."));
         _tcsncat_s (szTemp1, 512, szTemp2, _tcsnlen (szTemp2, 128));
         m_StatusText1.SetWindowText (szTemp1);
         m_StatusText2.SetWindowText(_T(""));

         if (m_sPortsConfigured)
         {
            for (i = 0; i < MAX_NUM_COMPORT; i++)
            {
               if (m_ComPorts[i].hCom == m_hCurrentHandle)
               {
                  m_ComPorts[i].hCom = NULL;
                  sUnConfigedPorts++;
               }
            }
         }
         else
            m_ComPorts[m_sCurrentPortIndex].hCom = NULL;

         if (sUnConfigedPorts == MAX_NUM_COMPORT)
         {
            m_sPortsConfigured = 0;
            m_StartThreadsPB.ShowWindow (SW_HIDE);
         }

         SecureZeroMemory (&m_dcbCurrentDCB, sizeof (DCB));
         m_dcbCurrentDCB.DCBlength = sizeof (DCB);
         // Fill in DCB: 2400 bps, 7 data bits, even parity, and 1 stop bit.
         m_dcbCurrentDCB.BaudRate = CBR_2400;
         m_dcbCurrentDCB.ByteSize = 7;
         m_dcbCurrentDCB.Parity = EVENPARITY;
         m_dcbCurrentDCB.StopBits = ONESTOPBIT;

         m_hCurrentHandle = (HANDLE)NULL;
         m_sHandleFlag = 0;

         SecureZeroMemory (&m_CurrentCommProp, sizeof (COMMPROP));

         m_sGotCommState = 0;
         m_sDCBBuiltFlag = 0;
         m_sCommStateSet = 0;
         m_sTimeoutsSet = 0;
         m_sGotTimeouts = 0;

         m_StructLBTitle.SetWindowText(_T(""));
         m_StructLB.ResetContent ();

      }  // End of CloseHandle success

   }  // End of if (m_hCurrentHandle)

}  // End of OnOpening1CloseHandle


void CScale3MainWin::OnOpening1RdWrtCurrent (void)
{

   DWORD                MyError = 0;
   DWORD                WrtBufLen = 0;
   DWORD                NumBytes = 0;
   DWORD                EventMask = 0;
   int                  i, iWrtBufLen, iCnt = 0;
   BOOL                 bResult1;
   TCHAR                szTemp1[STANDARD_STRING];
   BYTE                 aByte;
   BYTE                 aByteBuf[STANDARD_BUFFER];

   ClearStatusLines ();
   WrtBufLen = (DWORD)strlen(szCurrentCmd);
   iWrtBufLen = (int)WrtBufLen;
   bResult1 = WriteFile(
                  m_hCurrentHandle,
                  szCurrentCmd,
                  WrtBufLen,
                  &NumBytes,
                  &m_oCurrentOverlappedStruct);

#ifdef ASYNC
   bResult1 = GetOverlappedResult(m_hCurrentHandle, &m_oCurrentOverlappedStruct, &NumBytes, TRUE);
#endif

   if (!bResult1)
   {
      DisplaySystemError (&MyError);
      _stprintf_s(szTemp1, STANDARD_STRING, _T("%s%d"), _T("WriteFile failed with error number "), MyError);
      m_StatusText1.SetWindowText (szTemp1);
      return;
   }  // End of WriteFile failed

   NumBytes = 0;
   iCnt = 0;
   Sleep (WRITE_READ_DELAY);
   bResult1 = WaitCommEvent(m_hCurrentHandle, &EventMask, &m_oCurrentOverlappedStruct);

#ifdef ASYNC
   bResult1 = GetOverlappedResult(m_hCurrentHandle, &m_oCurrentOverlappedStruct, &NumBytes, TRUE);
#endif

   if (bResult1)
   {
      if (EventMask & EV_RXFLAG)
      {
         aByte = 0;
         aByteBuf[0] = 0;
         do
         {
            bResult1 = ReadFile(
                           m_hCurrentHandle,
                           &aByte,
                           1,
                           &NumBytes,
                           &m_oCurrentOverlappedStruct);

#ifdef ASYNC
            bResult1 = GetOverlappedResult(m_hCurrentHandle, &m_oCurrentOverlappedStruct, &NumBytes, TRUE);
#endif

            if (!bResult1)
            {
               DisplaySystemError (&MyError);
               _stprintf_s(szTemp1, STANDARD_STRING, _T("ReadFile failed at Count %d with error number %d"), iCnt, MyError);
               m_StatusText1.SetWindowText (szTemp1);
               return;
            }  // End of ReadFile failed

            if (NumBytes == 1)
               aByteBuf[iCnt++] = aByte;

         }
         while ((NumBytes == 1) && (iCnt < STANDARD_BUFFER));
         
         // This data representation is slightly different from the results of the other Read/Write I/O in
         // the program. I consider this operation to be more manual with the intent to see everything
         // the device is sending back.
         aByteBuf[iCnt++] = '\0';

         _stprintf_s(szTemp1, STANDARD_STRING,
                     _T("WriteFile/WaitCommEvent/ReadFile succeeded after %d ReadFile and with returned data :"),
                     iCnt);
         m_StatusText1.SetWindowText (szTemp1);

         /***************************************************************************************
         June 12th, 2015

         Unicode alert. My serial device only produces 8 bit bytes. It is too old to know anything about
         possible non-romance language use. So I have a mismatch when trying to use the data produced
         by the instrument in a 16 bit wide or 32 bit wide char user interface. I collect data as 8 bit
         bytes but must use Unicode strings. When I make use of _stprintf_s to transport
         the data/information to the user interface, I have to use an upper case S in the format specifier
         as opposed to a lower case S.

         ***************************************************************************************/

#ifdef BS2PIN16
         for (i = 0; (i < iWrtBufLen) && ((i + iWrtBufLen) < iCnt); i++)
            aByteBuf[i] = aByteBuf[i + WrtBufLen];
#endif

         _stprintf_s(szTemp1, STANDARD_STRING, _T("%S"), aByteBuf);
         m_StatusText2.SetWindowText (szTemp1);

      }  // End of if (EventMask & EV_RXFLAG)
      else
      {
         _stprintf_s(szTemp1, STANDARD_STRING,
                     _T("WaitCommEvent's EventMask != to EV_RXFLAG! ReadFile not done! No Data!"));
         m_StatusText1.SetWindowText (szTemp1);
         _stprintf_s(szTemp1, STANDARD_STRING, _T("WaitCommEvent's EventMask = %x "), EventMask);
         m_StatusText2.SetWindowText (szTemp1);
         return;
      }

   }  // End of if DCB's event-char received and no overlap
   else
   {
      DisplaySystemError (&MyError);
      _stprintf_s(szTemp1, STANDARD_STRING, _T("WaitCommEvent failed error number %d"), MyError);
      m_StatusText1.SetWindowText (szTemp1);
      return;
   }  // End of WaitCommEvent overlapped

}  // End of OnOpening1RdWrtCurrent


void CScale3MainWin::OnOpening1RdWrtConfigured (void)
{

   BYTE                 aByteBuf[STANDARD_BUFFER];
   TCHAR                szTemp1[STANDARD_STRING];
   int                  i, iResult1 = 1;

   ClearStatusLines ();
   if (m_sShowConfig)
   {
      m_ConfiguredLBTitle.ShowWindow (SW_SHOW);
      m_ConfiguredLB.ShowWindow (SW_SHOW);
      m_sShowConfig = 0;
   }
   m_ConfiguredLB.ResetContent ();

   for (i = 0; i < MAX_NUM_COMPORT; i++)
   {
      if (m_ComPorts[i].sDCBSet && m_ComPorts[i].sTimeoutsSet)
      {
         iResult1 = WrtReadPort (m_ComPorts[i].hCom,
                                 szCurrentCmd,
                                 &(m_ComPorts[i].OverLap),
                                 aByteBuf,
                                 STANDARD_BUFFER);
         if (!iResult1)
         {
            _stprintf_s(szTemp1, STANDARD_STRING, _T("%s - %S"), m_ComPorts[i].szName, aByteBuf);
            m_ConfiguredLB.AddString (szTemp1);
         }
      }

   }  // End of for (i = 0; i < MAX_NUM_COMPORT; i++)

}  // End of OnOpening1RdWrtConfigured


////////////////////////// Standard messages ////////////////////////////

void CScale3MainWin::OnSize (UINT nType, int cx, int cy)
{

   CFrameWnd::OnSize (nType, cx, cy);
   GetClientRect(&m_rect);
   m_iClientWidth = m_rect.Width ();
   m_iClientHeight = m_rect.Height ();

}  // End of OnSize


void CScale3MainWin::PostNcDestroy ()
{

   delete this;

}


void CScale3MainWin::OnPaint (void)
{

   CFrameWnd::OnPaint ();


}  // End of OnPaint method.


   /*These calls to ResumeThread cause the thread to execute the first executable line of code after the
   SuspendThread call. The SuspendThread call is made by the thread itself and the thread sits at the
   SuspendThread line until a ResumeThread call is made.*/
void CScale3MainWin::OnTimer (UINT nIDEvent)
{

   if (nIDEvent == THREADTIMER)  // DEFAULTTHREADTIMERVALUE (5000) or m_iThreadTimerValue
   {

      m_ConfiguredLB.ResetContent ();

      if (m_sThread1Running)
         m_Port1ThreadHandle->ResumeThread ();
      if (m_sThread2Running)
         m_Port2ThreadHandle->ResumeThread ();
      if (m_sThread3Running)
         m_Port3ThreadHandle->ResumeThread ();
      if (m_sThread4Running)
         m_Port4ThreadHandle->ResumeThread ();

   }

}  // End of OnTimer


//////////////////////// Control's Messages /////////////////////////////

void CScale3MainWin::OnSelchangePortList (void)
{

   short sNewSelection, sTemp1;

   sNewSelection = m_ComPortLB.GetCurSel ();

   m_ComPorts[m_sCurrentPortIndex].BaudRate = ulBaudRates[m_sCurrentBaudRateIndex];
   m_ComPorts[m_sCurrentPortIndex].BaudRateIndex = m_sCurrentBaudRateIndex;
   _stprintf_s(m_ComPorts[m_sCurrentPortIndex].Parity, 16, _T("%s"), (LPCTSTR)szParity[m_sCurrentParityIndex]);
   m_ComPorts[m_sCurrentPortIndex].ParityIndex = m_sCurrentParityIndex;
   m_ComPorts[m_sCurrentPortIndex].DataBits = sDataBits[m_sCurrentDataBitsIndex];
   m_ComPorts[m_sCurrentPortIndex].DataBitsIndex = m_sCurrentDataBitsIndex;
   m_ComPorts[m_sCurrentPortIndex].StopBits = fStopBits[m_sCurrentStopBitsIndex];
   m_ComPorts[m_sCurrentPortIndex].StopBits = m_sCurrentStopBitsIndex;

   m_ComPorts[m_sCurrentPortIndex].hCom = m_hCurrentHandle;
   memcpy (&(m_ComPorts[m_sCurrentPortIndex].CommTimeouts), &m_CurrentCommTimeouts, sizeof (COMMTIMEOUTS));
   memcpy (&(m_ComPorts[m_sCurrentPortIndex].dcb), &m_dcbCurrentDCB, sizeof (DCB));
   memcpy (&(m_ComPorts[m_sCurrentPortIndex].OverLap), &m_oCurrentOverlappedStruct, sizeof (OVERLAPPED));


   m_ComPorts[m_sCurrentPortIndex].sDCBSet = m_sCommStateSet;
   m_ComPorts[m_sCurrentPortIndex].sTimeoutsSet = m_sTimeoutsSet;

   if (!m_sPortsConfigured)
   {
      if (m_sHandleFlag && m_sCommStateSet && m_sTimeoutsSet)
      {
         m_sPortsConfigured = 1;
         m_StartThreadsPB.ShowWindow (SW_SHOW);
      }
   }

   sTemp1 = m_sCurrentPortIndex;
   m_sCurrentPortIndex = sNewSelection;

   m_sCurrentBaudRateIndex = m_ComPorts[m_sCurrentPortIndex].BaudRateIndex;
   m_sCurrentParityIndex = m_ComPorts[m_sCurrentPortIndex].ParityIndex;
   m_sCurrentDataBitsIndex = m_ComPorts[m_sCurrentPortIndex].DataBitsIndex;
   m_sCurrentStopBitsIndex = m_ComPorts[m_sCurrentPortIndex].StopBitsIndex;

   m_hCurrentHandle = m_ComPorts[m_sCurrentPortIndex].hCom;
   memcpy (&m_CurrentCommTimeouts, &(m_ComPorts[m_sCurrentPortIndex].CommTimeouts), sizeof (COMMTIMEOUTS));
   memcpy (&m_dcbCurrentDCB, &(m_ComPorts[m_sCurrentPortIndex].dcb), sizeof (DCB));
   memcpy (&m_oCurrentOverlappedStruct, &(m_ComPorts[m_sCurrentPortIndex].OverLap), sizeof (OVERLAPPED));

   if (m_hCurrentHandle)
      m_sHandleFlag = 1;
   else
      m_sHandleFlag = 0;

   SecureZeroMemory (&m_CurrentCommProp, sizeof (COMMPROP));

   m_sCommStateSet = m_ComPorts[m_sCurrentPortIndex].sDCBSet;
   m_sTimeoutsSet = m_ComPorts[m_sCurrentPortIndex].sTimeoutsSet;
   if (m_sCommStateSet)
   {
      m_sGotCommState = 1;
      m_sDCBBuiltFlag = 1;
   }
   else
   {
      m_sGotCommState = 0;
      m_sDCBBuiltFlag = 0;
   }

   if (m_sTimeoutsSet)
      m_sGotTimeouts = 1;
   else
      m_sGotTimeouts = 0;

   UpDateComboBoxes ();

   //m_StructLBTitle.SetWindowTextA ("");
   //m_StructLB.ResetContent ();
   if (!m_sShowStruct)
   {
      m_StructLBTitle.ShowWindow (SW_HIDE);
      m_StructLB.ShowWindow (SW_HIDE);
      m_sShowStruct = 1;
   }

   ClearStatusLines ();
   m_StatusText1.SetWindowText(_T("Current Device Path :"));
   m_StatusText2.SetWindowText (m_ComPorts[m_sCurrentPortIndex].szDevicePath);
   if (!m_sShowConfig)
   {
      m_ConfiguredLBTitle.ShowWindow (SW_HIDE);
      m_ConfiguredLB.ShowWindow (SW_HIDE);
      m_sShowConfig = 1;
   }

}  // End of OnSelchangePortList


void CScale3MainWin::OnCloseUpBaudRate (void)
{

   int               nSel;

   nSel = m_BaudRateCB.GetCurSel ();
   m_sCurrentBaudRateIndex = nSel;

}  // End of OnCloseUpBaudRate


void CScale3MainWin::OnCloseUpDataBits (void)
{

   int               nSel;

   nSel = m_DataBitsCB.GetCurSel ();
   m_sCurrentDataBitsIndex = nSel;

}  // End of OnCloseUpDataBits


void CScale3MainWin::OnCloseUpStopBits (void)
{

   int               nSel;

   nSel = m_StopBitsCB.GetCurSel ();
   m_sCurrentStopBitsIndex = nSel;

}  // End of OnCloseUpStopBits


void CScale3MainWin::OnCloseUpParity (void)
{

   int               nSel;

   nSel = m_ParityCB.GetCurSel ();
   m_sCurrentParityIndex = nSel;

}  // End of OnCloseUpParity


/*AfxBeginThread with the CREATE_SUSPENDED flag starts the worker thread. When you call ResumeThread
for the first time, the thread begins at the first executable line. It seems that the declarations and
assignments have been done.*/
void CScale3MainWin::OnClickedStartThreads (void)
{

   int                        i;
   short                      sTemp1 = 0;

   for (i = 0; i < MAX_NUM_COMPORT; i++)
   {
      if (m_ComPorts[i].hCom != NULL)
      {

         tp[i].hWnd = m_hWnd;
         strcpy_s(tp[i].Cmd, STANDARD_STRING, szCurrentCmd);
         //_stprintf_s(tp[i].Cmd, STANDARD_STRING, _T("%s"), (LPCTSTR)szCurrentCmd);
         tp[i].hCom = m_ComPorts[i].hCom;
         memcpy_s (&(tp[i].oOverlappedStruct), sizeof (OVERLAPPED),
                   &(m_ComPorts[i].OverLap), sizeof (OVERLAPPED));

         if (i == 0)
         {
            m_Port1ThreadHandle = AfxBeginThread (
                                    Port1Thread,
                                    (LPVOID)&i,
                                    THREAD_PRIORITY_NORMAL,
                                    0,
                                    CREATE_SUSPENDED);
            if (m_Port1ThreadHandle != NULL)
            {
               m_sThread1Running = 1;
               tp[0].hThread = m_Port1ThreadHandle->m_hThread;
               tp[0].sEndThreadFlag = 0;
            }
         }
         else if (i == 1)
         {
            m_Port2ThreadHandle = AfxBeginThread (
                                    Port2Thread,
                                    (LPVOID)&i,
                                    THREAD_PRIORITY_NORMAL,
                                    0,
                                    CREATE_SUSPENDED);
            if (m_Port2ThreadHandle != NULL)
            {
               m_sThread2Running = 1;
               tp[1].hThread = m_Port2ThreadHandle->m_hThread;
               tp[1].sEndThreadFlag = 0;
            }
         }
         else if (i == 2)
         {
            m_Port3ThreadHandle = AfxBeginThread (
                                    Port3Thread,
                                    (LPVOID)&i,
                                    THREAD_PRIORITY_NORMAL,
                                    0,
                                    CREATE_SUSPENDED);
            if (m_Port3ThreadHandle != NULL)
            {
               m_sThread3Running = 1;
               tp[2].hThread = m_Port3ThreadHandle->m_hThread;
               tp[2].sEndThreadFlag = 0;
            }
         }
         else if (i == 3)
         {
            m_Port4ThreadHandle = AfxBeginThread (
                                    Port4Thread,
                                    (LPVOID)&i,
                                    THREAD_PRIORITY_NORMAL,
                                    0,
                                    CREATE_SUSPENDED);
            if (m_Port4ThreadHandle != NULL)
            {
               m_sThread4Running = 1;
               tp[3].hThread = m_Port4ThreadHandle->m_hThread;
               tp[3].sEndThreadFlag = 0;
            }
         }

      }  // End of if (m_ComPorts[i].hCom != NULL)

   }  // End of for (i = 0; i < MAX_NUM_COMPORT; i++)

   sTemp1 = m_sThread1Running + m_sThread2Running + m_sThread3Running + m_sThread4Running;
   if (sTemp1)
   {
      m_StartThreadsPB.EnableWindow (false);
      m_ComPortLB.EnableWindow (false);
      m_UseThreadsPB.ShowWindow (SW_SHOW);
      m_TimerUseThreadsPB.ShowWindow (SW_SHOW);
      m_EndThreadsPB.ShowWindow (SW_SHOW);
      m_ThreadTimerEBTitle.ShowWindow (SW_SHOW);
      m_ThreadTimerEB.ShowWindow (SW_SHOW);
   }

}  // End of OnClickedStartThreads

/*These calls to ResumeThread cause the thread to execute the first executable line of code after the
SuspendThread call. The SuspendThread call is made by the thread itself and the thread sits at the
SuspendThread line until a ResumeThread call is made.*/
void CScale3MainWin::OnClickedUseThreads (void)
{

   ClearStatusLines ();
   if (m_sShowConfig)
   {
      m_ConfiguredLBTitle.ShowWindow (SW_SHOW);
      m_ConfiguredLB.ShowWindow (SW_SHOW);
      m_sShowConfig = 0;
   }
   m_ConfiguredLB.ResetContent ();

   if (m_sThread1Running)
      m_Port1ThreadHandle->ResumeThread ();
   if (m_sThread2Running)
      m_Port2ThreadHandle->ResumeThread ();
   if (m_sThread3Running)
      m_Port3ThreadHandle->ResumeThread ();
   if (m_sThread4Running)
      m_Port4ThreadHandle->ResumeThread ();

}  // End of OnClickedUseThreads


void CScale3MainWin::OnClickedTimerUseThreads (void)
{

   TCHAR                      szTemp1[16];
   int                        iTemp1 = DEFAULTTHREADTIMERVALUE;

   ClearStatusLines ();
   if (m_sShowConfig)
   {
      m_ConfiguredLBTitle.ShowWindow (SW_SHOW);
      m_ConfiguredLB.ShowWindow (SW_SHOW);
      m_sShowConfig = 0;
   }
   m_ConfiguredLB.ResetContent ();
   m_UseThreadsPB.EnableWindow (false);

   szTemp1[0] = '\0';
   m_ThreadTimerEB.GetWindowText (szTemp1, 16);
   iTemp1 = _tstoi (szTemp1);
   if ((iTemp1 < 1) || (iTemp1 > 100000))
   {
      m_iThreadTimerValue = DEFAULTTHREADTIMERVALUE;
      _stprintf_s (szTemp1, 16, _T("%d"), m_iThreadTimerValue);
      m_ThreadTimerEB.SetWindowText (szTemp1);
   }
   m_iThreadTimerValue = iTemp1;

   m_ThreadTimer = SetTimer (THREADTIMER, m_iThreadTimerValue, NULL);
   if (!m_ThreadTimer)
      KillTimer (m_ThreadTimer);

}  // End of OnClickedTimerUseThreads


/*Because the threads are using WaitCommEvent while the com port handles are configured for synchronous I/O,
there is the potential for the thread to be in the WaitCommEvent function forever. This could happen if a
comm port read is attempted on a commm port that is not actually attached to a live serial device. You don't
want to abandon a thread that is waiting for WaitCommEvent to return. Since I really want to use 
synchronous I/O, I avoid this situation by using the desktop only function CancelSynchronousIo before I
try killing the thread.
Another issue is the use of AfxEndThread. Why am I depending on the thread to make the call? Ideally, the
thread is suspended when I make this call to ResumeThread and I'm assuming the very next thing the thread
will do is compare a static global value to see if it should AfxEndThread. If my intent is to have this
function end the thread, then why not have the function call AfxEndThread? A possible scenerio is to have
this function issue a CancelSynchronousIo then wait for a bit, 1 sec maybe, and then issue the AfxEndThread.
But there must be a problem with the thread's suspend/resume counter. Here a typical problem scenerio.
The thread is in WaitCommEvent and is going to stay there because there is nn device connected to the serial
port. The thread's suspend/resume counter is at zero, but it might as will be in suspend mode. This function
comes along and does a ResumeThread. "If the function succeeds, the return value is the thread's
previous suspend count. If the function fails, the return value is (DWORD) -1. 
If the return value is zero, the specified thread was not suspended.
If the return value is 1, the specified thread was suspended but was restarted.
If the return value is greater than 1, the specified thread is still suspended."
If the thread is in WaitCommEvent, ResumeThread returns zero and this indicates that that I need to use
CancelSynchronousIO. That will clean up the serial IO and allow the parent to close the comport handle.
(If the comport handle is in WaitCommEvent and you attempt to close the comport handle the parent process
will lockup.) However, I still have to take care of terminating the thread. CancelSynchronousIO will
cause the thread to PostMessage an IO error because of erroring out of the WaitCommEvent. The thread will
then suspend itself. I will get the thread to AfxEndThread itself with another call to ResumeThread since
the sEndThreadFlag is set but the thread will not have suspended itself by the time I issue the next
ResumeThread. So I will wait for 400mS until I issue the ResumeThread. I'll try 12 times and while I should
flag the error condition if I go over 12 Waits, I'm not doing so right now.*/

void CScale3MainWin::EndThreads(void)
{

   BOOL                 bResult1;
   DWORD                MyError = 0;
   int                  iCnt = 0;

   if (m_sThread1Running)
   {
      tp[0].sEndThreadFlag = 1;
      MyError = 0;
      MyError = m_Port1ThreadHandle->ResumeThread();
      if (!MyError)
      {
         iCnt = 0;
         bResult1 = CancelSynchronousIo(tp[0].hThread);
         while ((iCnt <= 12) && (MyError != 1))
         {
            Sleep(WRITE_READ_DELAY);
            MyError = m_Port1ThreadHandle->ResumeThread();
            iCnt++;
         }
         //if (iCnt >= 12) Something is really wrong! You better do something!
      }
      m_sThread1Running = 0;
   }
   if (m_sThread2Running)
   {
      tp[1].sEndThreadFlag = 1;
      MyError = 0;
      MyError = m_Port2ThreadHandle->ResumeThread();
      if (!MyError)
      {
         iCnt = 0;
         bResult1 = CancelSynchronousIo(tp[1].hThread);
         while ((iCnt <= 12) && (MyError != 1))
         {
            Sleep(WRITE_READ_DELAY);
            MyError = m_Port2ThreadHandle->ResumeThread();
            iCnt++;
         }
         //if (iCnt >= 12) Something is really wrong! You better do something!
      }
      m_sThread2Running = 0;
   }
   if (m_sThread3Running)
   {
      tp[2].sEndThreadFlag = 1;
      MyError = 0;
      MyError = m_Port3ThreadHandle->ResumeThread();
      if (!MyError)
      {
         iCnt = 0;
         bResult1 = CancelSynchronousIo(tp[2].hThread);
         while ((iCnt <= 12) && (MyError != 1))
         {
            Sleep(WRITE_READ_DELAY);
            MyError = m_Port3ThreadHandle->ResumeThread();
            iCnt++;
         }
         //if (iCnt >= 12) Something is really wrong! You better do something!
      }
      m_sThread3Running = 0;
   }
   if (m_sThread4Running)
   {
      tp[3].sEndThreadFlag = 1;
      MyError = 0;
      MyError = m_Port4ThreadHandle->ResumeThread();
      if (!MyError)
      {
         iCnt = 0;
         bResult1 = CancelSynchronousIo(tp[3].hThread);
         while ((iCnt <= 12) && (MyError != 1))
         {
            Sleep(WRITE_READ_DELAY);
            MyError = m_Port4ThreadHandle->ResumeThread();
            iCnt++;
         }
         //if (iCnt >= 12) Something is really wrong! You better do something!
      }
      m_sThread4Running = 0;
   }

}  // EndThreads

void CScale3MainWin::OnClickedEndThreads (void)
{

   if (m_ThreadTimer != 0)
   {
      KillTimer (THREADTIMER);
      m_ThreadTimer = 0;
   }

   EndThreads ();

   m_TimerUseThreadsPB.EnableWindow (true);
   m_UseThreadsPB.ShowWindow (SW_HIDE);
   m_TimerUseThreadsPB.ShowWindow (SW_HIDE);
   m_EndThreadsPB.ShowWindow (SW_HIDE);
   m_ThreadTimerEBTitle.ShowWindow (SW_HIDE);
   m_ThreadTimerEB.ShowWindow (SW_HIDE);
   m_StartThreadsPB.EnableWindow (true);
   m_UseThreadsPB.EnableWindow (true);
   m_ComPortLB.EnableWindow (true);

   ClearStatusLines ();
   m_ConfiguredLBTitle.ShowWindow (SW_HIDE);
   m_ConfiguredLB.ShowWindow (SW_HIDE);
   m_sShowConfig = 1;

}  // End of OnClickedEndThreads


void CScale3MainWin::OnClickedTest (void)
{

   if (!m_ThreadTimer)
      m_ThreadTimer = SetTimer (THREADTIMER, m_iThreadTimerValue, NULL);
   else
   {
      KillTimer (m_ThreadTimer);
      m_ThreadTimer = 0;
   }

}  // End of OnClickedTest


////////////////////////////// Class Specific Functions /////////////////////////

void CScale3MainWin::FillComPortLB ()
{

   TCHAR          szString1[1024];
   int            i;

   CString        str;
   CSize          sz;
   int            dx = 0;
   TEXTMETRIC     tm;
   CDC            *pDC;
   CFont          *pFont, *pOldFont;

   m_ComPortLB.ResetContent ();

   //EnumerateComPorts ();
   //EnumerateXPComPorts ();
   EnumerateStartechWin8ComPorts();

   for (i = 0; i < MAX_NUM_COMPORT; i++)
   {
      if (m_ComPorts[i].ComFlag == FILLED)
      {
         szString1[0] = '\0';
         _stprintf_s (szString1, 1024, _T("%s %s  %s  %s  %s  %s"),
            m_ComPorts[i].szLocationInformation, m_ComPorts[i].szFriendlyName,
            m_ComPorts[i].szManufacturer, m_ComPorts[i].szDeviceDescrp,
            m_ComPorts[i].szDeviceObject, m_ComPorts[i].szDevicePath);
            m_ComPortLB.AddString (szString1);
      }
   }

   m_ComPortLB.SetCurSel (m_sCurrentPortIndex);
   // LB_ERR = -1, for your information

   // Select the listbox font, save the old font
   // Get the text metrics for avg char width
   pDC = m_ComPortLB.GetDC ();
   pFont = m_ComPortLB.GetFont ();
   pOldFont = pDC->SelectObject (pFont);
   pDC->GetTextMetrics (&tm);

   // Find the longest string in the list box.
   for (i = 0; i < m_ComPortLB.GetCount (); i++)
   {
      m_ComPortLB.GetText (i, str);
      sz = pDC->GetTextExtent (str);

      // Add the avg width to prevent clipping
      sz.cx += tm.tmAveCharWidth;

      if (sz.cx > dx)
         dx = sz.cx;
   }

   // Select the old font back into the DC
   pDC->SelectObject (pOldFont);
   m_ComPortLB.ReleaseDC(pDC);

   // Set the horizontal extent so every character of all strings 
   // can be scrolled to.
   m_ComPortLB.SetHorizontalExtent(dx);

}  // End of FillComPortLB


void CScale3MainWin::SetComboBoxes (void)
{

   short i;
   TCHAR szTemp1[32];

   m_BaudRateCB.ResetContent ();
   m_DataBitsCB.ResetContent ();
   m_StopBitsCB.ResetContent ();
   m_ParityCB.ResetContent ();

   for (i = 0; i < 13; i++)
   {

      szTemp1[0] = '\0';
      _stprintf_s (szTemp1, 32, _T("%d"), ulBaudRates[i]);
      m_BaudRateCB.AddString (szTemp1);

      if (i < 3)
      {
         szTemp1[0] = '\0';
         _stprintf_s (szTemp1, 32, _T("%1.1f"), fStopBits[i]);
         m_StopBitsCB.AddString (szTemp1);
      }
      if (i < 4)
      {
         szTemp1[0] = '\0';
         _stprintf_s (szTemp1, 32, _T("%d"), sDataBits[i]);
         m_DataBitsCB.AddString (szTemp1);
      }
      if (i < 5)
         m_ParityCB.AddString (szParity[i]);

   }

   // Set them to the known defaults for the MT Scales
   m_BaudRateCB.SetCurSel (m_sCurrentBaudRateIndex);
   m_DataBitsCB.SetCurSel (m_sCurrentDataBitsIndex);
   m_StopBitsCB.SetCurSel (m_sCurrentStopBitsIndex);
   m_ParityCB.SetCurSel (m_sCurrentParityIndex);

}  // End of SetComboBoxes


void CScale3MainWin::UpDateComboBoxes (void)
{

   // Set them to the known defaults for the MT Scales
   m_BaudRateCB.SetCurSel (m_sCurrentBaudRateIndex);
   m_DataBitsCB.SetCurSel (m_sCurrentDataBitsIndex);
   m_StopBitsCB.SetCurSel (m_sCurrentStopBitsIndex);
   m_ParityCB.SetCurSel (m_sCurrentParityIndex);

}  // End of UpDateComboBoxes


//   CStatic        m_StructEFTitle
//   CEdit          m_StructEF
//   DCB            m_dcbCurrentDCB;         CString DCBMembers[28]
//   COMMPROP       m_CurrentCommProp;       CString COMMPROPMembers[18]
//   COMMTIMEOUTS   m_CurrentCommTimeouts;   CString COMMTIMEOUTSMembers [5]
void CScale3MainWin::OutputProperties (void)
{

   TCHAR                   szTemp1[128];
   int                     i;

   if (m_sShowStruct)
   {
      m_StructLBTitle.ShowWindow (SW_SHOW);
      m_StructLB.ShowWindow (SW_SHOW);
      m_sShowStruct = 0;
   }

   m_StructLBTitle.SetWindowText (_T("Current Port's Properties"));
   m_StructLB.ResetContent ();
   i = 0;
   _stprintf_s(szTemp1, 128, _T("%s = %d"), (LPCTSTR)COMMPROPMembers[i++], m_CurrentCommProp.wPacketLength);
   m_StructLB.AddString (szTemp1);
   _stprintf_s(szTemp1, 128, _T("%s = %d"), (LPCTSTR)COMMPROPMembers[i++], m_CurrentCommProp.wPacketVersion);
   m_StructLB.AddString (szTemp1);
   _stprintf_s(szTemp1, 128, _T("%s = %d"), (LPCTSTR)COMMPROPMembers[i++], m_CurrentCommProp.dwServiceMask);
   m_StructLB.AddString (szTemp1);
   _stprintf_s(szTemp1, 128, _T("%s = %d"), (LPCTSTR)COMMPROPMembers[i++], m_CurrentCommProp.dwReserved1);
   m_StructLB.AddString (szTemp1);
   _stprintf_s(szTemp1, 128, _T("%s = %d"), (LPCTSTR)COMMPROPMembers[i++], m_CurrentCommProp.dwMaxTxQueue);
   m_StructLB.AddString (szTemp1);
   _stprintf_s(szTemp1, 128, _T("%s = %d"), (LPCTSTR)COMMPROPMembers[i++], m_CurrentCommProp.dwMaxRxQueue);
   m_StructLB.AddString (szTemp1);
   _stprintf_s(szTemp1, 128, _T("%s = 0x%X"), (LPCTSTR)COMMPROPMembers[i++], m_CurrentCommProp.dwMaxBaud);
   m_StructLB.AddString (szTemp1);
   _stprintf_s(szTemp1, 128, _T("%s = %d"), (LPCTSTR)COMMPROPMembers[i++], m_CurrentCommProp.dwProvSubType);
   m_StructLB.AddString (szTemp1);
   _stprintf_s(szTemp1, 128, _T("%s = 0x%X"), (LPCTSTR)COMMPROPMembers[i++], m_CurrentCommProp.dwProvCapabilities);
   m_StructLB.AddString (szTemp1);
   _stprintf_s(szTemp1, 128, _T("%s = 0x%X"), (LPCTSTR)COMMPROPMembers[i++], m_CurrentCommProp.dwSettableParams);
   m_StructLB.AddString (szTemp1);
   _stprintf_s(szTemp1, 128, _T("%s = 0x%X"), (LPCTSTR)COMMPROPMembers[i++], m_CurrentCommProp.dwSettableBaud);
   m_StructLB.AddString (szTemp1);
   _stprintf_s(szTemp1, 128, _T("%s = 0x%X"), (LPCTSTR)COMMPROPMembers[i++], m_CurrentCommProp.wSettableData);
   m_StructLB.AddString (szTemp1);
   _stprintf_s(szTemp1, 128, _T("%s = 0x%X"), (LPCTSTR)COMMPROPMembers[i++], m_CurrentCommProp.wSettableStopParity);
   m_StructLB.AddString (szTemp1);
   _stprintf_s(szTemp1, 128, _T("%s = %d"), (LPCTSTR)COMMPROPMembers[i++], m_CurrentCommProp.dwCurrentTxQueue);
   m_StructLB.AddString (szTemp1);
   _stprintf_s(szTemp1, 128, _T("%s = %d"), (LPCTSTR)COMMPROPMembers[i++], m_CurrentCommProp.dwCurrentRxQueue);
   m_StructLB.AddString (szTemp1);
   _stprintf_s(szTemp1, 128, _T("%s = %d"), (LPCTSTR)COMMPROPMembers[i++], m_CurrentCommProp.dwProvSpec1);
   m_StructLB.AddString (szTemp1);
   _stprintf_s(szTemp1, 128, _T("%s = %d"), (LPCTSTR)COMMPROPMembers[i++], m_CurrentCommProp.dwProvSpec2);
   m_StructLB.AddString (szTemp1);
   _stprintf_s (szTemp1, 128, _T("%s = %s"), (LPCTSTR)COMMPROPMembers[i++], m_CurrentCommProp.wcProvChar);
   m_StructLB.AddString (szTemp1);

}  // End of OutputProperties


void CScale3MainWin::OutputCommState (void)
{

   TCHAR                   szTemp1[128];
   int                     i;

   if (m_sShowStruct)
   {
      m_StructLBTitle.ShowWindow (SW_SHOW);
      m_StructLB.ShowWindow (SW_SHOW);
      m_sShowStruct = 0;
   }

   m_StructLBTitle.SetWindowText(_T("Current Port's DCB Structure"));
   m_StructLB.ResetContent ();
   i = 0;
   _stprintf_s (szTemp1, 128, _T("%s = %d"), (LPCTSTR)DCBMembers[i++], m_dcbCurrentDCB.DCBlength);
   m_StructLB.AddString (szTemp1);
   _stprintf_s (szTemp1, 128, _T("%s = %d"), (LPCTSTR)DCBMembers[i++], m_dcbCurrentDCB.BaudRate);
   m_StructLB.AddString (szTemp1);
   _stprintf_s (szTemp1, 128, _T("%s = %d"), (LPCTSTR)DCBMembers[i++], m_dcbCurrentDCB.fBinary);
   m_StructLB.AddString (szTemp1);
   _stprintf_s (szTemp1, 128, _T("%s = %d"), (LPCTSTR)DCBMembers[i++], m_dcbCurrentDCB.fParity);
   m_StructLB.AddString (szTemp1);
   _stprintf_s (szTemp1, 128, _T("%s = %d"), (LPCTSTR)DCBMembers[i++], m_dcbCurrentDCB.fOutxCtsFlow);
   m_StructLB.AddString (szTemp1);
   _stprintf_s (szTemp1, 128, _T("%s = %d"), (LPCTSTR)DCBMembers[i++], m_dcbCurrentDCB.fOutxDsrFlow);
   m_StructLB.AddString (szTemp1);
   _stprintf_s (szTemp1, 128, _T("%s = %d"), (LPCTSTR)DCBMembers[i++], m_dcbCurrentDCB.fDtrControl);
   m_StructLB.AddString (szTemp1);
   _stprintf_s (szTemp1, 128, _T("%s = %d"), (LPCTSTR)DCBMembers[i++], m_dcbCurrentDCB.fDsrSensitivity);
   m_StructLB.AddString (szTemp1);
   _stprintf_s (szTemp1, 128, _T("%s = %d"), (LPCTSTR)DCBMembers[i++], m_dcbCurrentDCB.fTXContinueOnXoff);
   m_StructLB.AddString (szTemp1);
   _stprintf_s (szTemp1, 128, _T("%s = %d"), (LPCTSTR)DCBMembers[i++], m_dcbCurrentDCB.fOutX);
   m_StructLB.AddString (szTemp1);
   _stprintf_s (szTemp1, 128, _T("%s = %d"), (LPCTSTR)DCBMembers[i++], m_dcbCurrentDCB.fInX);
   m_StructLB.AddString (szTemp1);
   _stprintf_s (szTemp1, 128, _T("%s = %d"), (LPCTSTR)DCBMembers[i++], m_dcbCurrentDCB.fErrorChar);
   m_StructLB.AddString (szTemp1);
   _stprintf_s (szTemp1, 128, _T("%s = %d"), (LPCTSTR)DCBMembers[i++], m_dcbCurrentDCB.fNull);
   m_StructLB.AddString (szTemp1);
   _stprintf_s (szTemp1, 128, _T("%s = %d"), (LPCTSTR)DCBMembers[i++], m_dcbCurrentDCB.fRtsControl);
   m_StructLB.AddString (szTemp1);
   _stprintf_s (szTemp1, 128, _T("%s = %d"), (LPCTSTR)DCBMembers[i++], m_dcbCurrentDCB.fAbortOnError);
   m_StructLB.AddString (szTemp1);
   _stprintf_s (szTemp1, 128, _T("%s = %d"), (LPCTSTR)DCBMembers[i++], m_dcbCurrentDCB.fDummy2);
   m_StructLB.AddString (szTemp1);
   
   //_stprintf_s (szTemp1, 128, _T("%s = %d"), DCBMembers[i++], m_dcbCurrentDCB.wReserved);
   //m_StructEF.ReplaceSel (szTemp1);
   i++;

   _stprintf_s (szTemp1, 128, _T("%s = %d"), (LPCTSTR)DCBMembers[i++], m_dcbCurrentDCB.XonLim);
   m_StructLB.AddString (szTemp1);
   _stprintf_s(szTemp1, 128, _T("%s = %d"), (LPCTSTR)DCBMembers[i++], m_dcbCurrentDCB.XoffLim);
   m_StructLB.AddString (szTemp1);
   _stprintf_s (szTemp1, 128, _T("%s = %d"), (LPCTSTR)DCBMembers[i++], m_dcbCurrentDCB.ByteSize);
   m_StructLB.AddString (szTemp1);
   _stprintf_s (szTemp1, 128, _T("%s = %d"), (LPCTSTR)DCBMembers[i++], m_dcbCurrentDCB.Parity);
   m_StructLB.AddString (szTemp1);
   _stprintf_s (szTemp1, 128, _T("%s = %d"), (LPCTSTR)DCBMembers[i++], m_dcbCurrentDCB.StopBits);
   m_StructLB.AddString (szTemp1);
   _stprintf_s (szTemp1, 128, _T("%s = %d"), (LPCTSTR)DCBMembers[i++], m_dcbCurrentDCB.XonChar);
   m_StructLB.AddString (szTemp1);
   _stprintf_s (szTemp1, 128, _T("%s = %d"), (LPCTSTR)DCBMembers[i++], m_dcbCurrentDCB.XoffChar);
   m_StructLB.AddString (szTemp1);
   _stprintf_s (szTemp1, 128, _T("%s = %d"), (LPCTSTR)DCBMembers[i++], m_dcbCurrentDCB.ErrorChar);
   m_StructLB.AddString (szTemp1);
   _stprintf_s (szTemp1, 128, _T("%s = %d"), (LPCTSTR)DCBMembers[i++], m_dcbCurrentDCB.EofChar);
   m_StructLB.AddString (szTemp1);
   _stprintf_s (szTemp1, 128, _T("%s = %d"), (LPCTSTR)DCBMembers[i++], m_dcbCurrentDCB.EvtChar);
   m_StructLB.AddString (szTemp1);
   _stprintf_s (szTemp1, 128, _T("%s = %d"), (LPCTSTR)DCBMembers[i++], m_dcbCurrentDCB.wReserved1);
   m_StructLB.AddString (szTemp1);

}  // End of OutputCommState


void CScale3MainWin::OutputCommTimeouts (void)
{

   TCHAR                   szTemp1[128];
   int                     i;

   if (m_sShowStruct)
   {
      m_StructLBTitle.ShowWindow (SW_SHOW);
      m_StructLB.ShowWindow (SW_SHOW);
      m_sShowStruct = 0;
   }

   m_StructLBTitle.SetWindowText (_T("Current Port's Timeouts"));
   m_StructLB.ResetContent ();
   i = 0;
   _stprintf_s(szTemp1, 128, _T("%s = %d"), (LPCTSTR)COMMTIMEOUTSMembers[i++], m_CurrentCommTimeouts.ReadIntervalTimeout);
   m_StructLB.AddString (szTemp1);
   _stprintf_s(szTemp1, 128, _T("%s = %d"), (LPCTSTR)COMMTIMEOUTSMembers[i++], m_CurrentCommTimeouts.ReadTotalTimeoutMultiplier);
   m_StructLB.AddString (szTemp1);
   _stprintf_s(szTemp1, 128, _T("%s = %d"), (LPCTSTR)COMMTIMEOUTSMembers[i++], m_CurrentCommTimeouts.ReadTotalTimeoutConstant);
   m_StructLB.AddString (szTemp1);
   _stprintf_s(szTemp1, 128, _T("%s = %d"), (LPCTSTR)COMMTIMEOUTSMembers[i++], m_CurrentCommTimeouts.WriteTotalTimeoutMultiplier);
   m_StructLB.AddString (szTemp1);
   _stprintf_s(szTemp1, 128, _T("%s = %d"), (LPCTSTR)COMMTIMEOUTSMembers[i++], m_CurrentCommTimeouts.WriteTotalTimeoutConstant);
   m_StructLB.AddString (szTemp1);

}  // End of OutputCommTimeouts


int CScale3MainWin::WrtReadPort (HANDLE PortHandle, char *Cmd, OVERLAPPED *OverLap, BYTE *Data, int DataLen)
{

   DWORD                MyError = 0;
   DWORD                WrtBufLen = 0;
   DWORD                NumBytes = 0;
   DWORD                EventMask = 0;
   int                  i, iWrtBufLen, iCnt = 0;
   BOOL                 bResult1;
   TCHAR                szTemp1[STANDARD_STRING];
   BYTE                 aByte;

   WrtBufLen = (DWORD)strlen (Cmd);
   iWrtBufLen = (int)WrtBufLen;
   bResult1 = WriteFile (
                  PortHandle,
                  Cmd,
                  WrtBufLen,
                  &NumBytes,
                  OverLap);

#ifdef ASYNC
   bResult1 = GetOverlappedResult(PortHandle, OverLap, &NumBytes, TRUE);
#endif

   if (!bResult1)
   {
      DisplaySystemError (&MyError);
      _stprintf_s (szTemp1,
                   STANDARD_STRING,
                   _T("%s%d"),
                   _T("WriteFile failed in function WrtReadPort with error number "), MyError);
      m_StatusText1.SetWindowText (szTemp1);
      return (1);
   }  // End of WriteFile failed

   NumBytes = 0;
   iCnt = 0;
   Sleep (WRITE_READ_DELAY);
   bResult1 = WaitCommEvent(PortHandle, &EventMask, OverLap);

#ifdef ASYNC
   bResult1 = GetOverlappedResult(PortHandle, OverLap, &NumBytes, TRUE);
#endif

   if (bResult1)
   {
      if (EventMask == EV_RXFLAG)
      {

         do
         {
            bResult1 = ReadFile (
                           PortHandle,
                           &aByte,
                           1,
                           &NumBytes,
                           OverLap);

#ifdef ASYNC
            bResult1 = GetOverlappedResult(PortHandle, OverLap, &NumBytes, TRUE);
#endif

            if (!bResult1)
            {
               DisplaySystemError (&MyError);
               _stprintf_s (szTemp1,
                            256,
                            _T("ReadFile failed in function WrtReadPort at Count %d with error number %d"),
                            iCnt, MyError);
               m_StatusText1.SetWindowText (szTemp1);
               return (2);
            }  // End of ReadFile failed

            if (NumBytes == 1)
               Data[iCnt++] = aByte;

         }
         while ((NumBytes == 1) && (iCnt < DataLen));
      }  // End of if (EventMask == EV_RXFLAG)


#ifdef BS2PIN16
      for (i = 0; (i < iWrtBufLen) && ((i + iWrtBufLen) < iCnt); i++)
         Data[i] = Data[i + WrtBufLen];
      iCnt -= iWrtBufLen;
#endif

      // I want to make sure I NULL terminate the data representation.
      if (iCnt >= 2)
         Data[iCnt - 2] = '\0';
      else if (iCnt >= 0)
         Data[iCnt] = '\0';

      return (0);

   }  // End of WaitCommEvent success and no overlap

   else
   {
      DisplaySystemError (&MyError);
      _stprintf_s (szTemp1,
                   STANDARD_STRING,
                   _T("WaitCommEvent failed in function WrtReadPort error number %d"), MyError);
      m_StatusText1.SetWindowText (szTemp1);
      return (3);
   }  // End of WaitCommEvent failure due to overlapped

}  // End of WrtReadPort


int CScale3MainWin::DoAllOfTheAboveStep1 (void)
{

   // Step1 - CreateFile

   DWORD                MyError = 0;
   TCHAR                szTemp1[STANDARD_STRING];

#ifdef ASYNC
   m_hCurrentHandle = CreateFile (
                           m_ComPorts[m_sCurrentPortIndex].szDevicePath,
                           GENERIC_READ | GENERIC_WRITE,
                           0,       // must be opened with exclusive-access
                           NULL,    // default security attributes
                           OPEN_EXISTING, // must use OPEN_EXISTING
                           FILE_FLAG_OVERLAPPED,  // asynchronous I/O
                           NULL);   // hTemplate must be NULL for comm devices
#else
   m_hCurrentHandle = CreateFile(
                           m_ComPorts[m_sCurrentPortIndex].szDevicePath,
                           GENERIC_READ | GENERIC_WRITE,
                           0,       // must be opened with exclusive-access
                           NULL,    // default security attributes
                           OPEN_EXISTING, // must use OPEN_EXISTING
                           0,       // synchronous I/O
                           NULL);   // hTemplate must be NULL for comm devices
#endif

   if (m_hCurrentHandle == INVALID_HANDLE_VALUE) 
   {
      DisplaySystemError (&MyError);
      _stprintf_s (szTemp1,
                   STANDARD_STRING,
                   _T("%hs%d%hs"),
                   "CreateFile failed in DoAllOfTheAboveStep1 with error number ", MyError,
                   ", the returned File Handle = INVALID_HANDLE_VALUE");
      m_StatusText1.SetWindowText (szTemp1);

      m_hCurrentHandle = (HANDLE)NULL;
      m_sHandleFlag = 0;
      return (0);
   }

   m_sHandleFlag = 1;
   return (1);

}  // End of DoAllOfTheAboveStep1


int CScale3MainWin::DoAllOfTheAboveStep2 (void)
{

   // Step2 - BuildCommDCB

   DWORD                MyError = 0;
   int                  iTemp1;
   BOOL                 bResult1;
   TCHAR                szTemp1[STANDARD_STRING];

   //SecureZeroMemory (&m_dcbCurrentDCB, sizeof (DCB));
   //m_dcbCurrentDCB.DCBlength = sizeof (DCB);

   iTemp1 = (int)fStopBits[m_sCurrentStopBitsIndex];
   //sprintf_s (szTemp1, STANDARD_STRING, _T("baud=%d parity=%s data=%d stop=%d"),
      //ulBaudRates[m_sCurrentBaudRateIndex],
      //szParity[m_sCurrentParityIndex],
      //sDataBits[m_sCurrentDataBitsIndex],
      //iTemp1);
   _stprintf_s(szTemp1, STANDARD_STRING, _T("%d,%s,%d,%d"),
      ulBaudRates[m_sCurrentBaudRateIndex],
      (LPCTSTR)szParity2[m_sCurrentParityIndex],
      sDataBits[m_sCurrentDataBitsIndex],
      iTemp1);
   bResult1 = BuildCommDCB (szTemp1, &m_dcbCurrentDCB);
   //bResult1 = BuildCommDCB ("2400,e,7,1", &m_dcbCurrentDCB);
   //m_dcbCurrentDCB.BaudRate = ulBaudRates[m_sCurrentBaudRateIndex];
   //m_dcbCurrentDCB.ByteSize = (BYTE)sDataBits[m_sCurrentDataBitsIndex];
   //m_dcbCurrentDCB.Parity = (BYTE)m_sCurrentParityIndex;
   //m_dcbCurrentDCB.StopBits = (BYTE)m_sCurrentStopBitsIndex;
   //bResult1 = true;

   if (!bResult1)
   {

      DisplaySystemError (&MyError);
      _stprintf_s (szTemp1,
                   STANDARD_STRING,
                   _T("%s%d"),
                   _T("BuildCommDCB in DoAllOfTheAboveStep2 failed with error number "), MyError);
      m_StatusText1.SetWindowText (szTemp1);

      m_sDCBBuiltFlag = 0;

      return (0);

   }  // End of BuildCommDCB failed

   m_sDCBBuiltFlag = 1;
   m_dcbCurrentDCB.EvtChar = '\n';
   return (1);

}  // End of DoAllOfTheAboveStep2


int CScale3MainWin::DoAllOfTheAboveStep3 (void)
{

   // Step3 - SetCommState

   DWORD                MyError = 0;
   BOOL                 bResult1;
   TCHAR                szTemp1[STANDARD_STRING];

   bResult1 = SetCommState (m_hCurrentHandle, &m_dcbCurrentDCB);
   if (!bResult1)
   {

      DisplaySystemError (&MyError);
      _stprintf_s (szTemp1,
                   STANDARD_STRING,
                   _T("%s%d"),
                   _T("SetCommState in DoAllOfTheAboveStep3 failed with error number "), MyError);
      m_StatusText1.SetWindowText (szTemp1);

      return (0);

   }  // End of SetCommState failed

   return (1);

}  // End of DoAllOfTheAboveStep3


int CScale3MainWin::DoAllOfTheAboveStep4 (void)
{

   // Step4 - SetCommMask

   DWORD                MyError = 0;
   BOOL                 bResult1;
   TCHAR                szTemp1[256];

   bResult1 = SetCommMask (m_hCurrentHandle, EV_RXFLAG);// | EV_ERR | EV_BREAK);
   if (!bResult1)
   {

      DisplaySystemError (&MyError);
      _stprintf_s (szTemp1,
                   STANDARD_STRING,
                   _T("%s%d"),
                   _T("SetCommMask in DoAllOfTheAboveStep4 failed with error number "), MyError);
      m_StatusText1.SetWindowText (szTemp1);

      return (0);

   }  // End of SetCommMask failed

   return (1);

}  // End of DoAllOfTheAboveStep4


int CScale3MainWin::DoAllOfTheAboveStep5 (void)
{

   // Step5 - member OVERLAPPED struct has its event created and is initialized.

   DWORD                MyError = 0;
   TCHAR                szTemp1[STANDARD_STRING];

   m_oCurrentOverlappedStruct.hEvent = NULL;
   m_oCurrentOverlappedStruct.hEvent = CreateEvent (
                                          NULL,
                                          true,
                                          false,
                                          NULL);
   m_oCurrentOverlappedStruct.Internal = 0;
   m_oCurrentOverlappedStruct.InternalHigh = 0;
   m_oCurrentOverlappedStruct.Offset = 0;
   m_oCurrentOverlappedStruct.OffsetHigh = 0;

   if (m_oCurrentOverlappedStruct.hEvent == NULL)
   {
      DisplaySystemError (&MyError);
      _stprintf_s (szTemp1,
                   STANDARD_STRING,
                   _T("The OVERLAPPED struct intialization failed in DoAllOfTheAboveStep5 with error %d."),
                   MyError);
      m_StatusText1.SetWindowText (szTemp1);
      m_sCommStateSet = 0;
      return (0);
   }

   m_sCommStateSet = 1;
   return (1);

}  // End of DoAllOfTheAboveStep5


int CScale3MainWin::DoAllOfTheAboveStep6 (void)
{

   // Step6 - SetCommTimeouts

   DWORD                MyError = 0;
   BOOL                 bResult1;
   TCHAR                szTemp1[STANDARD_STRING];

   m_CurrentCommTimeouts.ReadIntervalTimeout = MAXDWORD;
   m_CurrentCommTimeouts.ReadTotalTimeoutMultiplier = 0;
   m_CurrentCommTimeouts.ReadTotalTimeoutConstant = 0;
   m_CurrentCommTimeouts.WriteTotalTimeoutMultiplier = 10;
   m_CurrentCommTimeouts.WriteTotalTimeoutConstant = 1000;
   bResult1 = SetCommTimeouts (m_hCurrentHandle, &m_CurrentCommTimeouts);

   if (!bResult1)
   {
      DisplaySystemError (&MyError);
      _stprintf_s (szTemp1,
                   STANDARD_STRING,
                   _T("%s%d"),
                   _T("SetCommTimeouts failed in DoAllOfTheAboveStep6 with error number "), MyError);
      m_StatusText1.SetWindowText (szTemp1);
      m_sTimeoutsSet = 0;
      return (0);
   }

   m_sTimeoutsSet = 1;
   return (1);

}  // End of DoAllOfTheAboveStep6


void CScale3MainWin::DisplaySystemError (DWORD *TheError)
{

   DWORD MyError = 0;
   LPTSTR lpMsgBuf = (LPTSTR)NULL;

   MyError = GetLastError ();
   *TheError = MyError;
   FormatMessage ( 
      FORMAT_MESSAGE_ALLOCATE_BUFFER | 
         FORMAT_MESSAGE_FROM_SYSTEM | 
            FORMAT_MESSAGE_IGNORE_INSERTS,
      NULL,
      MyError,
      MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), // Default language
      (LPTSTR)&lpMsgBuf,
      0,
      NULL);

   m_StatusText2.SetWindowText (lpMsgBuf);

   if (lpMsgBuf != (LPTSTR)NULL)
      LocalFree (lpMsgBuf);

}  // End of DisplaySystemError


void CScale3MainWin::ClearStatusLines (void)
{

   m_StatusText1.SetWindowText (_T(""));
   m_StatusText2.SetWindowText (_T(""));

}  // End of DisplaySystemError


UINT Port1Thread (LPVOID pParams)
{

   BYTE                 Data[STANDARD_STRING];
   DWORD                MyError = 0;
   DWORD                WrtBufLen = 0;
   DWORD                NumBytes = 0;
   DWORD                EventMask = 0;
   int                  i, iWrtBufLen, iCnt = 0;
   BOOL                 bResult1;
   BYTE                 aByte;

write:
   EnterCriticalSection(&(tp[0].CritSec));
   WrtBufLen = (DWORD)strlen (tp[0].Cmd);
   iWrtBufLen = (int)WrtBufLen;
   bResult1 = WriteFile (
                  tp[0].hCom,
                  tp[0].Cmd,
                  WrtBufLen,
                  &NumBytes,
                  &(tp[0].oOverlappedStruct));

#ifdef ASYNC
   bResult1 = GetOverlappedResult(tp[0].hCom, &(tp[0].oOverlappedStruct), &NumBytes, TRUE);
#endif

   if (!bResult1)
   {
      GetSystemError (&MyError, tr[0].ErrorMsg2);
      _stprintf_s (tr[0].ErrorMsg1,
                   STANDARD_STRING,
                   _T("%s%d"),
                   _T("WriteFile failed in Port1Thread with error number "), MyError);
      tr[0].Data[0] = '\0';
      goto error;
   }  // End of WriteFile failed

   NumBytes = 0;
   iCnt = 0;
   Sleep (WRITE_READ_DELAY);
   bResult1 = WaitCommEvent(tp[0].hCom, &EventMask, &(tp[0].oOverlappedStruct));

#ifdef ASYNC
   bResult1 = GetOverlappedResult(tp[0].hCom, &(tp[0].oOverlappedStruct), &NumBytes, TRUE);
#endif

   if (bResult1)
   {
      if (EventMask == EV_RXFLAG)
      {

         do
         {
            bResult1 = ReadFile (
                           tp[0].hCom,
                           &aByte,
                           1,
                           &NumBytes,
                           &(tp[0].oOverlappedStruct));

#ifdef ASYNC
            bResult1 = GetOverlappedResult(tp[0].hCom, &(tp[0].oOverlappedStruct), &NumBytes, TRUE);
#endif

            if (!bResult1)
            {
               GetSystemError (&MyError, tr[0].ErrorMsg2);
               _stprintf_s (tr[0].ErrorMsg1,
                            STANDARD_STRING,
                            _T("ReadFile failed in Port1Thread at Count %d with error number %d"),
                            iCnt, MyError);
               tr[0].Data[0] = '\0';
               goto error;
            }  // End of ReadFile failed

            if (NumBytes == 1)
               Data[iCnt++] = aByte;

         }
         while ((NumBytes == 1) && (iCnt < STANDARD_STRING));
      }  // End of if (EventMask == EV_RXFLAG)

#ifdef BS2PIN16
      for (i = 0; (i < iWrtBufLen) && ((i + iWrtBufLen) < iCnt); i++)
         Data[i] = Data[i + WrtBufLen];
      iCnt -= iWrtBufLen;
#endif

      // I want to be sure to NULL terminate the data representation.
      if (iCnt >= 2)
         Data[iCnt - 2] = '\0';
      else if (iCnt >= 0)
         Data[iCnt] = '\0';

      strcpy_s (tr[0].Data, STANDARD_STRING, (char *)Data);
      tr[0].ErrorMsg1[0] = '\0';
      tr[0].ErrorMsg2[0] = '\0';
      ::PostMessage (tp[0].hWnd, WM_USER_PORT1_THREAD_FINISHED_DATA, 0, 0);
      goto suspend;

   }  // End of WaitCommEvent success and no overlap

   else
   {
      GetSystemError (&MyError, tr[0].ErrorMsg2);
      _stprintf_s (tr[0].ErrorMsg1,
                   STANDARD_STRING,
                   _T("WaitCommEvent failed in Port1Thread error number %d"), MyError);
      tr[0].Data[0] = '\0';
      goto error;
   }  // End of WaitCommEvent failure due to overlapped

error:
   ::PostMessage (tp[0].hWnd, WM_USER_PORT1_THREAD_FINISHED_ERRORS, 0, 0);

suspend:
   LeaveCriticalSection(&(tp[0].CritSec));
   SuspendThread (tp[0].hThread);
   if (tp[0].sEndThreadFlag)
      AfxEndThread (0);
   goto write;


}  // End of Port1Thread


UINT Port2Thread (LPVOID pParams)
{

   BYTE                 Data[STANDARD_STRING];
   DWORD                MyError = 0;
   DWORD                WrtBufLen = 0;
   DWORD                NumBytes = 0;
   DWORD                EventMask = 0;
   int                  i, iWrtBufLen, iCnt = 0;
   BOOL                 bResult1;
   BYTE                 aByte;

write:
   EnterCriticalSection(&(tp[1].CritSec));
   WrtBufLen = (DWORD)strlen (tp[1].Cmd);
   iWrtBufLen = (int)WrtBufLen;
   bResult1 = WriteFile (
                  tp[1].hCom,
                  tp[1].Cmd,
                  WrtBufLen,
                  &NumBytes,
                  &(tp[1].oOverlappedStruct));

#ifdef ASYNC
   bResult1 = GetOverlappedResult(tp[1].hCom, &(tp[1].oOverlappedStruct), &NumBytes, TRUE);
#endif

   if (!bResult1)
   {
      GetSystemError (&MyError, tr[1].ErrorMsg2);
      _stprintf_s (tr[1].ErrorMsg1,
                   STANDARD_STRING,
                   _T("%s%d"),
                   _T("WriteFile failed in Port2Thread with error number "), MyError);
      tr[1].Data[0] = '\0';
      goto error;
   }  // End of WriteFile failed

   NumBytes = 0;
   iCnt = 0;
   Sleep (WRITE_READ_DELAY);
   bResult1 = WaitCommEvent(tp[1].hCom, &EventMask, &(tp[1].oOverlappedStruct));

#ifdef ASYNC
   bResult1 = GetOverlappedResult(tp[1].hCom, &(tp[1].oOverlappedStruct), &NumBytes, TRUE);
#endif

   if (bResult1)
   {
      if (EventMask == EV_RXFLAG)
      {

         do
         {
            bResult1 = ReadFile (
                           tp[1].hCom,
                           &aByte,
                           1,
                           &NumBytes,
                           &(tp[1].oOverlappedStruct));

#ifdef ASYNC
            bResult1 = GetOverlappedResult(tp[1].hCom, &(tp[1].oOverlappedStruct), &NumBytes, TRUE);
#endif

            if (!bResult1)
            {
               GetSystemError (&MyError, tr[1].ErrorMsg2);
               _stprintf_s (tr[1].ErrorMsg1,
                            STANDARD_STRING,
                            _T("ReadFile failed in Port2Thread at Count %d with error number %d"),
                            iCnt, MyError);
               tr[1].Data[0] = '\0';
               goto error;
            }  // End of ReadFile failed

            if (NumBytes == 1)
               Data[iCnt++] = aByte;

         }
         while ((NumBytes == 1) && (iCnt < STANDARD_STRING));
      }  // End of if (EventMask == EV_RXFLAG)


#ifdef BS2PIN16
      for (i = 0; (i < iWrtBufLen) && ((i + iWrtBufLen) < iCnt); i++)
         Data[i] = Data[i + WrtBufLen];
      iCnt -= iWrtBufLen;
#endif

      // I want to make sure I NULL terminate the data representation.
      if (iCnt >= 2)
         Data[iCnt - 2] = '\0';
      else if (iCnt >= 0)
         Data[iCnt] = '\0';

      strcpy_s (tr[1].Data, STANDARD_STRING, (char *)Data);
      tr[1].ErrorMsg1[0] = '\0';
      tr[1].ErrorMsg2[0] = '\0';
      ::PostMessage (tp[1].hWnd, WM_USER_PORT2_THREAD_FINISHED_DATA, 0, 0);
      goto suspend;

   }  // End of WaitCommEvent success and no overlap

   else
   {
      GetSystemError (&MyError, tr[1].ErrorMsg2);
      _stprintf_s (tr[1].ErrorMsg1,
                   STANDARD_STRING,
                   _T("WaitCommEvent failed in Port2Thread error number %d"), MyError);
      tr[1].Data[0] = '\0';
      goto error;
   }  // End of WaitCommEvent failure due to overlapped

error:
   ::PostMessage (tp[1].hWnd, WM_USER_PORT2_THREAD_FINISHED_ERRORS, 0, 0);

suspend:
   LeaveCriticalSection(&(tp[1].CritSec));
   SuspendThread (tp[1].hThread);
   if (tp[1].sEndThreadFlag)
      AfxEndThread (0);
   goto write;


}  // End of Port2Thread


UINT Port3Thread (LPVOID pParams)
{

   BYTE                 Data[STANDARD_STRING];
   DWORD                MyError = 0;
   DWORD                WrtBufLen = 0;
   DWORD                NumBytes = 0;
   DWORD                EventMask = 0;
   int                  i, iWrtBufLen, iCnt = 0;
   BOOL                 bResult1;
   BYTE                 aByte;

write:
   EnterCriticalSection(&(tp[2].CritSec));
   WrtBufLen = (DWORD)strlen (tp[2].Cmd);
   iWrtBufLen = (int)WrtBufLen;
   bResult1 = WriteFile (
                  tp[2].hCom,
                  tp[2].Cmd,
                  WrtBufLen,
                  &NumBytes,
                  &(tp[2].oOverlappedStruct));

#ifdef ASYNC
   bResult1 = GetOverlappedResult(tp[2].hCom, &(tp[2].oOverlappedStruct), &NumBytes, TRUE);
#endif

   if (!bResult1)
   {
      GetSystemError (&MyError, tr[2].ErrorMsg2);
      _stprintf_s (tr[2].ErrorMsg1,
                   STANDARD_STRING,
                   _T("%s%d"), _T("WriteFile failed in Port3Thread with error number "), MyError);
      tr[2].Data[0] = '\0';
      goto error;
   }  // End of WriteFile failed

   NumBytes = 0;
   iCnt = 0;
   Sleep (WRITE_READ_DELAY);
   bResult1 = WaitCommEvent(tp[2].hCom, &EventMask, &(tp[2].oOverlappedStruct));

#ifdef ASYNC
   bResult1 = GetOverlappedResult(tp[2].hCom, &(tp[2].oOverlappedStruct), &NumBytes, TRUE);
#endif

   if (bResult1)
   {
      if (EventMask == EV_RXFLAG)
      {

         do
         {
            bResult1 = ReadFile (
                           tp[2].hCom,
                           &aByte,
                           1,
                           &NumBytes,
                           &(tp[2].oOverlappedStruct));

#ifdef ASYNC
            bResult1 = GetOverlappedResult(tp[2].hCom, &(tp[2].oOverlappedStruct), &NumBytes, TRUE);
#endif

            if (!bResult1)
            {
               GetSystemError (&MyError, tr[2].ErrorMsg2);
               _stprintf_s (tr[2].ErrorMsg1,
                            STANDARD_STRING,
                            _T("ReadFile failed in Port3Thread at Count %d with error number %d"),
                            iCnt, MyError);
               tr[2].Data[0] = '\0';
               goto error;
            }  // End of ReadFile failed

            if (NumBytes == 1)
               Data[iCnt++] = aByte;

         }
         while ((NumBytes == 1) && (iCnt < STANDARD_STRING));
      }  // End of if (EventMask == EV_RXFLAG)


#ifdef BS2PIN16
      for (i = 0; (i < iWrtBufLen) && ((i + iWrtBufLen) < iCnt); i++)
         Data[i] = Data[i + WrtBufLen];
      iCnt -= iWrtBufLen;
#endif

      // I want to make sure I NULL terminate the data representation.
      if (iCnt >= 2)
         Data[iCnt - 2] = '\0';
      else if (iCnt >= 0)
         Data[iCnt] = '\0';

      strcpy_s (tr[2].Data, STANDARD_STRING, (char *)Data);
      tr[2].ErrorMsg1[0] = '\0';
      tr[2].ErrorMsg2[0] = '\0';
      ::PostMessage (tp[2].hWnd, WM_USER_PORT3_THREAD_FINISHED_DATA, 0, 0);
      goto suspend;

   }  // End of WaitCommEvent success and no overlap

   else
   {
      GetSystemError (&MyError, tr[2].ErrorMsg2);
      _stprintf_s (tr[2].ErrorMsg1,
                   STANDARD_STRING,
                   _T("WaitCommEvent failed in Port3Thread error number %d"), MyError);
      tr[2].Data[0] = '\0';
      goto error;
   }  // End of WaitCommEvent failure due to overlapped

error:
   ::PostMessage (tp[2].hWnd, WM_USER_PORT3_THREAD_FINISHED_ERRORS, 0, 0);

suspend:
   LeaveCriticalSection(&(tp[2].CritSec));
   SuspendThread (tp[2].hThread);
   if (tp[2].sEndThreadFlag)
      AfxEndThread (0);
   goto write;


}  // End of Port3Thread


UINT Port4Thread (LPVOID pParams)
{

   BYTE                 Data[STANDARD_STRING];
   DWORD                MyError = 0;
   DWORD                WrtBufLen = 0;
   DWORD                NumBytes = 0;
   DWORD                EventMask = 0;
   int                  i, iWrtBufLen, iCnt = 0;
   BOOL                 bResult1;
   BYTE                 aByte;

write:
   EnterCriticalSection (&(tp[3].CritSec));
   WrtBufLen = (DWORD)strlen (tp[3].Cmd);
   iWrtBufLen = (int)WrtBufLen;
   bResult1 = WriteFile (
                  tp[3].hCom,
                  tp[3].Cmd,
                  WrtBufLen,
                  &NumBytes,
                  &(tp[3].oOverlappedStruct));

#ifdef ASYNC
   bResult1 = GetOverlappedResult(tp[3].hCom, &(tp[3].oOverlappedStruct), &NumBytes, TRUE);
#endif

   if (!bResult1)
   {
      GetSystemError (&MyError, tr[3].ErrorMsg2);
      _stprintf_s (tr[3].ErrorMsg1,
                   STANDARD_STRING,
                   _T("%s%d"),
                   _T("WriteFile failed in Port4Thread with error number "), MyError);
      tr[3].Data[0] = '\0';
      goto error;
   }  // End of WriteFile failed

   NumBytes = 0;
   iCnt = 0;
   Sleep (WRITE_READ_DELAY);
   bResult1 = WaitCommEvent (tp[3].hCom, &EventMask, &(tp[3].oOverlappedStruct));

#ifdef ASYNC
   bResult1 = GetOverlappedResult(tp[3].hCom, &(tp[3].oOverlappedStruct), &NumBytes, TRUE);
#endif

   if (bResult1)
   {
      if (EventMask == EV_RXFLAG)
      {

         do
         {
            bResult1 = ReadFile (
                           tp[3].hCom,
                           &aByte,
                           1,
                           &NumBytes,
                           &(tp[3].oOverlappedStruct));

#ifdef ASYNC
            bResult1 = GetOverlappedResult(tp[3].hCom, &(tp[3].oOverlappedStruct), &NumBytes, TRUE);
#endif

            if (!bResult1)
            {
               GetSystemError (&MyError, tr[3].ErrorMsg2);
               _stprintf_s (tr[3].ErrorMsg1,
                            STANDARD_STRING,
                            _T("ReadFile failed in Port4Thread at Count %d with error number %d"),
                            iCnt, MyError);
               tr[3].Data[0] = '\0';
               goto error;
            }  // End of ReadFile failed

            if (NumBytes == 1)
               Data[iCnt++] = aByte;

         }
         while ((NumBytes == 1) && (iCnt < STANDARD_STRING));
      }  // End of if (EventMask == EV_RXFLAG)


#ifdef BS2PIN16
      for (i = 0; (i < iWrtBufLen) && ((i + iWrtBufLen) < iCnt); i++)
         Data[i] = Data[i + WrtBufLen];
      iCnt -= iWrtBufLen;
#endif

      // I want to make sure I NULL terminate the data representation.
      if (iCnt >= 2)
         Data[iCnt - 2] = '\0';
      else if (iCnt >= 0)
         Data[iCnt] = '\0';

      strcpy_s (tr[3].Data, STANDARD_STRING, (char *)Data);
      tr[3].ErrorMsg1[0] = '\0';
      tr[3].ErrorMsg2[0] = '\0';
      ::PostMessage (tp[3].hWnd, WM_USER_PORT4_THREAD_FINISHED_DATA, 0, 0);
      goto suspend;

   }  // End of WaitCommEvent success and no overlap

   else
   {
      GetSystemError (&MyError, tr[3].ErrorMsg2);
      _stprintf_s (tr[3].ErrorMsg1,
                   STANDARD_STRING,
                   _T("WaitCommEvent failed in Port4Thread error number %d"), MyError);
      tr[3].Data[0] = '\0';
      goto error;
   }  // End of WaitCommEvent failure due to overlapped

error:
   ::PostMessage (tp[3].hWnd, WM_USER_PORT4_THREAD_FINISHED_ERRORS, 0, 0);

suspend:
   LeaveCriticalSection (&(tp[3].CritSec));
   SuspendThread (tp[3].hThread);
   if (tp[3].sEndThreadFlag)
      AfxEndThread (0);
   goto write;


}  // End of Port4Thread


void GetSystemError (DWORD *TheError, TCHAR *TheErrorString)
{

   DWORD MyError = 0;
   LPTSTR lpMsgBuf = (LPTSTR)NULL;

   MyError = GetLastError ();
   *TheError = MyError;
   FormatMessage ( 
      FORMAT_MESSAGE_ALLOCATE_BUFFER | 
         FORMAT_MESSAGE_FROM_SYSTEM | 
            FORMAT_MESSAGE_IGNORE_INSERTS,
      NULL,
      MyError,
      MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), // Default language
      (LPTSTR)&lpMsgBuf,
      0,
      NULL);

   _tcscpy_s (TheErrorString, STANDARD_STRING, lpMsgBuf);

   if (lpMsgBuf != (LPTSTR)NULL)
      LocalFree (lpMsgBuf);

}  // End of GetSystemError




/***************** Post Notes *****************/
/***********************************************

1) COMMPROP structure for the StarTech ICUSB2324X USB to RS-232 adapter

typedef struct _COMMPROP {
   WORD wPacketLength;
   WORD wPacketVersion;
   DWORD dwServiceMask;
   DWORD dwReserved1;
   DWORD dwMaxTxQueue;
   DWORD dwMaxRxQueue;
   DWORD dwMaxBaud;
   DWORD dwProvSubType;
   DWORD dwProvCapabilities;
   DWORD dwSettableParams;
   DWORD dwSettableBaud;
   WORD wSettableData;
   WORD wSettableStopParity;
   DWORD dwCurrentTxQueue;
   DWORD dwCurrentRxQueue;
   DWORD dwProvSpec1;
   DWORD dwProvSpec2;
   WCHAR wcProvChar[1];
} COMMPROP;

wPacketLength = 64
wPacketVersion = 2
dwServiceMask = 1
dwReserved1 = 0
dwMaxTxQueue = 0              Maximum size of the driver's internal output buffer,
                              in bytes. A value of zero indicates that no
                              maximum value is imposed by the serial provider.

dwMaxRxQueue = 0              Maximum size of the driver's internal input buffer,
                              in bytes. A value of zero indicates that no
                              maximum value is imposed by the serial provider.

dwMaxBaud = 0x10000000        Maximum allowable baud rate, in bits per second (bps).
                              0x10000000 = BAUD_USER, Programmable baud rate.

dwProvSubType = 1             Communications-provider type.
                              0x00000001 = PST_RS232, RS-232 serial port

dwProvCapabilities = 0xff     Capabilities offered by the provider.
                              This member can be one of the following values.
                              Guess it can also be a combination.
                              0x0001 = PCF_DTRDSR, DTR (data-terminal-ready)/DSR (data-set-ready) supported
                              0x0080 = PCF_INTTIMEOUTS, Interval time-outs supported
                              0x0008 = PCF_PARITY_CHECK, Parity checking supported
                              0x0004 = PCF_RLSD, RLSD (receive-line-signal-detect) supported
                              0x0002 = PCF_RTSCTS, RTS (request-to-send)/CTS (clear-to-send) supported
                              0x0020 = PCF_SETXCHAR, Settable XON/XOFF supported
                              0x0040 = PCF_TOTALTIMEOUTS, Total (elapsed) time-outs supported
                              0x0010 = PCF_XONXOFF, XON/XOFF flow control supported
                              ______
                              0x00ff

dwSettableParams = 0x7f       Communications parameter that can be changed.
                              This member can be one of the following values.
                              Guess it can also be a combination.
                              0x0002 = SP_BAUD, Baud rate
                              0x0004 = SP_DATABITS, Data bits
                              0x0010 = SP_HANDSHAKING, Handshaking (flow control)
                              0x0001 = SP_PARITY, Parity
                              0x0020 = SP_PARITY_CHECK, Parity checking
                              0x0040 = SP_RLSD, RLSD (receive-line-signal-detect)
                              0x0008 = SP_STOPBITS, Stop bits
                              ______
                              0x007f

dwSettableBaud = 0x10000000   Baud rates that can be used.
                              For values, see the dwMaxBaud member. 

wSettableData = 0xf           Number of data bits that can be set.
                              This member can be one of the following values. 
                              Guess it can also be a combination.
                              0x0001 = DATABITS_5, 5 data bits
                              0x0002 = DATABITS_6, 6 data bits
                              0x0004 = DATABITS_7, 7 data bits
                              0x0008 = DATABITS_8, 8 data bits
                              ______
                              0x000f

wSettableStopParity = 0x1f07  Stop bit and parity settings that can be selected.
                              This member can be one of the following values.
                              Guess it can also be a combination.
                              0x0001 = STOPBITS_10, 1 stop bit
                              0x0002 = STOPBITS_15, 1.5 stop bits
                              0x0004 = STOPBITS_20, 2 stop bits
                              0x0100 = PARITY_NONE, no parity
                              0x0200 = PARITY_ODD, odd parity
                              0x0400 = PARITY_EVEN, even parity
                              0x0800 = PARITY_MARK, mark parity
                              0x1000 = PARITY_SPACE, space parity
                              ______
                              0x1f07

dwCurrentTxQueue = 0          Size of the driver's internal output buffer, in bytes.
                              A value of zero indicates that the value is unavailable. 

dwCurrentRxQueue = 4096       Size of the driver's internal input buffer, in bytes.
                              A value of zero indicates that the value is unavailable.

dwProvSpec1 = 0               Provider-specific data.
                              Applications should ignore this member unless they have
                              detailed information about the format of the data
                              required by the provider. 
                              Set this member to COMMPROP_INITIALIZED before calling
                              the GetCommProperties function to indicate that the
                              wPacketLength member is already valid.

dwProvSpec2 = 0               Provider-specific data.
                              Applications should ignore this member unless they have
                              detailed information about the format of the data
                              required by the provider.

wcProvChar[1] = ""            Provider-specific data.
                              Applications should ignore this member unless they have
                              detailed information about the format of the data
                              required by the provider.


***********************************************/
/**********************************************/



// End of module CScale3MainWin.CPP