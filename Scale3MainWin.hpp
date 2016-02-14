

/***************************************************************


System:     BioWindow -- Data Acquisition, Analysis, and Report Generator


File:       Scale3MainWin.HPP


Procedure:  Defines all of the methods and members contained in the
            Scale3MainWin class.


Programmer: James D. Matthews


Date:       Oct. 1, 2002



Copyright (C) 2002, Modular Instruments, Inc.

=================================================================

                        Revision Block
                        --------------

   ECO No.     Date     Init.    Procedure and Description of Change
   --------    -------- -----    -----------------------------------


*****************************************************************/


#if !defined (_SCALE3MAINWIN_HPP)
#define _SCALE3MAINWIN_HPP

#define STANDARD_BUFFER                256
#define STANDARD_STRING                256
#define WRITE_READ_DELAY               400
#define THREADTIMER                    256
#define DEFAULTTHREADTIMERVALUE        5000

//#include <fcntl.h>      /* Needed only for _O_RDWR definition */
//#include <io.h>
//#include <share.h>
//#include <stdlib.h>
//#include <stdio.h>

#include "resource.h"
#include "EnumComPorts.h"

#define WM_USER_PORT_THREAD_FINISHED_DATA          WM_USER + 0x100
#define WM_USER_PORT_THREAD_FINISHED_ERRORS        WM_USER + 0x101
#define WM_USER_PORT1_THREAD_FINISHED_DATA         WM_USER + 0x102
#define WM_USER_PORT2_THREAD_FINISHED_DATA         WM_USER + 0x103
#define WM_USER_PORT3_THREAD_FINISHED_DATA         WM_USER + 0x104
#define WM_USER_PORT4_THREAD_FINISHED_DATA         WM_USER + 0x105
#define WM_USER_PORT1_THREAD_FINISHED_ERRORS       WM_USER + 0x106
#define WM_USER_PORT2_THREAD_FINISHED_ERRORS       WM_USER + 0x107
#define WM_USER_PORT3_THREAD_FINISHED_ERRORS       WM_USER + 0x108
#define WM_USER_PORT4_THREAD_FINISHED_ERRORS       WM_USER + 0x109

UINT Port1Thread (LPVOID pParams);
UINT Port2Thread (LPVOID pParams);
UINT Port3Thread (LPVOID pParams);
UINT Port4Thread (LPVOID pParams);
void GetSystemError (DWORD *TheError, TCHAR *TheErrorString);

/*******************************************************
Unicode issue!

The question is weither the serial device is Unicode or ANSI. Does the device expect
Unicode or ANSI communication? By defining Cmd as TCHAR you're leaving that decission
up to the compilier. While I respect Microsoft's VS C++ compilier, I do not attribute
psycic powers to it. So this assumption, like most assumptions, must be incorrect.

Since the microcontoller in the device is programmed by another entity, how can we
provide general purpose communications?

*******************************************************/

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

typedef struct _THREADPARAMS
{
   HWND              hWnd;
   HANDLE            hCom;
   HANDLE            hThread;
   char              Cmd[STANDARD_BUFFER]; // What is sent to the device.
   OVERLAPPED        oOverlappedStruct;
   short             sEndThreadFlag;
   CRITICAL_SECTION  CritSec;
} THREADPARAMS;

typedef struct _THREADRETURNS
{
   TCHAR             ErrorMsg1[STANDARD_STRING]; // Used by the program not the device.
   TCHAR             ErrorMsg2[STANDARD_STRING]; // Used by the program not the device.
   char              Data[STANDARD_STRING]; // What is received from the device
} THREADRETURNS;

/*Static globals used between threads and the GUI. Better be sure one can't hit them while the other is
using them, (you can never be sure). Not using thread synchonization of some type, CriticalSection, Mutex,
can only end in segmentation faults, locked-up boxes, and dismal failure.
Is that true? Doesn't the design of seperate globals for each thread institute a sort of thread to thread
synchronization primitive?
This design insures that one thread doesn't try to mess with another thread's memory space. I admit that
this can result in wasteful memory use, (multiple iterations of the same information such as 
the parent's handle or the serial device's command). It seems to me that the only threat of collison can
occur between the parent process and a thread.
I would argue that the current program design insures that the filling of THREADPARAMS by the parent
and the use of THREADPARAMS by a thread occur at different times and will not overlap. Also, the thread only
ever reads THREADPARAMS. The thread never writes into it. I seem to recall that reading a memory segment while
another process writes into it is somehow less egregious than two prcesses trying to write into the same
memory segment at the sametime.
Am I on thin ice with the use of THREADRETURNS? Again I would argue that program design prohibits the parent
from accessing THREADRETURNS at the same time that the thread is using it, (and vice-versa). I believe this
because the thread uses PostMessage after filling THREADRETURNS and then suspends itself. Also when you
take into account the low throughput and general lathargic requirements of the program, collisons between
parent and thread seem improbable.
I'm not suggesting that thread synchronization isn't a prime objective. It is a requirement anytime
multithreading is used. However, for this not ready for primetime cowboy coding prototype I feel that I can
get away with ignoring it. Is bad programming ever justified? NO, of course not. However, its easier to talk
about it then to impliment it. I recognize my indolent nature and by this time admit that I could have
implemented thread synchronization of somesort instead of trying to pointlessly justifiy bad programming.
OK, I added a CRITICAL_SECTION type called CritSec to the THREADPARAMS structure.
I InitializeCriticalSectionAndSpinCount in the CScale3MainWin constructor and DeleteCriticalSection in the
destructor. At the top of the thread, I have the threads EnterCriticalSection. I have the threads
LeaveCriticalSection immediately before they suspend themselves. I have the parent EnterCriticalSection
at the top of the functions that handle the thread messages and the parent does LeaveCriticalSection
at the end of those functions.*/

static THREADPARAMS           tp[MAX_NUM_COMPORT];
static THREADRETURNS          tr[MAX_NUM_COMPORT];


class CScale3MainWin : public CFrameWnd, public CComPorts
{

public:
   CScale3MainWin ();

private:

   ~CScale3MainWin ();

   // Over Rides

   afx_msg int OnCreate (LPCREATESTRUCT lpcs);
   afx_msg void OnClose (void);
   afx_msg void OnSize (UINT nType, int cx, int cy);
   afx_msg void OnPaint (void);
   afx_msg void OnTimer (UINT nIDEvent);

   virtual void PostNcDestroy ();

   // User Messages
   long OnPort1ThreadFinishedData (WPARAM wParam, LPARAM lParam);
   long OnPort2ThreadFinishedData (WPARAM wParam, LPARAM lParam);
   long OnPort3ThreadFinishedData (WPARAM wParam, LPARAM lParam);
   long OnPort4ThreadFinishedData (WPARAM wParam, LPARAM lParam);
   long OnPort1ThreadFinishedErrors (WPARAM wParam, LPARAM lParam);
   long OnPort2ThreadFinishedErrors (WPARAM wParam, LPARAM lParam);
   long OnPort3ThreadFinishedErrors (WPARAM wParam, LPARAM lParam);
   long OnPort4ThreadFinishedErrors (WPARAM wParam, LPARAM lParam);

   // Menu Commands

   void OnUpdateOpening1ShutDown (CCmdUI *pCmdUI);
   void OnOpening1ShutDown (void);
   void OnUpdateOpening1CreateFile (CCmdUI *pCmdUI);
   void OnOpening1CreateFile (void);
   void OnUpdateOpening1GetCommProps (CCmdUI *pCmdUI);
   void OnOpening1GetCommProps (void);
   void OnUpdateOpening1GetCommState (CCmdUI *pCmdUI);
   void OnOpening1GetCommState (void);
   void OnUpdateOpening1SetCommState (CCmdUI *pCmdUI);
   void OnOpening1SetCommState (void);
   void OnUpdateOpening1GetCommTimeouts (CCmdUI *pCmdUI);
   void OnOpening1GetCommTimeouts (void);
   void OnUpdateOpening1SetCommTimeouts (CCmdUI *pCmdUI);
   void OnOpening1SetCommTimeouts (void);
   void OnUpdateOpening1DoAllOfTheAbove (CCmdUI *pCmdUI);
   void OnOpening1DoAllOfTheAbove (void);
   void OnUpdateOpening1CloseHandle (CCmdUI *pCmdUI);
   void OnOpening1CloseHandle (void);
   void OnUpdateOpening1RdWrtCurrent (CCmdUI *pCmdUI);
   void OnOpening1RdWrtCurrent (void);
   void OnUpdateOpening1RdWrtConfigured (CCmdUI *pCmdUI);
   void OnOpening1RdWrtConfigured (void);
   void OnSelchangePortList (void);
   void OnCloseUpBaudRate (void);
   void OnCloseUpDataBits (void);
   void OnCloseUpStopBits (void);
   void OnCloseUpParity (void);
   void OnClickedStartThreads (void);
   void OnClickedUseThreads (void);
   void OnClickedTimerUseThreads (void);
   void OnClickedEndThreads (void);
   void OnClickedTest (void);

   // Routines

   void FillComPortLB ();
   void SetComboBoxes (void);
   void UpDateComboBoxes (void);
   void OutputProperties (void);
   void OutputCommState (void);
   void OutputCommTimeouts (void);
   int WrtReadPort (HANDLE PortHandle, char *Cmd, OVERLAPPED *OverLap, BYTE *Data, int DataLen);
   int DoAllOfTheAboveStep1 (void);
   int DoAllOfTheAboveStep2 (void);
   int DoAllOfTheAboveStep3 (void);
   int DoAllOfTheAboveStep4 (void);
   int DoAllOfTheAboveStep5 (void);
   int DoAllOfTheAboveStep6 (void);
   void DisplaySystemError (DWORD *TheError);
   void ClearStatusLines (void);
   void EndThreads (void);

private:
   CString strMyClass;

// Opening Menus
   CMenu m_menuOpeningMenu1;

// MI2 Icon Stuff
   HICON m_MI2Icon1;
   HICON m_MI2Icon2;

// We'll use these guys to keep track of the frame window's size as
// we receive WM_SIZE messages. Important for re-painting.
   CRect m_rect;
   int m_iClientHeight, m_iClientWidth;


private:

   UINT_PTR       m_ThreadTimer;
   int            m_cxChar, m_cyChar, m_iThreadTimerValue;
   short          m_sCurrentPortIndex, m_sCurrentBaudRateIndex,
                  m_sCurrentParityIndex, m_sCurrentDataBitsIndex,
                  m_sCurrentStopBitsIndex, m_sHandleFlag, m_sGotCommState,
                  m_sDCBBuiltFlag, m_sCommStateSet, m_sGotTimeouts,
                  m_sTimeoutsSet, m_sPortsConfigured, m_sShowConfig,
                  m_sShowStruct, m_sThread1Running, m_sThread2Running,
                  m_sThread3Running, m_sThread4Running;
   DCB            m_dcbCurrentDCB;
   COMMPROP       m_CurrentCommProp;
   COMMTIMEOUTS   m_CurrentCommTimeouts;
   HANDLE         m_hCurrentHandle;
   OVERLAPPED     m_oCurrentOverlappedStruct;
   CWinThread     *m_Port1ThreadHandle, *m_Port2ThreadHandle,
                  *m_Port3ThreadHandle, *m_Port4ThreadHandle;

   CFont m_fontMain;

   CStatic m_ComPortLBTitle;
   CListBox m_ComPortLB;

   CStatic m_BaudRateCBTitle;
   CComboBox m_BaudRateCB;
   CStatic m_DataBitsCBTitle;
   CComboBox m_DataBitsCB;
   CStatic m_StopBitsCBTitle;
   CComboBox m_StopBitsCB;
   CStatic m_ParityCBTitle;
   CComboBox m_ParityCB;

   CStatic m_StatusText1;
   CStatic m_StatusText2;

   CStatic m_StructLBTitle;
   CListBox m_StructLB;

   CStatic m_ConfiguredLBTitle;
   CListBox m_ConfiguredLB;

   CButton  m_StartThreadsPB;
   CButton  m_UseThreadsPB;
   CButton  m_TimerUseThreadsPB;
   CButton  m_EndThreadsPB;

   CStatic m_ThreadTimerEBTitle;
   CEdit m_ThreadTimerEB;

   CButton  m_TestPB;

   DECLARE_MESSAGE_MAP ();

};  // End of class CScale3MainWin


#endif

