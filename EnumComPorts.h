


/***************************************************************


System:     BioRep -- Data Acquisition, Analysis, and Report Generator


File:       EnumComPorts.HPP


Procedure:  Defines all of the methods and members contained in the
            CComPorts class.


Programmer: James D. Matthews


Date:       Oct. 1, 2002



Copyright (C) 2002, Modular Instruments, Inc.

=================================================================

                        Revision Block
                        --------------

   ECO No.     Date     Init.    Procedure and Description of Change
   --------    -------- -----    -----------------------------------


*****************************************************************/


#if !defined (_EnumComPorts_H)
#define _EnumComPorts_H

#include "tchar.h"
#include <setupapi.h>
//#include <devguid.h>
//#include <regstr.h>
#include <ntddser.h>
//#include <objbase.h>
//#include <initguid.h>


#define COM_NAME_SIZE 32
#define MAX_NUM_COMPORT 4
#define MAX_KEY_LENGTH 256
#define MAX_VALUE_NAME 16384


typedef enum
{
   EMPTY,
   FILLED,
} COM_FLAG;

typedef struct _COM_INFO
{
   COM_FLAG ComFlag;
   bool bUsageFlag;
   TCHAR szDevicePath[MAX_KEY_LENGTH];    // The unintellegable string used in the CreateFile call.
   TCHAR szFriendlyName[64];
   TCHAR szManufacturer[64];
   TCHAR szDeviceDescrp[64];
   TCHAR szDeviceObject[64];
   TCHAR szLocationInformation[64];
   TCHAR szName[32];
   ULONG BaudRate;
   short BaudRateIndex;
   TCHAR Parity[16];
   short ParityIndex;
   short DataBits;
   short DataBitsIndex;
   float StopBits;
   short StopBitsIndex;
   short sDCBSet;
   short sTimeoutsSet;
   HANDLE hCom;
   DCB dcb;
   COMMTIMEOUTS CommTimeouts;
   OVERLAPPED OverLap;
}	COM_INFO;

class CComPorts
{

public:

   CComPorts ();
   ~CComPorts ();

public:
   ULONG m_ulActiveCount;
   int m_iCurrentPort;
   COM_INFO m_ComPorts[MAX_NUM_COMPORT];
   int EnumerateStartechWin8ComPorts ();
   int EnumerateStartechXPComPorts ();
   int AreTherSerialPorts ();
   int AreTherStartechSerialPorts ();

};  // End of class CComPorts

#endif 

