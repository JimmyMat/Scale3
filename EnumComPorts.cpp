


/***************************************************************


System:     BioWindow -- Data Acquisition, Analysis, and Report Generator


File:       EnumComPorts.CPP


Procedure:  Implements all of the methods contained in the CComPorts
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


/*************************Additional Notes************************

ComPorts are enumerated out of the registry.

Function EnumerateComPorts is responsible for enumerating the available
comports. Enumeration means to get the string representation of the comport.
This string will be used to obtain the communications handle to the comport
via the CreateFile API function. These strings are located in the registry
under the following keys :

HKEY_LOCAL_MACHINE\HARDWARE\DEVICEMAP\SERIALCOMM

This key has the following values on an XP, SP3, laptop with no OEM serial ports
and with the StarTech 4 Port USB to RS-232 Serial Adapter, (ICUSB2324X), with
associated drivers installed :

      Value Name              Type                 Value Data
   \Device\Serial0            REG_SZ                  COM1
   \Device\Serial1            REG_SZ                  COM2
   \Device\Serial2            REG_SZ                  COM3
   \Device\Serial3            REG_SZ                  COM4

The Value's Data is the required string and is extracted and stored as a string
in the szName element of the COM_INFO structure.

As an added diagnostic aid, the COM_INFO structure also contains the member element
szSettings. The registry also contains the settings of the comports. These are
found under the following key:

HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS NT\CURRENTVERSION\PORTS

On the same computer, this key has the following values:

      Value Name              Type                 Value Data
         COM1:                REG_SZ               9600,n,8,1
         COM2:                REG_SZ               9600,n,8,1
         COM3:                REG_SZ               9600,n,8,1
         COM4:                REG_SZ               9600,n,8,1
         COM5:                REG_SZ               9600,n,8,1
         COM6:                REG_SZ               9600,n,8,1

This key also has values for network connections, LPT1, LPT2, LPT3, FILE,
Ne00, Ne01, and XPSPort, (14 total values, all values except COM1 - COM6
have no data). The function EnumerateComPortSettings extracts and assigns
the Value Data to the szSettings element of the appropeate COM_INFO structure
based on a comparison between the Value Name and the szName element. The settings
are stored as the whole string; no attempt is made to parse the settings string
and make the appropriate assignment to a SERIAL structure.
Comport settings will be modified during use with the MT scales. As a diagnostic
aid, it may be informative to know the initial comport settings before any use.


The MT, (Mettler Toledo), scale has a Standard Interface Command Set, (SICS), which has
multiple levels, (from Basic to whatever, documentation from 2000 lists 4 levels).
We will only use Level 0 commands and during actual use we'll only ever use one
command. The command is, "SI\r\n".
The SICS requires a specific command format and provides a specific command
response format.
The scale also has specific com port settings which will have to be set for
a selected port on the computer. The settings are, "2400,e,7,1" 


Had some real problems with getting valid enumerations of the com ports. Trying to use
the values returned from HKEY_LOCAL_MACHINE\HARDWARE\DEVICEMAP\SERIALCOMM did not always
work. There are two specific instances. First, if the value had double digits, such as
COM12, CreateFile would choke on it. Second, if the computer lost power, was restarted, or
shutdown, the values returned could not be used with CreateFile, (CreateFile would return
"The system could not find the file specified.". Placing the computer into Standby was OK.)
In this second instance, the only remedy
was to use Device Manager to open the properties of each of the comports and using the
StarTech driver software in the Configuration Tab, change the Com Port Assignment
of each com port. After exiting Device Manager, the returned values would then work.

This was unexpected and unacceptable. After research, I finally realized that what worked
under NT would not get the job done under XP. In Win2000 and up you have to use the SetupDi
functions found in the DDK and defined in SETUPAPI.H. You need an ID that is persistent across
system restarts. These are provided as Device Instance IDs. These IDs have a Device Path
associated with them. This Device Path is a string, (a TCHAR array of varying length dependent
on the device). Each com port will have a Device Path. In CreateFile, you must use this
Device Path string and not the value from HKEY_LOCAL_MACHINE\HARDWARE\DEVICEMAP\SERIALCOMM.
You can use that value as a string representation of a user friendly name, which I do. The
Device Path is not user friendly.

03/04/2010
While getting ready to do the installation and training on 03/08/2010, I set up a new Asus EPC1000HE
as a ultra-portable target machine. I loaded XP, SP3, along with all of the Asus supplied drivers for
the 1000HE's hardware. This included BlueTooth.
While testing out this MI2 supplied utility, SCALE3.EXE, I discovered that BlueTooth creates a
com port for blue tooth. After installing the StarTech drivers, the 1000HE has 5 serial ports listed
under HKEY_LOCAL_MACHINE\HARDWARE\DEVICEMAP\SERIALCOMM.

Function EnumerateComPorts:

Using RegEnumValue for cValues times, (cValues is the 8th parameter used in the RegQueryInfoKey and is the
number of values for the key "HARDWARE\\DEVICEMAP\\SERIALCOMM"), produces the following list.
\Device\BTPort0      COM3
\Device\Serial0      COM4
\Device\Serial1      COM5
\Device\Serial2      COM6
\Device\Serial3      COM7
EnumerateComPorts just grabed the first four com ports since I assumed that the computer would
only have the four Startech based com ports. This isn't a good assumption so I modified that part
of EnumerateComPorts that assigned the com port name to the individual m_ComPorts structures. I also
added a new m_ComPorts structure member, szRegistryName, so I could record the registry name of
each port for display.

               if (m_ulActiveCount < MAX_NUM_COMPORT)
               {

                  for (j = 0; j < MAX_NUM_COMPORT; j++)
                  {

                     if (m_ComPorts[j].ComFlag == FILLED)  // Been filled already
                        continue;

                     sprintf_s (szTestString, 32, "\\Device\\Serial%d", j);
                     iLen = (int)strnlen_s (szTestString, 32);
                     iResult1 = strncmp (ValueName, szTestString, iLen);
                     if (!iResult1)
                     {
			               m_ComPorts[j].ComFlag = FILLED;
			               strcpy_s (m_ComPorts[j].szName, COM_NAME_SIZE, (char *)&ValueData);
			               strcpy_s (m_ComPorts[j].szRegistryName, MAX_PATH, (char *)&ValueName);
			               m_ulActiveCount++;
                     }

                  }

               }

I still make assumptions.
The greatest weakness is the asumption that ValueName is always the string "\Device\Serial0",
or, "\Device\Serial1", or "\Device\Serial2", or "\Device\Serial3"
for the com ports that the user wants to use with my application.
The next greatest weakness is the asumption that there is a one-to-one
correspondence between the RegistryName and the physical port sequence on the startech
box, (that \Device\Serial0 corresponse to Port 1 on the box, \Device\Serial1 corresponse to Port 2,
etc...). This will have consequences in the important function EnumerateXPComPorts.

Function EnumerateXPComPorts:

EnumerateXPComPorts provides the all important szDevicePath which is the user un-friendly string
used in the CreateFile file that provides the handle to the com port, (as explained in the paragraph
directly above this 03/04 revision note). The problem is that the order of com ports is different
from the order of com ports provided by function EnumerateComPorts.
The function EnumerateComPorts lists the comports in chronological order using the RegEnumValue
function. The function EnumerateXPComPorts uses either the DDK function SetupDiEnumDeviceInfo or the
DDK function SetupDiEnumDeviceInterfaces. Either DDK function produces a list of com ports that is
not in chronological order. In this particular instance they both produce the following listing.
COM4
COM6
COM7
COM5
COM3
I must make a second assumption that COM4 from the DDK function is COM4 from the RegEnumValue,
COM5 from the DDK function is COM5 from the RegEnumValue. etc... This is vital since I'll be
placing the all important DevicePath found only with the DDK function in the m_ComPorts structure
that has the matching COM? name.

I modified EnumerateXPComPorts to provide the user friendly name for the com port, the manufacturer,
and the device description. This required the addition of three new members to the m_ComPorts
structure, szFriendlyName, szManufacturer, szDeviceDescrp, and szDeviceObject.
I will also modify the entire front-end so that this information is displayed in the list box.

Overall:

My first assumption is that the com ports of interest to the user of this program are going to
be listed as \Device\Serial0, \Device\Serial1, \Device\Serial2, and \Device\Serial3 in the
registry key "HKEY_LOCAL_MACHINE\HARDWARE\DEVICEMAP\SERIALCOMM".

My second assumption is that
\Device\Serial0 - COM4 corresponds to Port 1 on the StarTech box
\Device\Serial1 - COM5 corresponds to Port 2 on the StarTech box
\Device\Serial2 - COM6 corresponds to Port 3 on the StarTech box
\Device\Serial3 - COM7 corresponds to Port 4 on the StarTech box
Where COM4, COM5, COM6, and COM7 are the com port names provided in the program. This means the
user associates COM4 with Port 1, COM5 with Port 2, COM6 with Port 3, and COM7 with Port 4.

My third assumption is that I can use the com port name to associate the all important device
path with the corresponding m_ComPorts structure. I will use the com port name, (COM4, COM5,
COM6, and COM7), as the link between the RegEnumValue function and the DDK function.

I suppose I don't need to ever use the RegEnumValue or any Reg--- function. However, I don't
presently have any serial devices to hook up to the StarTech box. This means I don't have any
way to test. If there is time while I'm on site I will modify SCALE3.EXE so that it uses just the
DDK functions.

Also the function EnumerateComPortSettings is not needed. I will remove it and the corresponding
m_ComPorts structure member szSettings. Comport settings are modified during use with the MT scales.
These intial settings that appear to be static in the registry are of no use. Just examine the registry
at "SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Ports" if these initial values are of interest.

03/08/2010

ON-SITE MODIFICATIONS

USAISR security does not allow 3rd party programs to use Reg functions. So I had to remove
use of EnumerateComPorts. SCALE3 now uses just EnumerateXPComPorts.

EnumerateXPComPorts:

EnumerateXPComPorts now uses just the SetupDi functions. I no longer need the szRegistryName
member in the m_ComPorts structure. I am changing the use of the member szName. szName will now
refer to the StarTech box's port and not the com port.
I use the SetupDiGetDeviceRegistryProperty with the
SPDRP_PHYSICAL_DEVICE_OBJECT_NAME to get a string that resemble "\Device\SerBus?". These strings
seem to pertain to just the StarTech com ports. The strings come out as follows.
\Device\SerBus3
\Device\SerBus2
\Device\SerBus1
\Device\SerBus0

Note the lack of numerical order. There is a one to one correspondance between the numbers and the
StarTech box's port numbers. This means that the following order is true.
\Device\SerBus0 -> Port 1
\Device\SerBus1 -> Port 2
\Device\SerBus2 -> Port 3
\Device\SerBus3 -> Port 4

Overall:

Greatly simplified this module, EnumComPorts.

04/03/2015
See notes above new function EnumerateStartechWin8ComPorts.
Also renamed function EnumerateXPComPorts to EnumerateStartechXPComPorts. Did this to maintain Startech
dependancy. Don't know what other manufactures' drivers will produce in the strings returned by
SetupDiGetDeviceRegistryProperty.
*****************************************************************/


#include "stdafx.h"
#include "EnumComPorts.h"

CComPorts::CComPorts ()
{

   int            i;

   for (i = 0; i < MAX_NUM_COMPORT; i++)
   {

      /*******************************************************************************************

      03/22/2015

      Just trying to initialize things. However, if I try to use either a
      memset (&m_ComPorts[i], 0, sizeof (m_ComPorts)); or a
      SecureZeroMemory(&m_ComPorts[i], sizeof(m_ComPorts));, the heap is corrupted
      and you crash on the first Afx routine call. If you initialize each piece of the
      m_ComPorts structure one at a time as I am doing here, memory integrity is
      maintained and subsequent Afx calls succeed. I don't know why but I suspect it has
      something to do with:
      class CScale3MainWin : public CFrameWnd, public CComPorts
      That inheritance of CComPorts means CComPorts' constructor is called before/at the same time
      as CScale3MainWin's constructor where the first Afx call is made via the
      AfxRegisterWndClass routine.

      ********************************************************************************************/

      SecureZeroMemory (&(m_ComPorts[i].CommTimeouts), sizeof (COMMTIMEOUTS));
      SecureZeroMemory (&(m_ComPorts[i].dcb), sizeof (DCB));
      m_ComPorts[i].dcb.DCBlength = sizeof (DCB);
      // Fill in DCB: 9600 bps, 7 data bits, even parity, and 1 stop bit.
      m_ComPorts[i].dcb.BaudRate = CBR_9600;
      m_ComPorts[i].dcb.ByteSize = 7;
      m_ComPorts[i].dcb.Parity = EVENPARITY;
      m_ComPorts[i].dcb.StopBits = ONESTOPBIT;

      SecureZeroMemory(&(m_ComPorts[i].szDevicePath), sizeof(m_ComPorts[i].szDevicePath));
      SecureZeroMemory(&(m_ComPorts[i].szFriendlyName), sizeof(m_ComPorts[i].szFriendlyName));
      SecureZeroMemory(&(m_ComPorts[i].szManufacturer), sizeof(m_ComPorts[i].szManufacturer));
      SecureZeroMemory(&(m_ComPorts[i].szDeviceDescrp), sizeof(m_ComPorts[i].szDeviceDescrp));
      SecureZeroMemory(&(m_ComPorts[i].szDeviceObject), sizeof(m_ComPorts[i].szDeviceObject));
      SecureZeroMemory(&(m_ComPorts[i].szLocationInformation), sizeof(m_ComPorts[i].szLocationInformation));
      SecureZeroMemory(&(m_ComPorts[i].szName), sizeof(m_ComPorts[i].szName));

      m_ComPorts[i].hCom = (HANDLE)NULL;
      m_ComPorts[i].BaudRate = 9600;
      m_ComPorts[i].BaudRateIndex = 6;
      _tcscpy_s(m_ComPorts[i].Parity, 16, _T("EVEN"));
      m_ComPorts[i].ParityIndex = 2;
      m_ComPorts[i].DataBits = 7;
      m_ComPorts[i].DataBitsIndex = 2;
      m_ComPorts[i].StopBits = 1;
      m_ComPorts[i].StopBitsIndex = 0;
      m_ComPorts[i].sDCBSet = 0;
      m_ComPorts[i].sTimeoutsSet = 0;

#ifdef BS2PIN16
      m_ComPorts[i].BaudRate = 9600;
      m_ComPorts[i].BaudRateIndex = 6;
      _tcscpy_s(m_ComPorts[i].Parity, 16, _T("NONE"));
      m_ComPorts[i].ParityIndex = 0;
      m_ComPorts[i].DataBits = 8;
      m_ComPorts[i].DataBitsIndex = 3;
      m_ComPorts[i].StopBits = 1;
      m_ComPorts[i].StopBitsIndex = 0;
#endif

      m_ComPorts[i].ComFlag = EMPTY;

   }

   m_ulActiveCount = m_iCurrentPort = 0;

}


CComPorts::~CComPorts ()
{
 
}


int CComPorts::AreTherSerialPorts ()
{

   DWORD                            MyError = 0;
   HDEVINFO                         hDevInfo = INVALID_HANDLE_VALUE;

   hDevInfo = SetupDiGetClassDevs(
      &GUID_DEVINTERFACE_COMPORT,
      NULL,
      NULL,
      DIGCF_PRESENT | DIGCF_DEVICEINTERFACE);

   if (hDevInfo == INVALID_HANDLE_VALUE)
   {

      MyError = GetLastError();
      HRESULT_FROM_SETUPAPI(MyError);
      LPTSTR lpMsgBuf = (LPTSTR)NULL;
      FormatMessage(
         FORMAT_MESSAGE_ALLOCATE_BUFFER |
         FORMAT_MESSAGE_FROM_SYSTEM |
         FORMAT_MESSAGE_IGNORE_INSERTS,
         NULL,
         MyError,
         MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), // Default language
         (LPTSTR)&lpMsgBuf,
         0,
         NULL);

      // In debug, check lpMsgBuf before you get rid of it.

      if (lpMsgBuf != (LPTSTR)NULL)
      {
         LocalFree(lpMsgBuf);
      }

      return (0);

   }  // End of if (hDevInfo == INVALID_HANDLE_VALUE)

   SetupDiDestroyDeviceInfoList(hDevInfo);
   return (1);

}   // End of AreThereSerialPorts

int CComPorts::AreTherStartechSerialPorts()
{

   BOOL                             bResult1;
   int                              iCount, iLen;
   TCHAR                            szTemp[MAX_PATH];
   DWORD                            dwDevIndex, dwPropertyType,
                                    dwRequiredSize, MyError = 0;

   HDEVINFO                         hDevInfo = INVALID_HANDLE_VALUE;
   SP_DEVINFO_DATA                  spDevInfoData;

   long                             lResult1 = 0;
   TCHAR                            szTestString[COM_NAME_SIZE];

   iCount = 0;
   spDevInfoData.cbSize = sizeof(SP_DEVINFO_DATA);

   hDevInfo = SetupDiGetClassDevs(
      &GUID_DEVINTERFACE_COMPORT,
      NULL,
      NULL,
      DIGCF_PRESENT | DIGCF_DEVICEINTERFACE);

   if (hDevInfo == INVALID_HANDLE_VALUE)
   {

      MyError = GetLastError();
      HRESULT_FROM_SETUPAPI(MyError);
      LPTSTR lpMsgBuf = (LPTSTR)NULL;
      FormatMessage(
         FORMAT_MESSAGE_ALLOCATE_BUFFER |
         FORMAT_MESSAGE_FROM_SYSTEM |
         FORMAT_MESSAGE_IGNORE_INSERTS,
         NULL,
         MyError,
         MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), // Default language
         (LPTSTR)&lpMsgBuf,
         0,
         NULL);

      // In debug, check lpMsgBuf before you get rid of it.

      if (lpMsgBuf != (LPTSTR)NULL)
      {
         LocalFree(lpMsgBuf);
      }

      return (0);

   }  // End of if (hDevInfo == INVALID_HANDLE_VALUE)

   dwDevIndex = 0;
   while (SetupDiEnumDeviceInfo(
      hDevInfo,
      dwDevIndex,
      &spDevInfoData))
   {

      dwPropertyType = dwRequiredSize = 0;
      bResult1 = SetupDiGetDeviceRegistryProperty(
         hDevInfo,
         &spDevInfoData,
         SPDRP_DEVICEDESC,
         &dwPropertyType,
         (LPBYTE)szTemp,
         MAX_PATH,
         &dwRequiredSize);

      if (bResult1)
      {

         _stprintf_s(szTestString, COM_NAME_SIZE, _T("Startech.com Serial Adapter"));
         iLen = (int)_tcsnlen(szTestString, COM_NAME_SIZE);
         lResult1 = _tcsnccmp(szTemp, szTestString, iLen);
         if (!lResult1)
            iCount++;

      }  // End of if (!lResult1)

      dwDevIndex++;

   }  // End of while SetupDiEnumDeviceInfo loop

   SetupDiDestroyDeviceInfoList(hDevInfo);
   return (iCount);

}   // End of AreThereStartechSerialPorts

int CComPorts::EnumerateStartechXPComPorts()
{

   BOOL                             bResult1;
   int                              i, iLen;
   TCHAR                            szTemp[MAX_PATH];
   DWORD                            dwBufferSize, dwDevIndex, dwPropertyType,
                                    dwRequiredSize, MyError = 0;
   HDEVINFO                         hDevInfo = INVALID_HANDLE_VALUE;
   SP_DEVICE_INTERFACE_DATA         spDevInterfaceData;
   SP_DEVICE_INTERFACE_DETAIL_DATA  *pspDevInterfaceDetailData = NULL;
   SP_DEVINFO_DATA                  spDevInfoData;

   long                             lResult1 = 0;
   TCHAR                            szTestString[COM_NAME_SIZE];

   spDevInterfaceData.cbSize = sizeof (SP_DEVICE_INTERFACE_DATA);
   spDevInfoData.cbSize = sizeof (SP_DEVINFO_DATA);

   hDevInfo = SetupDiGetClassDevs (
                  &GUID_DEVINTERFACE_COMPORT,
                  NULL,
                  NULL,
                  DIGCF_PRESENT | DIGCF_DEVICEINTERFACE);

   if (hDevInfo == INVALID_HANDLE_VALUE)
   {

      MyError = GetLastError ();
      HRESULT_FROM_SETUPAPI (MyError);
      LPTSTR lpMsgBuf = (LPTSTR)NULL;
      FormatMessage ( 
         FORMAT_MESSAGE_ALLOCATE_BUFFER | 
            FORMAT_MESSAGE_FROM_SYSTEM | 
               FORMAT_MESSAGE_IGNORE_INSERTS,
         NULL,
         MyError,
         MAKELANGID (LANG_NEUTRAL, SUBLANG_DEFAULT), // Default language
         (LPTSTR)&lpMsgBuf,
         0,
         NULL);

      // In debug, check lpMsgBuf before you get rid of it.

      if (lpMsgBuf != (LPTSTR)NULL)
      {
         LocalFree (lpMsgBuf);
      }

      return (0);

   }  // End of if (hDevInfo == INVALID_HANDLE_VALUE)

   dwDevIndex = 0;
   while (SetupDiEnumDeviceInfo (
            hDevInfo,
            dwDevIndex,
            &spDevInfoData))
   {

      dwPropertyType = dwRequiredSize = 0;
      bResult1 = SetupDiGetDeviceRegistryProperty (
                                 hDevInfo,
                                 &spDevInfoData,
                                 SPDRP_PHYSICAL_DEVICE_OBJECT_NAME,
                                 &dwPropertyType,
                                 (LPBYTE)szTemp,
                                 MAX_PATH,
                                 &dwRequiredSize);
      if (bResult1)
      {

         for (i = 0; i < MAX_NUM_COMPORT; i++)
         {

            if (m_ComPorts[i].ComFlag == FILLED)  // Been filled already
               continue;

            _stprintf_s(szTestString, COM_NAME_SIZE, _T("\\Device\\SerBus%d"), i);
            iLen = (int)_tcsnlen (szTestString, COM_NAME_SIZE);
            lResult1 = _tcsnccmp (szTemp, szTestString, iLen);
            if (!lResult1)
            {

			      _tcscpy_s (m_ComPorts[i].szDeviceObject, 64, szTemp);

               dwPropertyType = dwRequiredSize = 0;
               bResult1 = SetupDiGetDeviceRegistryProperty (
                                 hDevInfo,
                                 &spDevInfoData,
                                 SPDRP_FRIENDLYNAME,
                                 &dwPropertyType,
                                 (LPBYTE)szTemp,
                                 MAX_PATH,
                                 &dwRequiredSize);
               _tcscpy_s (m_ComPorts[i].szFriendlyName, 64, szTemp);

               dwPropertyType = dwRequiredSize = 0;
               bResult1 = SetupDiGetDeviceRegistryProperty (
                                 hDevInfo,
                                 &spDevInfoData,
                                 SPDRP_MFG,
                                 &dwPropertyType,
                                 (LPBYTE)szTemp,
                                 MAX_PATH,
                                 &dwRequiredSize);
               _tcscpy_s (m_ComPorts[i].szManufacturer, 64, szTemp);

               dwPropertyType = dwRequiredSize = 0;
               bResult1 = SetupDiGetDeviceRegistryProperty (
                                 hDevInfo,
                                 &spDevInfoData,
                                 SPDRP_DEVICEDESC,
                                 &dwPropertyType,
                                 (LPBYTE)szTemp,
                                 MAX_PATH,
                                 &dwRequiredSize);
               _tcscpy_s (m_ComPorts[i].szDeviceDescrp, 64, szTemp);

               SetupDiEnumDeviceInterfaces (
                     hDevInfo,
                     NULL,
                     &GUID_DEVINTERFACE_COMPORT,
                     dwDevIndex,
                     &spDevInterfaceData);

               dwBufferSize = 0;
               SetupDiGetDeviceInterfaceDetail (
                     hDevInfo,
                     &spDevInterfaceData,
                     NULL,
                     0,
                     &dwBufferSize,
                     NULL);

               if ((GetLastError () == ERROR_INSUFFICIENT_BUFFER) && (dwBufferSize > 0))
               {
                  if (pspDevInterfaceDetailData)
                     free (pspDevInterfaceDetailData);
                  pspDevInterfaceDetailData =
                     (SP_DEVICE_INTERFACE_DETAIL_DATA *)calloc (dwBufferSize, sizeof (char));
                  pspDevInterfaceDetailData->cbSize = sizeof (SP_DEVICE_INTERFACE_DETAIL_DATA);
                  SetupDiGetDeviceInterfaceDetail (
                        hDevInfo,
                        &spDevInterfaceData,
                        pspDevInterfaceDetailData,
                        dwBufferSize,
                        NULL,
                        NULL);
                  if (dwBufferSize < MAX_KEY_LENGTH)
                     _stprintf_s (m_ComPorts[i].szDevicePath,
                                  MAX_KEY_LENGTH, pspDevInterfaceDetailData->DevicePath);
               }

               m_ulActiveCount++;
               m_ComPorts[i].ComFlag = FILLED;
               _stprintf_s (m_ComPorts[i].szName, 32, _T("Port %d"), (i + 1));
               break;

            }  // End of if (!lResult1)

         }  // End of for (i = 0; i < MAX_NUM_COMPORT; i++)

      }  // End of if (bResult1)

      dwDevIndex++;

   }  // End of while SetupDiEnumDeviceInfo loop

   SetupDiDestroyDeviceInfoList (hDevInfo);
   if (pspDevInterfaceDetailData)
      free (pspDevInterfaceDetailData);
   return (1);

}  // End of EnumerateStartechXPComPorts

/************************************************************************************

04/03/2015

As feared, the assumption that SPDRP_PHYSICAL_DEVICE_OBJECT_NAME would return a consistent
string that I could hardcode a comparison string with the for loop iteration variable suffixed,
has been broken. The SPDRP_PHYSICAL_DEVICE_OBJECT_NAME does not end with a predictable number.
further research shows that the SPDRP_LOCATION_INFORMATION does return a predictable string with
a number suffix that is predictable although not in order.

The following are the results of an investigation of what the Win7/8 driver puts in the registry.
Using the SetupDiGetDeviceRegistryProperty and 10 of the DDK defined SPDRP_ #defines. I investigated
these ten SPDRP_'s based on the description of what they should return in the documentation on the
SetupDiGetDeviceRegistryProperty function, (all but one return a string).
Note - dwDeviceIndex is just a local whose value is kept between 0 to 3.
I am working only with a four port Startech serial to usb hub so the maximum number of
serial ports is 4, as #define MAX_NUM_COMPORT in EnumComPorts.h.

SPDRP_ #define sent to SetupDiGetDeviceRegistryProperty
1)    SPDRP_CLASS
2)    SPDRP_DEVICEDESC
3)    SPDRP_DEVTYPE
4)    SPDRP_ENUMERATOR_NAME
5)    SPDRP_FRIENDLYNAME
6)    SPDRP_LOCATION_INFORMATION
7)    SPDRP_MFG
8)    SPDRP_PHYSICAL_DEVICE_OBJECT_NAME
9)    SPDRP_SERVICE
10)   SPDRP_UI_NUMBER
11)   SPDRP_UI_NUMBER_DESC_FORMAT

1, 2, 4, 7, and 9 return the same string for each of the four ports.
3, 10, and 11 return a failure. Each produces the system error message "The data is invalid.". This
   indicates that the driver ignores this registry entry.
5, 6, and 8 return unique strings that seem to be port specific and require further investigation.

SPDRP_LOCATION_INFORMATION returns a string from the Startech driver as follows:
   Port_#000x.Hub_#000y
   where x is a port number and y is hub number, duh.
For example, on my remote debug machine with the four port unit, part #ICUSB2324X, the following strings
are returned:
   dwDevIndex value     Location Information       Freindly Name                          Physical Device Name
         0              Port_#0002.Hub_#0006       Startech.com Serial Adapter (COM4)     \\Device\\USBPDO-9
         1              Port_#0003.Hub_#0006       Startech.com Serial Adapter (COM5)     \\Device\\USBPDO-10
         2              Port_#0004.Hub_#0006       Startech.com Serial Adapter (COM6)     \\Device\\USBPDO-11
         3              Port_#0001.Hub_#0006       Startech.com Serial Adapter (COM3)     \\Device\\USBPDO-7
As you can see, the Physical Device Name is useless for sorting and interpretation via a predictable string
comparison. Did I mention predictability? Well that is out of the question as the Physical Device Name has a
tendency to change from one boot to the next as well. The friendly name has COM3 thru COM4 but this too can
change from boot to boot. Also the user can manually set this COMx via the Device Manager.
Location Information is the string to use. I have not seen it change from boot to boot. The actual device
has its RS232 connectors labeled PORT1, PORT2, PORT3, and PORT4 so the corelation between registry value
and actual device is clear. This is very important as it allows you to easily generate a GUI.

I have also added code to make sure we're looking at Startech.com equipment via the new functions
AreThereSerialPorts and AreThereStartechSerialPorts. These functions are called as soon as the
frame window is created. If either returns false the program is terminated by a WM_CLOSE message.
This is purposely redundant since I want serial port information weither or not a Startech device
is attached.

*************************************************************************************/

int CComPorts::EnumerateStartechWin8ComPorts()
{

   BOOL                             bResult1;
   int                              i, iFilledComPortStrucs, iGotOne, iLen;
   TCHAR                            szTemp[MAX_PATH];
   DWORD                            dwBufferSize, dwDevIndex, dwPropertyType,
                                    dwRequiredSize, MyError = 0;

   HDEVINFO                         hDevInfo = INVALID_HANDLE_VALUE;
   SP_DEVICE_INTERFACE_DATA         spDevInterfaceData;
   SP_DEVICE_INTERFACE_DETAIL_DATA  *pspDevInterfaceDetailData = NULL;
   SP_DEVINFO_DATA                  spDevInfoData;

   long                             lResult1 = 0;
   TCHAR                            szTestString[COM_NAME_SIZE];

   spDevInterfaceData.cbSize = sizeof(SP_DEVICE_INTERFACE_DATA);
   spDevInfoData.cbSize = sizeof(SP_DEVINFO_DATA);

   hDevInfo = SetupDiGetClassDevs(
      &GUID_DEVINTERFACE_COMPORT,
	  //&GUID_DEVINTERFACE_USB_DEVICE,
      NULL,
      NULL,
      DIGCF_PRESENT | DIGCF_DEVICEINTERFACE);

   if (hDevInfo == INVALID_HANDLE_VALUE)
   {

      MyError = GetLastError();
      HRESULT_FROM_SETUPAPI(MyError);
      LPTSTR lpMsgBuf = (LPTSTR)NULL;
      FormatMessage(
         FORMAT_MESSAGE_ALLOCATE_BUFFER |
         FORMAT_MESSAGE_FROM_SYSTEM |
         FORMAT_MESSAGE_IGNORE_INSERTS,
         NULL,
         MyError,
         MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), // Default language
         (LPTSTR)&lpMsgBuf,
         0,
         NULL);

      // In debug, check lpMsgBuf before you get rid of it.

      if (lpMsgBuf != (LPTSTR)NULL)
      {
         LocalFree(lpMsgBuf);
      }

      return (0);

   }  // End of if (hDevInfo == INVALID_HANDLE_VALUE)

   dwDevIndex = 0;
   iFilledComPortStrucs = 0;
   while (SetupDiEnumDeviceInfo(
      hDevInfo,
      dwDevIndex,
      &spDevInfoData) && (iFilledComPortStrucs < MAX_NUM_COMPORT))
   {

      dwPropertyType = dwRequiredSize = 0;
      bResult1 = SetupDiGetDeviceRegistryProperty(
         hDevInfo,
         &spDevInfoData,
         SPDRP_LOCATION_INFORMATION,
         &dwPropertyType,
         (LPBYTE)szTemp,
         MAX_PATH,
         &dwRequiredSize);

      if (bResult1)
      {

         iGotOne = 0;
         for (i = 0; i < MAX_NUM_COMPORT; i++)
         {

            _stprintf_s(szTestString, COM_NAME_SIZE, _T("Port_#000%d"), (i + 1));
            iLen = (int)_tcsnlen(szTestString, COM_NAME_SIZE);
            lResult1 = _tcsnccmp(szTemp, szTestString, iLen);
            if (!lResult1)
            {
               iGotOne = 1;
               break;
            }

         }

         if (!iGotOne)
         {
            dwDevIndex++;
            continue;
         }

         else
         {

            if (m_ComPorts[i].ComFlag == FILLED)  // Been filled already
               continue;

            iFilledComPortStrucs++;
            _tcscpy_s(m_ComPorts[i].szLocationInformation, 64, szTemp);

            dwPropertyType = dwRequiredSize = 0;
            bResult1 = SetupDiGetDeviceRegistryProperty(
                  hDevInfo,
                  &spDevInfoData,
                  SPDRP_PHYSICAL_DEVICE_OBJECT_NAME,
                  &dwPropertyType,
                  (LPBYTE)szTemp,
                  MAX_PATH,
                  &dwRequiredSize);
            _tcscpy_s(m_ComPorts[i].szDeviceObject, 64, szTemp);

            dwPropertyType = dwRequiredSize = 0;
            bResult1 = SetupDiGetDeviceRegistryProperty(
                  hDevInfo,
                  &spDevInfoData,
                  SPDRP_FRIENDLYNAME,
                  &dwPropertyType,
                  (LPBYTE)szTemp,
                  MAX_PATH,
                  &dwRequiredSize);
            _tcscpy_s(m_ComPorts[i].szFriendlyName, 64, szTemp);

            dwPropertyType = dwRequiredSize = 0;
            bResult1 = SetupDiGetDeviceRegistryProperty(
                  hDevInfo,
                  &spDevInfoData,
                  SPDRP_MFG,
                  &dwPropertyType,
                  (LPBYTE)szTemp,
                  MAX_PATH,
                  &dwRequiredSize);
            _tcscpy_s(m_ComPorts[i].szManufacturer, 64, szTemp);

            dwPropertyType = dwRequiredSize = 0;
            bResult1 = SetupDiGetDeviceRegistryProperty(
                  hDevInfo,
                  &spDevInfoData,
                  SPDRP_DEVICEDESC,
                  &dwPropertyType,
                  (LPBYTE)szTemp,
                  MAX_PATH,
                  &dwRequiredSize);
            _tcscpy_s(m_ComPorts[i].szDeviceDescrp, 64, szTemp);

            SetupDiEnumDeviceInterfaces(
                  hDevInfo,
                  NULL,
                  &GUID_DEVINTERFACE_COMPORT,
                  dwDevIndex,
                  &spDevInterfaceData);

            dwBufferSize = 0;
            SetupDiGetDeviceInterfaceDetail(
                  hDevInfo,
                  &spDevInterfaceData,
                  NULL,
                  0,
                  &dwBufferSize,
                  NULL);

            if ((GetLastError() == ERROR_INSUFFICIENT_BUFFER) && (dwBufferSize > 0))
            {
               if (pspDevInterfaceDetailData)
                  free(pspDevInterfaceDetailData);
               pspDevInterfaceDetailData =
                  (SP_DEVICE_INTERFACE_DETAIL_DATA *)calloc(dwBufferSize, sizeof(char));
               pspDevInterfaceDetailData->cbSize = sizeof(SP_DEVICE_INTERFACE_DETAIL_DATA);
               SetupDiGetDeviceInterfaceDetail(
                     hDevInfo,
                     &spDevInterfaceData,
                     pspDevInterfaceDetailData,
                     dwBufferSize,
                     NULL,
                     NULL);
               if (dwBufferSize < MAX_KEY_LENGTH)
                  _stprintf_s(m_ComPorts[i].szDevicePath,
                              MAX_KEY_LENGTH, pspDevInterfaceDetailData->DevicePath);
            }

            m_ulActiveCount++;
            m_ComPorts[i].ComFlag = FILLED;
            _stprintf_s(m_ComPorts[i].szName, 32, _T("Port %d"), (i + 1));

         }  // End of else

      }  // End of if (bResult1)

      dwDevIndex++;

   }  // End of while SetupDiEnumDeviceInfo loop

   SetupDiDestroyDeviceInfoList(hDevInfo);
   if (pspDevInterfaceDetailData)
      free(pspDevInterfaceDetailData);
   return (1);

}  // End of EnumerateStartechWin8ComPorts

/******************************************************************************************************/

int CComPorts::EnumerateStartechWin8USBPorts()
{

   BOOL                             bResult1;
   int                              i, iGotOne, iLen;
   TCHAR                            szTemp[MAX_PATH];
   DWORD                            dwBufferSize, dwDevIndex, dwPropertyType,
                                    dwRequiredSize, MyError = 0;

   HDEVINFO                         hDevInfo = INVALID_HANDLE_VALUE;
   SP_DEVICE_INTERFACE_DATA         spDevInterfaceData;
   SP_DEVICE_INTERFACE_DETAIL_DATA  *pspDevInterfaceDetailData = NULL;
   SP_DEVINFO_DATA                  spDevInfoData;

   long                             lResult1 = 0;
   TCHAR                            szTestString[COM_NAME_SIZE];

   spDevInterfaceData.cbSize = sizeof(SP_DEVICE_INTERFACE_DATA);
   spDevInfoData.cbSize = sizeof(SP_DEVINFO_DATA);

   hDevInfo = SetupDiGetClassDevs(
	   &GUID_DEVINTERFACE_USB_DEVICE,
      NULL,
      NULL,
      DIGCF_PRESENT | DIGCF_DEVICEINTERFACE);

   if (hDevInfo == INVALID_HANDLE_VALUE)
   {

      MyError = GetLastError();
      HRESULT_FROM_SETUPAPI(MyError);
      LPTSTR lpMsgBuf = (LPTSTR)NULL;
      FormatMessage(
         FORMAT_MESSAGE_ALLOCATE_BUFFER |
         FORMAT_MESSAGE_FROM_SYSTEM |
         FORMAT_MESSAGE_IGNORE_INSERTS,
         NULL,
         MyError,
         MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), // Default language
         (LPTSTR)&lpMsgBuf,
         0,
         NULL);

      // In debug, check lpMsgBuf before you get rid of it.

      if (lpMsgBuf != (LPTSTR)NULL)
      {
         LocalFree(lpMsgBuf);
      }

      return (0);

   }  // End of if (hDevInfo == INVALID_HANDLE_VALUE)

   dwDevIndex = 0;
   while (SetupDiEnumDeviceInfo(
      hDevInfo,
      dwDevIndex,
      &spDevInfoData))
   {

      dwPropertyType = dwRequiredSize = 0;
      bResult1 = SetupDiGetDeviceRegistryProperty(
         hDevInfo,
         &spDevInfoData,
         SPDRP_LOCATION_INFORMATION,
         &dwPropertyType,
         (LPBYTE)szTemp,
         MAX_PATH,
         &dwRequiredSize);

      if (bResult1)
      {

         iGotOne = 0;
         for (i = 0; i < MAX_NUM_COMPORT; i++)
         {

            _stprintf_s(szTestString, COM_NAME_SIZE, _T("Port_#000%d"), (i + 1));
            iLen = (int)_tcsnlen(szTestString, COM_NAME_SIZE);
            lResult1 = _tcsnccmp(szTemp, szTestString, iLen);
            if (!lResult1)
            {
               iGotOne = 1;
               break;
            }

         }

         if (!iGotOne)
         {
            dwDevIndex++;
            continue;
         }

         else
         {

            dwPropertyType = dwRequiredSize = 0;
            bResult1 = SetupDiGetDeviceRegistryProperty(
                  hDevInfo,
                  &spDevInfoData,
                  SPDRP_PHYSICAL_DEVICE_OBJECT_NAME,
                  &dwPropertyType,
                  (LPBYTE)szTemp,
                  MAX_PATH,
                  &dwRequiredSize);

            dwPropertyType = dwRequiredSize = 0;
            bResult1 = SetupDiGetDeviceRegistryProperty(
                  hDevInfo,
                  &spDevInfoData,
                  SPDRP_FRIENDLYNAME,
                  &dwPropertyType,
                  (LPBYTE)szTemp,
                  MAX_PATH,
                  &dwRequiredSize);

            dwPropertyType = dwRequiredSize = 0;
            bResult1 = SetupDiGetDeviceRegistryProperty(
                  hDevInfo,
                  &spDevInfoData,
                  SPDRP_MFG,
                  &dwPropertyType,
                  (LPBYTE)szTemp,
                  MAX_PATH,
                  &dwRequiredSize);

            dwPropertyType = dwRequiredSize = 0;
            bResult1 = SetupDiGetDeviceRegistryProperty(
                  hDevInfo,
                  &spDevInfoData,
                  SPDRP_DEVICEDESC,
                  &dwPropertyType,
                  (LPBYTE)szTemp,
                  MAX_PATH,
                  &dwRequiredSize);

            SetupDiEnumDeviceInterfaces(
                  hDevInfo,
                  NULL,
                  &GUID_DEVINTERFACE_USB_DEVICE,
				      dwDevIndex,
                  &spDevInterfaceData);

            dwBufferSize = 0;
            SetupDiGetDeviceInterfaceDetail(
                  hDevInfo,
                  &spDevInterfaceData,
                  NULL,
                  0,
                  &dwBufferSize,
                  NULL);

            if ((GetLastError() == ERROR_INSUFFICIENT_BUFFER) && (dwBufferSize > 0))
            {
               if (pspDevInterfaceDetailData)
                  free(pspDevInterfaceDetailData);
               pspDevInterfaceDetailData =
                  (SP_DEVICE_INTERFACE_DETAIL_DATA *)calloc(dwBufferSize, sizeof(char));
               pspDevInterfaceDetailData->cbSize = sizeof(SP_DEVICE_INTERFACE_DETAIL_DATA);
               SetupDiGetDeviceInterfaceDetail(
                     hDevInfo,
                     &spDevInterfaceData,
                     pspDevInterfaceDetailData,
                     dwBufferSize,
                     NULL,
                     NULL);
               if (dwBufferSize < MAX_KEY_LENGTH)
                  _stprintf_s(szTemp,
                              MAX_KEY_LENGTH, pspDevInterfaceDetailData->DevicePath);
            }

            m_ulActiveCount++;

         }  // End of else

      }  // End of if (bResult1)

      dwDevIndex++;

   }  // End of while SetupDiEnumDeviceInfo loop

   SetupDiDestroyDeviceInfoList(hDevInfo);
   if (pspDevInterfaceDetailData)
      free(pspDevInterfaceDetailData);
   return (1);

}  // End of EnumerateStartechWin8USBPorts

// End of EnumComPorts.CPP


