# Scale3
VS2013, MFC based app that does serial IO via Startech USB to RS232 serial adapter in Win8.1

Company:
Modular Instruments, Inc.
1436 Unionville-Wawaset Road
West Chester, PA 19382

Programmer:
James Matthews
Begin - March 3, 2015
End   - June 21, 2015

Description:
SCALE3 is a test bed project. It is testing the use of a Startech.com Professional
4 port USB to RS232 serial adaptor, Part# ICUSB2324X. This device will be used as
part of a proprietary life sciences laboratory data acquisition and reporting system.
This project investigates what is required to use this device under Windows 8.1 in
a MFC C++ program.
Its main areas of investigation include:
	1) Installation of the Startech device under Windows 8.1.
	2) Intialization and use of the Startech device within an MFC program under Windows 8.1.
	3) RS232 serial port I/O API's within an MFC program under Windows 8.1
	4) Multi threaded programming under Windows 8.1.
	5) Development of MCU test hardware to simulate the target serial lab instrument.
	6) Testing using all of the componets.

History:
A user of the proprietary data acquisition system needs to incorporate an RS232 serial
lab instrument. This instrument is a Metlier Toledo lab scale. Thus the name Scale3. The
user will need to incorporate up to four of these instruments. Thus the 4 port adaptor.
The Startech device was selected based on the TI chipset it uses and the quality/ruggedness
of the adaptor's enclosure.
The proprietary data acquisition system's software must now communicate with the existing
proprietary D/A hardware and the RS232 serial lab instrument at the same time. Thus the use
of muti threads is required.
Since the use of actual Metlier Toledo Scales for testing purposes is out of the question,
a realistic physical simulation device is required. The simulator must use RS232 serial I/O
and be capable of duplicating the lab instruments functions. Thus the use of an MCU based
hardware which can be programed to simulate these functions.

Processcie:
The Startech USB to RS232 serial adapter must be installed sucessfully.
There are three source files. Scale3.cpp, Scale3MainWin.cpp, and EnumComPorts.cpp.
Scale3.cpp is responsible for InitInstance of a CwinApp which in turn creates a new CScale3MainWin.
Scale3MainWin.cpp is responsible for taking care of the CScale3MainWin class. CScale3MainWin is a
CFrameWnd that I flesh out with controls and their associated functions. Its a straight up Windows
program and I intentionally  didn't bother to break it up so its over 3400 lines long. It has a ton
of comments. One of the controls is a list box which shows the available com ports. When OnCreate
fills this listbox, it calls a function, EnumerateStartechWin8ComPorts. this function takes up most
of the next source file.
EnumComPorts.cpp is responsible for reading the registry, and filling out the CComPorts class. The
most important thing here is getting the user unfriendly string m_ComPorts[?].szDevicePath. This is
the string used by CreateFile to produce a handle to a com port. All I/O calls use this handle.
No handle, no I/O. This file also has many comments.

Simulator:
I needed a real RS232 device. I had some 1990's era Parrallax boards populated with BS2 Stamp MCUs.
These boards are used for programming the BS2 and interfacing with it. They used RS232 to communicate
with the computer via db9 straight thru cables. These were perfect and I programmed two boards with
programs that simulate the required actions of the Mettler Toledo SR1600 scale. The programs duplicate
the command sequence of the scale. You must set the com ports to 9600 baud 8N1 to use these.

Keywords & Concepts:
RS232 serial I/O in Windows 8.1.
Multi threaded I/O in Windows 8.1.
C++ program using MFC in a Visual Studio 2013 IDE.
Ancient MCU programming via the serial port using ParraLLax hardware, editor, and MTScale1.bs2 & MTScale2.bs2.

Read the source file comments. I don't include any of the VS2013 files since I don't think they'll be relavent
to any other environment but my own. Just make a new empty MFC project and then place the files in it.
