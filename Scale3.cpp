#include "stdafx.h"
#include "Scale3.h"
#include "Scale3MainWin.hpp"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

CScale3App ScaleApp;

BOOL CScale3App::InitInstance ()
{

   m_pMainWnd = new CScale3MainWin;
   m_pMainWnd->ShowWindow (m_nCmdShow);
   m_pMainWnd->UpdateWindow ();
   return true;

}


// End of module Scale3.cpp