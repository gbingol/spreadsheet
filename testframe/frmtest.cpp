#include "frmtest.h" 

#include <codecvt>
#include <locale>


#include <wx/artprov.h>
#include <wx/sstream.h>
#include <wx/txtstrm.h>
#include <wx/dir.h>
#include <wx/url.h>


#include "icell.h"




frmTest::frmTest(wxWindow* parent): wxFrame(nullptr, wxID_ANY,"" )
{
	wxTheApp->SetTopWindow(this);

	m_Workbook = new ICELL::CWorkbook(this);
	m_Workbook->AddNewWorksheet();
	m_Workbook->Show();
	
	auto szrMain = new wxBoxSizer(wxVERTICAL);
	szrMain->Add(m_Workbook, 1, wxEXPAND);
	SetSizer(szrMain);
	Layout();
}