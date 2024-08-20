#include "frmtest.h" 


#include "icell.h"




frmTest::frmTest(wxWindow* parent): wxFrame(nullptr, wxID_ANY,"" )
{
	wxTheApp->SetTopWindow(this);

	m_Workbook = new ICELL::CWorkbook(this);
	m_Workbook->AddNewWorksheet();
	
	auto szrMain = new wxBoxSizer(wxVERTICAL);
	szrMain->Add(m_Workbook, 1, wxEXPAND);
	SetSizer(szrMain);
	Layout();
}