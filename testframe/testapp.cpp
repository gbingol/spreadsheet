#include "testapp.h"


#include "frmtest.h"



wxIMPLEMENT_APP(TestApp);


bool TestApp::OnInit()
{
		
	try {
		m_frmTest = new frmTest(nullptr);
		m_frmTest->Show();
		m_frmTest->Maximize();
		
	}	
	catch(std::exception& e) {
		wxMessageBox(e.what());
		return false;
	}

	return true;
}