#pragma once

#include <filesystem>

#include <wx/app.h>
#include <wx/cmdline.h>

class frmTest;

class TestApp : public wxApp
{
protected:
	virtual bool OnInit() override;
	
private:
	frmTest *m_frmTest{nullptr};
};


wxDECLARE_APP(TestApp);

