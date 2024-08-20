#pragma once

#include <filesystem>
#include <fstream>


#include <wx/wx.h>


namespace ICELL 
{ 
	class CWorkbook; 
}




class frmTest: public wxFrame
{

public:
	frmTest(wxWindow* parent);
	~frmTest() = default;

private:
		

	ICELL::CWorkbook* m_Workbook{ nullptr };
	
};


