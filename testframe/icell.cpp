#include "icell.h"

#include <string>
#include <sstream>
#include <codecvt>
#include <locale>

#include <wx/clipbrd.h>
#include <wx/xml/xml.h>
#include <wx/artprov.h>
#include <wx/tokenzr.h>
#include <wx/fontenum.h>
#include <wx/colordlg.h>

#include <grid/ws_funcs.h>
#include <grid/rangebase.h>
#include <grid/worksheetbase.h>
#include <grid/workbookbase.h>
#include <grid/ws_cell.h>
#include <grid/undoredo.h>
#include <grid/events.h>




namespace ICELL
{
	CWorksheet::CWorksheet(wxWindow* parent,
		CWorkbook* workbook,
		wxWindowID id,
		const wxPoint& pos,
		const wxSize& size,
		long style,
		wxString WindowName,
		int nrows,
		int ncols)
		:grid::CWorksheetBase(parent, workbook, id, pos, size, style, WindowName, nrows, ncols)
	{
		m_Workbook = workbook;
		m_ParentWnd = parent;

		SetWSName(WindowName.ToStdWstring());

		//Bind(wxEVT_GRID_SELECT_CELL, &CWorksheet::OnSelectCell, this);
		Bind(wxEVT_GRID_CELL_RIGHT_CLICK, &CWorksheet::OnRightClick, this);
	}

	CWorksheet::~CWorksheet() = default;



	void CWorksheet::OnRightClick(wxGridEvent& event)
	{
		if (GetNumSelRows() == 1 && GetNumSelCols() == 1)
			return;

		if (m_ContextMenu) {
			delete m_ContextMenu;
			m_ContextMenu = nullptr;
		}

		m_ContextMenu = new wxMenu();

		auto Menu_Copy = m_ContextMenu->Append(wxID_ANY, "Copy");
		Menu_Copy->SetBitmap(wxArtProvider::GetBitmap(wxART_COPY));
		m_ContextMenu->Bind(wxEVT_MENU, [this](wxCommandEvent& e) { Copy(); }, Menu_Copy->GetId());

		auto Menu_Cut = m_ContextMenu->Append(wxID_ANY, "Cut");
		Menu_Cut->SetBitmap(wxArtProvider::GetBitmap(wxART_CUT));
		m_ContextMenu->Bind(wxEVT_MENU, [this](wxCommandEvent& e) { Cut(); }, Menu_Cut->GetId());

		if (!IsSelection()) {
			m_ContextMenu->AppendSeparator();
			
			auto Menu_Paste = m_ContextMenu->Append(wxID_ANY, "Paste");
			Menu_Paste->SetBitmap(wxArtProvider::GetBitmap(wxART_PASTE));
			Menu_Paste->Enable(false);

			if (!wxTheClipboard->IsSupported(wxDF_INVALID))
			{
				Menu_Paste->Enable(true);
				m_ContextMenu->Bind(wxEVT_MENU, [this](wxCommandEvent &event) { 
					Paste(); 
				}, Menu_Paste->GetId());

				if (grid::SupportsXML()) {
					auto PasteVal = m_ContextMenu->Append(wxID_ANY, "Paste Values");
					auto PasteFrmt = m_ContextMenu->Append(wxID_ANY, "Paste Format");

					m_ContextMenu->Bind(wxEVT_MENU, [this](wxCommandEvent& ) { 
						GetWorkbook()->PasteValues(wxDF_TEXT); 
					}, PasteVal->GetId());

					m_ContextMenu->Bind(wxEVT_MENU, [this](wxCommandEvent& ) { 
						GetWorkbook()->PasteFormat(grid::XMLDataFormat()); 
					}, PasteFrmt->GetId());
				}
			}
		}

		PopupMenu(m_ContextMenu);
		event.Skip();
	}



	/******************************************************* */

	CWorksheetNtbk::CWorksheetNtbk(CWorkbook* parent) :
		grid::CWorksheetNtbkBase(parent), m_Workbook{ parent }
	{
	}


	CWorksheetNtbk::~CWorksheetNtbk() = default;


	grid::CWorksheetBase* CWorksheetNtbk::CreateWorksheet(
		wxWindow* wnd, 
		const std::wstring& Label, 
		int nrows, 
		int ncols) const
	{
		return  new CWorksheet(wnd, 
					m_Workbook, 
					wxID_ANY, 
					wxDefaultPosition, 
					wxDefaultSize, 0, Label, nrows, ncols);;
	}



	/******************************************************************/


	CWorkbook::CWorkbook(wxWindow* parent): grid::CWorkbookBase(parent)
	{
		m_WSNtbk = new CWorksheetNtbk(this);

		auto szrMain = new wxBoxSizer(wxVERTICAL);
		szrMain->Add(m_WSNtbk, 1, wxEXPAND, 0);
		SetSizerAndFit(szrMain);
		Layout();

	}

	CWorkbook::~CWorkbook() = default;
}