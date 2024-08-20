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
	wxDEFINE_EVENT(ssEVT_WB_PAGECHANGED, wxAuiNotebookEvent);
	wxDEFINE_EVENT(ssEVT_GRID_SELECTION_BEGUN, wxGridRangeSelectEvent);



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

		Bind(wxEVT_GRID_SELECT_CELL, &CWorksheet::OnSelectCell, this);

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

			PopupMenu(m_ContextMenu);

			return;
		}


		PopupMenu(m_ContextMenu);

		event.Skip();
	}



	void CWorksheet::OnSelectCell(wxGridEvent& event)
	{
		wxString StatBarInfo;
		StatBarInfo << grid::ColNumtoLetters(event.GetCol() + 1) << (event.GetRow() + 1);

		event.Skip();
	}





	/******************************************************* */

	CWorksheetNtbk::CWorksheetNtbk(CWorkbook* parent) :
		grid::CWorksheetNtbkBase(parent), m_Workbook{ parent }
	{
		Bind(wxEVT_AUINOTEBOOK_PAGE_CHANGED, &CWorksheetNtbk::OnPageChanged, this);	
	}


	CWorksheetNtbk::~CWorksheetNtbk() = default;


	void CWorksheetNtbk::OnPageChanged(wxAuiNotebookEvent& evt) 
	{
		m_ActiveWS = FindWorksheet(evt.GetSelection());
		m_ActiveWS->ClearSelection();
		m_ActiveWS->SetFocus();

		wxAuiNotebookEvent PageEvt(evt);
		PageEvt.SetEventType(ssEVT_WB_PAGECHANGED);
		ProcessWindowEvent(PageEvt);
	}


	void CWorksheetNtbk::OnTabRightDown(wxAuiNotebookEvent& evt)
	{

		PopupMenu(m_ContextMenu);
	}

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
		//start with yellow color (changes as user selects a new one)
		m_FillColor = wxColor(255, 255, 0); 

		//start with red color
		m_FontColor = wxColor(255, 0, 0); 

		m_WSNtbk = new CWorksheetNtbk(this);

		auto szrMain = new wxBoxSizer(wxVERTICAL);
		szrMain->Add(m_WSNtbk, 1, wxEXPAND, 0);
		SetSizerAndFit(szrMain);
		Layout();

		Bind(ssEVT_WB_UNDOREDO, &CWorkbook::OnUndoRedoStackChanged, this);

	}


	CWorkbook::~CWorkbook() = default;



	void CWorkbook::OnUpdateUI(wxUpdateUIEvent & evt)
	{
		auto ws = GetActiveWS();

		int row = ws->GetGridCursorRow();
		int col = ws->GetGridCursorCol();

		wxFont font = ws->GetCellFont(row, col);

		int evtID = evt.GetId();

		if (evtID == ID_FONTBOLD)
			evt.Check(font.GetWeight() == wxFONTWEIGHT_BOLD);

		else if (evtID == ID_FONTITALIC)
			evt.Check(font.GetStyle() == wxFONTSTYLE_ITALIC);

		else if (evtID == ID_FONTSIZE)
			m_ComboFontSize->SetValue(std::to_string(font.GetPointSize()));
		
		else if (evtID == ID_FONTFACE)
			m_ComboFontFace->SetValue(font.GetFaceName());
	}



	void CWorkbook::OnFontChanged(wxCommandEvent & event) //Font size or face
	{
		int evtID = event.GetId();

		if (evtID == ID_FONTSIZE)
		{
			long fontSize = 11;

			if (!m_ComboFontSize->GetStringSelection().ToLong(&fontSize)){
				wxMessageBox("ERROR: Not a valid font size!");
				return;
			}

			ChangeFontSize(fontSize);
		}

		else if (evtID == ID_FONTFACE)
		{
			wxString fontFace = m_ComboFontFace->GetStringSelection();
			wxFont fnt;
			fnt.SetFaceName(fontFace);

			if (!fnt.IsOk())
			{
				wxMessageBox("Font face is not valid!");
				return;
			}

			ChangeFontFace(fontFace);
		}
	}


	void CWorkbook::OnFillFontColor(wxAuiToolBarEvent& event)
	{
		if (!event.IsDropDownClicked())
		{
			if (event.GetId() == ID_FILLCOLOR)
				ChangeCellBGColor(m_FillColor);
			else
				ChangeTextColor(m_FontColor);

			return;
		}

		wxColourDialog dlg(wxTheApp->GetTopWindow());
		dlg.SetExtraStyle(wxTAB_TRAVERSAL);

		if (dlg.ShowModal() != wxID_OK)
			return;
		
		const auto& Color = dlg.GetColourData().GetColour();
		if (event.GetId() == ID_FILLCOLOR)
		{
			m_FillColor = Color;
			ChangeCellBGColor(m_FillColor);
		}
		else
		{
			m_FontColor = Color;
			ChangeTextColor(m_FontColor);
		}
	}




	void CWorkbook::OnUndoRedoStackChanged(wxCommandEvent& event)
	{
		if (!GetUndoStack().empty())
		{
			auto top = GetUndoStack().top().get();
		}

		if (!GetRedoStack().empty())
		{
			auto top = GetRedoStack().top().get();
		}
	}


}