#pragma once

#include <vector>
#include <list>
#include <map>
#include <set>
#include <filesystem>


#include <wx/wx.h>
#include <wx/aui/auibook.h>
#include <wx/aui/aui.h>
#include <wx/popupwin.h>
#include <wx/tglbtn.h>
#include <wx/colordlg.h>

#include <grid/worksheetbase.h>
#include <grid/ntbkbase.h>
#include <grid/workbookbase.h>





namespace grid
{
	class CWorkbookBase;
}



namespace ICELL
{
	class CWorksheetNtbk;
	class CWorkbook;

	wxDECLARE_EVENT(ssEVT_WB_PAGECHANGED, wxAuiNotebookEvent);
	wxDECLARE_EVENT(ssEVT_GRID_SELECTION_BEGUN, wxGridRangeSelectEvent);


	class CWorksheet :public grid::CWorksheetBase
	{
	public:
		CWorksheet(wxWindow* panel,
			CWorkbook* workbook,
			wxWindowID id,
			const wxPoint& pos,
			const wxSize& size,
			long style,
			wxString WindowName,
			int nrows = 1000,
			int ncols = 100);

		virtual ~CWorksheet();


		CWorkbook* GetWorkbook() const {
			return m_Workbook;
		}


		wxMenu* GetContextMenu() const
		{
			return m_ContextMenu;
		}


	protected:
		void OnRightClick(wxGridEvent& event);

		void OnSelectCell(wxGridEvent& event);

	protected:
		CWorkbook* m_Workbook = nullptr;
		wxMenu* m_ContextMenu{nullptr};

	private:
		wxWindow* m_ParentWnd;
	};


	/********************************************************* */

	class CWorksheetNtbk : public grid::CWorksheetNtbkBase
	{

	public:
		CWorksheetNtbk(CWorkbook* parent);
		virtual ~CWorksheetNtbk();

		grid::CWorksheetBase* CreateWorksheet(wxWindow* wnd, const std::wstring& Label, int nrows, int ncols) const;

		auto GetContextMenu() const
		{
			return m_ContextMenu;
		}


	protected:
		void OnPageChanged(wxAuiNotebookEvent& evt);
		void OnTabRightDown(wxAuiNotebookEvent& evt) override;

	private:
		CWorkbook* m_Workbook{nullptr};
	};



	/************************************************************** */


	class CWorkbook : public grid::CWorkbookBase
	{
	public:
		CWorkbook(wxWindow* parent);
		virtual ~CWorkbook();

		CWorksheetNtbk* GetWorksheetNotebook() const override
		{
			return (CWorksheetNtbk*)m_WSNtbk;
		}

		bool AddNewWorksheet(const std::wstring& tblname = L"", int nrows = 1000, int ncols = 50) override
		{
			return GetWorksheetNotebook()->AddNewWorksheet(tblname, nrows, ncols);
		}

		void ChangeCellAlignment(int ID)
		{
			grid::CWorkbookBase::ChangeCellAlignment(ID);
		}


	protected:

		void OnUpdateUI(wxUpdateUIEvent& evt);
		void OnFontChanged(wxCommandEvent& event); //Font size or face
		void OnFillFontColor(wxAuiToolBarEvent& event);
		void OnUndoRedoStackChanged(wxCommandEvent& event);
		

	private:
		wxComboBox* m_ComboFontFace, *m_ComboFontSize;
		wxColour m_FillColor, m_FontColor;

		const int ID_FILLCOLOR = wxNewId();
		const int ID_FONTCOLOR = wxNewId();
		const int ID_FONTBOLD = wxNewId();
		const int ID_FONTITALIC = wxNewId();
		const int ID_FONTSIZE = wxNewId();
		const int ID_FONTFACE = wxNewId();

		const int ID_PASTE = wxNewId();
		const int ID_UNDO = wxNewId();
		const int ID_REDO = wxNewId();
	};
}