#pragma once

#include <string>
#include <wx/wx.h>
#include <wx/grid.h>
#include <wx/xml/xml.h>

#include "dllimpexp.h"



namespace grid
{
	class Cell;

	class DLLGRID CellFormat
	{
	public:
		CellFormat() = default;
		CellFormat(Cell* cell);

		bool operator==(const CellFormat& other) const;

		void SetFont(const wxFont& font) {
			m_Font = font;
		}

		wxFont GetFont() const {
			return m_Font;
		}

		void SetBackgroundColor(const wxColor& color) {
			m_BGColor = color;
		}

		wxColor GetBackgroundColor() const {
			return m_BGColor;
		}

		void SetTextColor(const wxColor& color) {
			m_TextColor = color;
		}

		wxColor GetTextColor() const
		{
			return m_TextColor;
		}


		void SetAlignment(int horiz, int vert);

		int GetHAlign() const {
			return m_HorAlign;
		}

		int GetVAlign() const {
			return m_VerAlign;
		}

	private:
		wxFont m_Font{ wxNullFont };
		wxColor m_BGColor{ wxNullColour }, m_TextColor{ wxNullColour };
		int m_HorAlign{ 0 }, m_VerAlign{ 0 };

	};




	class DLLGRID Cell
	{

	public:
		static Cell FromXMLNode(const wxGrid* const ws, const wxXmlNode* CellNode);

		//Get topleft and bottom right coordinates
		static std::pair<wxGridCellCoords, wxGridCellCoords>
			Get_TLBR(const std::vector<Cell>& vecCells);

	public:
		Cell() = default;
		Cell(const wxGrid* ws, int row = -1, int col = -1);

		bool operator==(const Cell& other) const;

		const wxGrid* GetWorksheetBase() const
		{
			return m_WSBase;
		}

		int GetRow() const {
			return m_Row;
		}

		void SetRow(int row) {
			m_Row = row;
		}

		int GetCol() const {
			return m_Column;
		}

		void SetCol(int col) {
			m_Column = col;
		}

		std::wstring GetValue() const {
			return m_Value;
		}

		void SetValue(const std::wstring& val) {
			m_Value = val;
		}

		CellFormat GetFormat() const {
			return m_Format;
		}

		void SetFormat(const CellFormat& format) {
			m_Format = format;
		}


		CellFormat GetDefaultFormat() const;
		std::wstring ToXMLString() const;

	private:
		int m_Row{ -1 }, m_Column{ -1 };
		std::wstring m_Value{};

		const wxGrid* m_WSBase{ nullptr };
		CellFormat m_Format;
	};
}