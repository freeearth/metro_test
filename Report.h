//---------------------------------------------------------------------------

#ifndef ReportH
#define ReportH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <ADODB.hpp>
#include <DB.hpp>
#include <DBCtrls.hpp>
#include "cxButtons.hpp"
#include "cxLookAndFeelPainters.hpp"
#include <Menus.hpp>
#include <vector.h>
//---------------------------------------------------------------------------
class TFormReports : public TForm
{
__published:	// IDE-managed Components
	TLabel *Label2;
	TADOConnection *ADOConnection2;
	TADOQuery *ADOQuery2;
	TComboBox *ComboBoxDepartments;
	TcxButton *cxButtonExit;
	TcxButton *cxButtonMakeXlsReport;
	TLabel *Label1;
	TEdit *EditXlsName;
	TADOQuery *ADOQuery3;
	void __fastcall FormCreate(TObject *Sender);
	void __fastcall cxButtonExitClick(TObject *Sender);
	void __fastcall cxButtonMakeXlsReportClick(TObject *Sender);
private:	// User declarations
public:		// User declarations
	__fastcall TFormReports(TComponent* Owner);
    Variant vVarApp,vVarBooks,vVarBook,vVarSheets,vVarSheet,vVarCells,vVarCell;
    bool fStart;
    AnsiString    vAsCurDir;
    void __fastcall vBorder(Variant& vVarCell,int Weight,
                 int LineStyle,int ColorIndex);
  	void __fastcall vFont(Variant& vVarCell,int HAlignment,
                        int VAlignment,int Size,
                        int ColorIndex,int Name,
                        TColor Color,int Style,
                        int Strikline,int Underline);
};
//---------------------------------------------------------------------------
extern PACKAGE TFormReports *FormReports;
extern String path_to_mdb;
//---------------------------------------------------------------------------
#endif
