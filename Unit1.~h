//---------------------------------------------------------------------------

#ifndef Unit1H
#define Unit1H
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <Menus.hpp>
#include <Dialogs.hpp>
//ДЛЯ РАБОТЫ С INI ФАЙЛАМИ
#include <IniFiles.hpp>
#include "cxButtons.hpp"
#include "cxLookAndFeelPainters.hpp"
#include <vector.h>
#include <math.h>
#include <ADODB.hpp>
#include <DB.hpp>
#include <DBGrids.hpp>
#include <Grids.hpp>

#include <cstdlib>
#include <iostream>
#include <fstream>
#include <iomanip>

#include <stdio.h>
#include <windows.h>
#include <winbase.h>
#include <conio>
//ЗАГОЛОВОЧНЫЙ ФАЙЛ ДЛЯ ОПРЕДЕЛЕНИЯ ПРИСУТСТВИЯ ЗАПУЩЕННОГО ПРОЦЕССА В СИСТЕМЕ
#include <TlHelp32.h>
using namespace std;


//---------------------------------------------------------------------------
class TForm1 : public TForm
{
__published:	// IDE-managed Components
	TMainMenu *MainMenu1;
	TMenuItem *N1;
	TMenuItem *NChooseBinFile;
	TOpenDialog *ChooseBinFile;
	TLabel *Label1;
	TEdit *EditDBFile;
	TLabel *Label2;
	TEdit *EditDBName;
	TcxButton *CxBtnWriteFromFileToDB;
	TADOConnection *ADOConnection1;
	TADOQuery *ADOQuery1;
	TDataSource *DataSource1;
	TcxButton *cxBtnReport;
	void __fastcall NChooseBinFileClick(TObject *Sender);
	void __fastcall FormCreate(TObject *Sender);
	void __fastcall CxBtnWriteFromFileToDBClick(TObject *Sender);
	void __fastcall FormClose(TObject *Sender, TCloseAction &Action);
	void __fastcall cxBtnReportClick(TObject *Sender);
private:	// User declarations
public:		// User declarations
	__fastcall TForm1(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TForm1 *Form1;

//Переменная с расположением и названием выбранного bin файла, содержащего данные для БД
extern String bin_file_path;
//Переменная с расположением и названием БД, из ini файла
extern String path_to_mdb;

//---------------------------------------------------------------------------
#endif
