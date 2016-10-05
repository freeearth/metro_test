//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop
#include <ComObj.hpp>
#include "Report.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "cxButtons"
#pragma link "cxLookAndFeelPainters"
#pragma resource "*.dfm"
TFormReports *FormReports;
//---------------------------------------------------------------------------
__fastcall TFormReports::TFormReports(TComponent* Owner)
	: TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TFormReports::FormCreate(TObject *Sender)
{
    try {
    	ComboBoxDepartments->Style = csDropDownList;
		String pr,dpr;
    	pr="MSDataShape.1";
    	dpr="Microsoft.Jet.OLEDB.4.0";
    	const String ConnStr = "Provider=%s;Data Provider=%s;Data Source=%s";
    	if (!ADOConnection2->Connected) {
    		ADOConnection2->ConnectionString = Format (ConnStr,ARRAYOFCONST((pr,dpr,path_to_mdb)));
        	ADOConnection2->LoginPrompt = False;
		}
        ADOQuery2->SQL->Clear();
        ADOQuery2->SQL->Text = "SELECT * FROM Departments";
        ADOQuery2->Open();
    	if (ADOQuery2->RecordCount == 0) {
        	MessageBox (NULL,"Справочник Отделы пустой\n Окно будет закрыто!\nВнесите несколько записей в справочник Отделы и повторите!","Ошибка",MB_OK|MB_ICONERROR);
            FormReports->Close();
        }
        for (int i = 1; i < ADOQuery2->RecordCount + 1; i++) {
        	String id = ADOQuery2->FieldByName("Id")->Value;
            String Department_n = ADOQuery2->FieldByName("Department_n")->Value;
            ComboBoxDepartments->AddItem(Department_n,(TObject*) new String(id));
            ADOQuery2->Next();
        }
        ComboBoxDepartments->ItemIndex = 0;
		ADOQuery2->Close();
    }
    catch (Exception *ex) {
    	if (ADOQuery2->Active) {
        	ADOQuery2->Close();
        }
    	MessageBox (NULL,ex->Message.c_str(),"Ошибка!",MB_OK|MB_ICONERROR);
	}
}
//---------------------------------------------------------------------------
void __fastcall TFormReports::cxButtonExitClick(TObject *Sender)
{
	Close();
}
//---------------------------------------------------------------------------

void __fastcall TFormReports::cxButtonMakeXlsReportClick(TObject *Sender)
{
    if (EditXlsName->Text.IsEmpty()) {
    		MessageBox (NULL,"Введите имя файла для отчёта","Ошибка xls!",MB_OK|MB_ICONERROR);
            return;
    }
    else {
    	try {
    		ADOQuery2->SQL->Clear();
        	ADOQuery2->SQL->Text = "SELECT * FROM Employees WHERE Department=:Department";
            String dep = * (String*)ComboBoxDepartments->Items->Objects[ComboBoxDepartments->ItemIndex];
        	ADOQuery2->Parameters->ParamByName("Department")->Value= dep;
        	ADOQuery2->Open();
            vector < vector<String> > report_data(ADOQuery2->RecordCount);
            int rows_num =  ADOQuery2->RecordCount;
            if (rows_num == 0) {
            	throw Exception("В базе данных нет записей о сотрудниках!");
            }
    		for (int i = 1; i < ADOQuery2->RecordCount + 1; i++) {
        		String FIO = ADOQuery2->FieldByName("FIO")->Value;
                String Department = ComboBoxDepartments->Text;

                String id_p = ADOQuery2->FieldByName("Position")->Value;
				ADOQuery3->SQL->Clear();
                ADOQuery3->SQL->Text = "SELECT Position_n FROM Positions WHERE Id=:Id_p";
                ADOQuery3->Parameters->ParamByName("Id_p")->Value=id_p;
                ADOQuery3->Open();
                String Position_n;
                if (ADOQuery3->FieldByName("Position_n")->Value.IsNull())  {
                	throw Exception("Должности №" +id_p +" не существует!");
                }
                else {
                	Position_n = ADOQuery3->FieldByName("Position_n")->Value;
                }
                ADOQuery3->Close();

                String Salary= ADOQuery2->FieldByName("Salary")->Value;
                String P_vehicle = ADOQuery2->FieldByName("Personal_vehicle")->Value;
                P_vehicle=="True"?P_vehicle = "Да":P_vehicle = "Нет";


            	String D_licence = ADOQuery2->FieldByName("Driver_licence")->Value;
                D_licence=="True"?D_licence = "Да":D_licence = "Нет";

                String Pasport_s_n = ADOQuery2->FieldByName("Passport_s_n")->Value;

                report_data[i-1].push_back(FIO);
                report_data[i-1].push_back(Department);
                report_data[i-1].push_back(Position_n);
                report_data[i-1].push_back(Salary);
                report_data[i-1].push_back(P_vehicle);
                report_data[i-1].push_back(D_licence);
                report_data[i-1].push_back(Pasport_s_n);
                ADOQuery2->Next();
        	}




            Variant v;
 			if(!fStart) {
				vVarApp=CreateOleObject("Excel.Application");
 				fStart=true;
    		}
 	
    		vVarApp.OlePropertySet("Visible",true);
 			vVarBooks=vVarApp.OlePropertyGet("Workbooks");
 			vVarApp.OlePropertySet("SheetsInNewWorkbook",1);
 			vVarBooks.OleProcedure("Add");
    		vVarBook=vVarBooks.OlePropertyGet("Item",1);
    		vVarSheets=vVarBook.OlePropertyGet("Worksheets") ;
 			vVarSheet=vVarSheets.OlePropertyGet("Item",1);
            String ListName = "Отчёт по отделу "+ComboBoxDepartments->Text;
 			vVarSheet.OlePropertySet("Name",ListName.c_str());
 			vVarSheet=vVarSheets.OlePropertyGet("Item",1);
 			vVarSheet.OleProcedure("Activate");
 
 			//Заносим номер дня в первую строку таблицы
  			vVarCell=vVarSheet.OlePropertyGet("Cells").
            OlePropertyGet("Item",4,1);
  			vBorder(vVarCell,2,1,55);
  			vFont(vVarCell,-4108,-4108,12,37,1,clBlack,0,0,0);
            vVarCell.OlePropertySet("Value","ФИО");
            vVarCell.OlePropertySet("ColumnWidth",40);
            vVarCell.OlePropertyGet("Font").OlePropertySet("Bold",true);

            	//Заносим номер дня в первую строку таблицы
  			vVarCell=vVarSheet.OlePropertyGet("Cells").
            OlePropertyGet("Item",5,1);
  			vBorder(vVarCell,2,1,55);
  			vFont(vVarCell,-4108,-4108,12,37,1,clBlack,0,0,0);
            vVarCell.OlePropertySet("Value","Отдел");
            vVarCell.OlePropertySet("ColumnWidth",40);
            vVarCell.OlePropertyGet("Font").OlePropertySet("Bold",true);


            	//Заносим номер дня в первую строку таблицы
  			vVarCell=vVarSheet.OlePropertyGet("Cells").
            OlePropertyGet("Item",6,1);
  			vBorder(vVarCell,2,1,55);
  			vFont(vVarCell,-4108,-4108,12,37,1,clBlack,0,0,0);
            vVarCell.OlePropertySet("Value","Должность");
            vVarCell.OlePropertySet("ColumnWidth",40);
            vVarCell.OlePropertyGet("Font").OlePropertySet("Bold",true);

            	//Заносим номер дня в первую строку таблицы
  			vVarCell=vVarSheet.OlePropertyGet("Cells").
            OlePropertyGet("Item",7,1);
  			vBorder(vVarCell,2,1,55);
  			vFont(vVarCell,-4108,-4108,12,37,1,clBlack,0,0,0);
            vVarCell.OlePropertySet("Value","Оклад");
            vVarCell.OlePropertySet("ColumnWidth",40);
            vVarCell.OlePropertyGet("Font").OlePropertySet("Bold",true);

            	//Заносим номер дня в первую строку таблицы
  			vVarCell=vVarSheet.OlePropertyGet("Cells").
            OlePropertyGet("Item",8,1);
  			vBorder(vVarCell,2,1,55);
  			vFont(vVarCell,-4108,-4108,12,37,1,clBlack,0,0,0);
            vVarCell.OlePropertySet("Value","Личный автомобиль");
            vVarCell.OlePropertySet("ColumnWidth",40);
            vVarCell.OlePropertyGet("Font").OlePropertySet("Bold",true);

            	//Заносим номер дня в первую строку таблицы
  			vVarCell=vVarSheet.OlePropertyGet("Cells").
            OlePropertyGet("Item",9,1);
  			vBorder(vVarCell,2,1,55);
  			vFont(vVarCell,-4108,-4108,12,37,1,clBlack,0,0,0);
            vVarCell.OlePropertySet("Value","Права автомобилиста");
            vVarCell.OlePropertySet("ColumnWidth",40);
            vVarCell.OlePropertyGet("Font").OlePropertySet("Bold",true);

            	//Заносим номер дня в первую строку таблицы
  			vVarCell=vVarSheet.OlePropertyGet("Cells").
            OlePropertyGet("Item",10,1);
  			vBorder(vVarCell,2,1,55);
  			vFont(vVarCell,-4108,-4108,12,37,1,clBlack,0,0,0);
            vVarCell.OlePropertySet("Value","Паспортные данные");
            vVarCell.OlePropertySet("ColumnWidth",40);
            vVarCell.OlePropertyGet("Font").OlePropertySet("Bold",true);

 	for (int i=2; i < rows_num+2; i++)
 	{

            //Заносим номер дня в первую строку таблицы
  			vVarCell=vVarSheet.OlePropertyGet("Cells").
            OlePropertyGet("Item",4,i);
  			vBorder(vVarCell,2,1,55);
  			vFont(vVarCell,-4108,-4108,12,37,1,clBlue,0,0,0);
            vVarCell.OlePropertySet("Value",report_data[i-2][0].c_str());
            vVarCell.OlePropertySet("ColumnWidth",40);

            	//Заносим номер дня в первую строку таблицы
  			vVarCell=vVarSheet.OlePropertyGet("Cells").
            OlePropertyGet("Item",5,i);
  			vBorder(vVarCell,2,1,55);
  			vFont(vVarCell,-4108,-4108,12,37,1,clRed,0,0,0);
            vVarCell.OlePropertySet("Value",report_data[i-2][1].c_str());
            vVarCell.OlePropertySet("ColumnWidth",40);

            	//Заносим номер дня в первую строку таблицы
  			vVarCell=vVarSheet.OlePropertyGet("Cells").
            OlePropertyGet("Item",6,i);
  			vBorder(vVarCell,2,1,55);
  			vFont(vVarCell,-4108,-4108,12,37,1,clYellow,0,0,0);
            vVarCell.OlePropertySet("Value",report_data[i-2][2].c_str());
            vVarCell.OlePropertySet("ColumnWidth",40);

            	//Заносим номер дня в первую строку таблицы
  			vVarCell=vVarSheet.OlePropertyGet("Cells").
            OlePropertyGet("Item",7,i);
  			vBorder(vVarCell,2,1,55);
  			vFont(vVarCell,-4108,-4108,12,37,1,clGreen,0,0,0);
            vVarCell.OlePropertySet("Value",report_data[i-2][3].c_str());
            vVarCell.OlePropertySet("ColumnWidth",40);

            	//Заносим номер дня в первую строку таблицы
  			vVarCell=vVarSheet.OlePropertyGet("Cells").
            OlePropertyGet("Item",8,i);
  			vBorder(vVarCell,2,1,55);
  			vFont(vVarCell,-4108,-4108,12,37,1,clLime,0,0,0);
            vVarCell.OlePropertySet("Value",report_data[i-2][4].c_str());
            vVarCell.OlePropertySet("ColumnWidth",40);

            	//Заносим номер дня в первую строку таблицы
  			vVarCell=vVarSheet.OlePropertyGet("Cells").
            OlePropertyGet("Item",9,i);
  			vBorder(vVarCell,2,1,55);
  			vFont(vVarCell,-4108,-4108,12,37,1,clPurple,0,0,0);
            vVarCell.OlePropertySet("Value",report_data[i-2][5].c_str());
            vVarCell.OlePropertySet("ColumnWidth",40);

            	//Заносим номер дня в первую строку таблицы
  			vVarCell=vVarSheet.OlePropertyGet("Cells").
            OlePropertyGet("Item",10,i);
  			vBorder(vVarCell,2,1,55);
  			vFont(vVarCell,-4108,-4108,12,37,1,clBlack,0,0,0);
            vVarCell.OlePropertySet("Value",report_data[i-2][6].c_str());
            vVarCell.OlePropertySet("ColumnWidth",40);

            
		}
		//Отключить вывод сообщений с вопросами типа "Заменить файл..."
 		vVarApp.OlePropertySet("DisplayAlerts",false);

 		String programm_name= ParamStr(0);
     	String CurrentDir = ExtractFilePath(programm_name);

 		String vAsCurDir1=CurrentDir + EditXlsName->Text+".xls";
 		vVarApp.OlePropertyGet("Workbooks").OlePropertyGet("Item",1).
        OleProcedure("SaveAs",vAsCurDir1.c_str());
        ShowMessage("Отчёт "+vAsCurDir1+" успешно сфомирован!");
 		//Закрыть открытое приложение Excel
 		//vVarApp.OleProcedure("Quit");
		}
    	catch (Exception *ex) {
    		if (ADOQuery2->Active) {
        		ADOQuery2->Close();
        	}
            if (ADOQuery3->Active) {
        		ADOQuery3->Close();
        	}
        	MessageBox (NULL,ex->Message.c_str(),"Ошибка!",MB_OK|MB_ICONERROR);
		}
    }


}
//---------------------------------------------------------------------------



//Функция определяет все параметры рисования квадратной рамки
//вокруг ячейки или группе выделенных ячеек
void __fastcall TFormReports::vBorder(Variant& vVarCell,int Weight,
                   int LineStyle,int ColorIndex)
{
 for(int i=8; i <= 10; i++)
 {
  switch(LineStyle)
  {
   case 1:
   case -4115:
   case 4:
   case 5:
   case -4118:
   case -4119:
   case 13:
   case -4142:
    vVarCell.OlePropertyGet("Borders",10).
            OlePropertySet("LineStyle",LineStyle);
   break;
   default:
    vVarCell.OlePropertyGet("Borders",i).
            OlePropertySet("LineStyle",1);
  }
  switch(Weight)
  {
   case 1:
   case -4138:
   case 2:
   case 4:
    vVarCell.OlePropertyGet("Borders",i).
             OlePropertySet("Weight",Weight);
   break;
   default:
    vVarCell.OlePropertyGet("Borders",i).
             OlePropertySet("Weight",1);
  }
  vVarCell.OlePropertyGet("Borders",i).
           OlePropertySet("ColorIndex",ColorIndex);
 }
}


//Функция определяет все параметры шрифта, заливку, подчеркивание
//и выравнивание текста в ячейках или группе выделенных ячеек
void __fastcall TFormReports::vFont(Variant& vVarCell,int HAlignment,
              int VAlignment,int Size,int ColorIndex,
              int Name,TColor Color,int Style,
              int Strikline,int Underline)
{
  //Выравнивание
  switch(HAlignment)
  {
   case -4108:
   case 7:
   case -4117:
   case 5:
   case 1:
   case -4130:
   case -4131:
   case -4152:
    vVarCell.OlePropertySet("HorizontalAlignment",HAlignment);
   break;
  }
  switch(VAlignment)
  {
   case -4108:
   case 7:
   case -4117:
   case 5:
   case 1:
   case -4130:
   case -4131:
   case -4152:
    vVarCell.OlePropertySet("VerticalAlignment",VAlignment);
   break;
  }
  //Размер шрифта
  vVarCell.OlePropertyGet("Font").
           OlePropertySet("Size",Size);
  //Цвет шрифта
  vVarCell.OlePropertyGet("Font").
           OlePropertySet("Color",Color);
  //Имя щрифта(Можно включаь сколько угодно)
  switch(Name)
  {
   case 1:
    vVarCell.OlePropertyGet("Font").
             OlePropertySet("Name","Arial");
   break;
   case 2:
    vVarCell.OlePropertyGet("Font").
             OlePropertySet("Name","Times New");
   break;
   default:
    vVarCell.OlePropertyGet("Font").
             OlePropertySet("Name","System");
  }
  //Заливка ячейки
  vVarCell.OlePropertyGet("Interior").
           OlePropertySet("ColorIndex",ColorIndex);
  //Стиль шрифта
  switch(Style)
  {
   case 1:
    vVarCell.OlePropertyGet("Font").OlePropertySet("Bold",true);
    vVarCell.OlePropertyGet("Font").OlePropertySet("Italic",false);
   break;
   case 2:
    vVarCell.OlePropertyGet("Font").OlePropertySet("Bold",false);
    vVarCell.OlePropertyGet("Font").OlePropertySet("Italic",true);
   break;
   case 3:
    vVarCell.OlePropertyGet("Font").OlePropertySet("Bold",true);
    vVarCell.OlePropertyGet("Font").OlePropertySet("Italic",true);
   break;
   default:
   vVarCell.OlePropertyGet("Font").OlePropertySet("Bold",false);
   vVarCell.OlePropertyGet("Font").OlePropertySet("Italic",false);
  }
  //Зачеркивание и индексы
  switch(Strikline)
  {
   case 1: //Зачеркнутый
    vVarCell.OlePropertyGet("Font").
             OlePropertySet("Strikethrough",true);
   break;
   case 2://Верхний индекс
    vVarCell.OlePropertyGet("Font").
             OlePropertySet("Superscript",true);

   break;
   case 3://Верхний индекс
    vVarCell.OlePropertyGet("Font").
             OlePropertySet("Subscript",true);

   break;
   case 4://Нижний индекс
    vVarCell.OlePropertyGet("Font").
             OlePropertySet("Subscript",true);

   break;
   case 5://Без линий
    vVarCell.OlePropertyGet("Font").
            OlePropertySet("OutlineFont",true);

   break;
   case 6://C тенью
    vVarCell.OlePropertyGet("Font").
             OlePropertySet("Shadow",true);

   break;
   default://Без линий
    vVarCell.OlePropertyGet("Font").
             OlePropertySet("OutlineFont",true);
  }
  //Подчеркивание
  switch(Underline)
  {
   case 2:
   case 4:
   case 5:
   case -4119:
    vVarCell.OlePropertyGet("Font").
             OlePropertySet("Underline",Underline);
   break;
  }
}
