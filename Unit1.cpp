//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "Unit1.h"
#include "Report.cpp"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "cxButtons"
#pragma link "cxLookAndFeelPainters"
#pragma link "Access_2K_SRVR"
#pragma resource "*.dfm"
TForm1 *Form1;
String bin_file_path="";
String path_to_mdb="";
/*
 *Разделить строку Str String(AnsiString) на массив по разделителю Delimeter
*/
TStringList * ExplodeStr(String Str,String Delimeter) {
	TStringList *stringList = new TStringList;
    stringList->Clear();
	stringList->Delimiter = *Delimeter.c_str();
    stringList->DelimitedText=Str;
    return stringList;
}

// функция разбивает по-указанным разделителям, maked by xAtom from CyberForum
void split(TStringList* lout, char* str, const char* separator) {
      //СТРАННО РАБОТАЕТ - МЕНЯЕТ ВХОДНОЕ ЗНАЧЕНИЕ
      for(char* tok = strtok(str, separator); tok; tok = strtok(NULL, separator)) {
          lout->Add(tok);
      }
      return;
}

//ведение логов
void insert_into_log_file (String Error) {
    TDateTime CurrTime = Time().CurrentDateTime();
	String CTime = CurrTime.FormatString("yyyy-mm-dd hh:mm:ss");
    String programm_name= ParamStr(0);
    String CurrentDirLog = ExtractFilePath(programm_name)+"db_errors.log";
    const int length = CurrentDirLog.Length()+1;
    char *log_file = new char[length];
    strcpy(log_file, CurrentDirLog.c_str());
    fstream file;

    file.open(log_file, std::fstream::ate | std::fstream::out | std::fstream::app);
	if (!file ) {
      	file.open(log_file,  fstream::in | fstream::out | fstream::trunc);
        file <<"["<<CTime.c_str()<<"]:  "<<Error.c_str()<<endl;
        file.close();

       } 
    else {
    	file <<"["<<CTime.c_str()<<"]:  "<<Error.c_str()<<endl;
        file.close();
	}
	return;
}

//---------------------------------------------------------------------------
__fastcall TForm1::TForm1(TComponent* Owner)
	: TForm(Owner)
{
}
//---------------------------------------------------------------------------

void __fastcall TForm1::NChooseBinFileClick(TObject *Sender)
{
	if (ChooseBinFile->Execute() == true) {
    for (int i=0;i<Form1->ChooseBinFile->Files->Count;i++) {
        	bin_file_path = Form1->ChooseBinFile->Files->Strings[i];
        }
        Form1->EditDBFile->Text = bin_file_path;
    }
}
//---------------------------------------------------------------------------
void __fastcall TForm1::FormCreate(TObject *Sender)
{
    try {
		//название и путь к запускаемой программе
     	String programm_name= ParamStr(0);
     	String CurrentDir = ExtractFilePath(programm_name);
     	//ПОЛНОЕ ИМЯ INI ФАЙЛА С ПУТЁМ
     	String INI_NAME;
     	INI_NAME=CurrentDir+"database.ini";
     	//СОЗДАЁМ ОБЪЕКТ INWI
     	TIniFile *path_db = new TIniFile(INI_NAME);
     	//ПРИ НЕЗАПОЛНЕННЫХ СДЕЛАЕМ ИХ ЗНАЧЕНИЕ Undefined
     	//ЗАПИСЬ ИЗ INI В СТРОКУ ЗНАЧЕНИЯ ПЕРЕМЕННОЙ DB_PATH==путь к базе данных
     	path_to_mdb=path_db->ReadString("Configure","DB_PATH","Undefined");
     	//путь к БД не заполнен
     	if (AnsiCompareStr("Undefined",path_to_mdb)==0||path_to_mdb.IsEmpty()) {
     	MessageBox (NULL,"Не найдены или некорректно заполнены настройки файла database.ini.\n Программа будет завершена!\n Исправьте ошибку и запустите программу заново!","Ошибка конфигурации",MB_OK|MB_ICONERROR);
     		exit(1);
     	}
     	EditDBName->Text = path_to_mdb;
    }
    catch (Exception *ex) {
    	MessageBox (NULL,ex->Message.c_str(),"Ошибка!",MB_OK|MB_ICONERROR);
	}
}
//---------------------------------------------------------------------------


void __fastcall TForm1::CxBtnWriteFromFileToDBClick(TObject *Sender)
{
    if (!bin_file_path.IsEmpty()) {
        try {
    		ifstream is (bin_file_path.c_str(), ifstream::binary);
			is.seekg (0, is.end);
			int length = is.tellg();
            char * buffer = new char [length];
 			is.seekg (0, is.beg);
   			is.read (buffer, length);

            /*
            	предполагаемое число строк для записи,
                с учётом того, что в каждой строке 7 значений:
                Отдел
   				Должность
    			ФИО
  				Оклад
   				Личный автомобиль
  				Права
   				Паспорт

            */
            int row_count = ceil(length/91);//максимальная длина строки для записи - 91 символ
            /*
            	rows[0][0] Отдел
                rows[0][1] Должность
                rows[0][2] ИО
                rows[0][3] Оклад
                rows[0][4] Личный автомобиль
                rows[0][5] Права
                rows[0][6] Паспорт
            */
            vector < vector<char*> > rows(row_count);
            //до последней строки, не включая её
            int i_beg = 0;int i_end;
            /*
            	номер строки для записи.
                начинаем с нулевой.
                предполагается, что в одной строке 7 переменых для записи.
                все они разделены символами \0 (Null character)
            */
            int row_num = 0;
            int word_num = 0;//счётчик слов
            for (int i=0;i<length;i++) {
                if (buffer[i] == '\0'&&row_num<row_count) {
                	i_end = i;
                    char *bf = new char [i_end+1];
                    for (int j = 0;j<=i_end-i_beg;j++) {
                    	bf[j] = buffer[j+i_beg];
                    }
                    rows[row_num].push_back(bf);
					word_num++;
                    //переводим каретку дальше, если следующий - нулевой символ
                    if ((i+1)<length) {
                    	while (buffer[i + 1] == '\0') {
                    		i++;
                    	}
                    }	    
                    i_beg = i+1;
                }
                //запись последнего слова в последней стоке для записи
                if ( i + 1 == length) {
                	i_end = i;
                    char *bf = new char [i_end+1];
                    for (int j = 0;j<=i_end-i_beg;j++) {
                    	bf[j] = buffer[j+i_beg];
                    }
                    rows[row_num].push_back(bf);
                }

                if (i!= 0) {
                	//при достижении каждого 7-ого слова
                	if (word_num == 7) {
                		//переходим на новую строку записи
                		row_num++;
                        word_num = 0;
                	}
                }    
            }
            try {
            	String pr,dpr;
 				pr="MSDataShape.1";
 				dpr="Microsoft.Jet.OLEDB.4.0";
 				const String ConnStr = "Provider=%s;Data Provider=%s;Data Source=%s";
 				if (!ADOConnection1->Connected) {
  					ADOConnection1->ConnectionString = Format (ConnStr,ARRAYOFCONST((pr,dpr,path_to_mdb)));
                    ADOConnection1->LoginPrompt = False;

                }

                for (int i = 0;i<row_count;i++) {
                	if (rows[i].size()>0) {
                           /*
            				rows[0][0] Отдел
                            rows[0][1] Должность
                			rows[0][2] ФИО
                			rows[0][3] Оклад
                			rows[0][4] Личный автомобиль
                			rows[0][5] Права
                			rows[0][6] Паспорт
                           */
                    	String Department = rows[i][0];
                        //ShowMessage(Department);
                        ADOQuery1->SQL->Clear();
                        ADOQuery1->SQL->Text = "SELECT Id FROM Departments WHERE Department_n=:Department";
                        ADOQuery1->Parameters->ParamByName("Department")->Value=Department;
                        ADOQuery1->Open();
                        String Department_int;
                        if (ADOQuery1->FieldByName("Id")->Value.IsNull()) {
                        	ADOQuery1->Close();
                        	throw Exception("Отдела "+Department+" нет в справочнике отделов");
                        }
                        else {
							Department_int = ADOQuery1->FieldByName("Id")->Value;
                        }
                        ADOQuery1->Close();


                       String Position = rows[i][1];
                       ADOQuery1->SQL->Clear();
                       ADOQuery1->SQL->Text = "SELECT Id FROM Positions WHERE Position_n=:Position";
                       ADOQuery1->Parameters->ParamByName("Position")->Value=Position;
                       ADOQuery1->Open();
                       String Position_int;
                       if (ADOQuery1->FieldByName("Id")->Value.IsNull()) {
                       		ADOQuery1->Close();
                       		throw Exception("Должности "+Position+" нет в справочнике должностей");
                       }    
                       else {
                       		Position_int = ADOQuery1->FieldByName("Id")->Value;
                       }
                       ADOQuery1->Close();

                       ADOQuery1->SQL->Clear();
                       /*ADOQuery1->SQL->Add("INSERT INTO 'Employees'(FIO,Department,Position,Salary,Personal_vehicle,Driver_licence,Passport_s_n)" );
                       ADOQuery1->SQL->Add(" VALUES (:FIO, :Department,:Position,:Salary,:Personal_vehicle,:Driver_licence,:Passport_s_n)");
                       ADOQuery1->Parameters->ParamByName("FIO")->Value = "'"+(String)rows[i][2]+"'";
                       ADOQuery1->Parameters->ParamByName("Department") ->Value = Department_int;
                       ADOQuery1->Parameters->ParamByName("Position") ->Value = Position_int;
                       ADOQuery1->Parameters->ParamByName("Salary") ->Value = rows[i][3];
                       */
                       String p_vehicle, d_licence;

                       (String)rows[i][4]=="Да"?p_vehicle = "true":p_vehicle = "false";
                       (String)rows[i][5]=="Да"?d_licence = "true":d_licence = "false";
                       String txt;
                       txt = "INSERT INTO [Employees]([FIO],[Department],[Position],[Salary],[Personal_vehicle],[Driver_licence],[Passport_s_n]) VALUES('"+(String)rows[i][2]+"',"+Department_int+","+Position_int+","+rows[i][3]+","+p_vehicle+","+d_licence+",'"+rows[i][6]+"');";
                       ADOQuery1->SQL->Text = txt;
                       ADOQuery1->ExecSQL();
                       ADOQuery1->Close();
                       ShowMessage("Данные успешно записаны!");
                       }
                }
            }
            catch (Exception *ex) {
            	if (ADOQuery1->Active) {
                	ADOQuery1->Close();
                }
                insert_into_log_file(ex->Message);
            	//MessageBox (NULL,ex->Message.c_str(),"Ошибка!",MB_OK|MB_ICONERROR);
            }


        }
        catch (Exception *ex) {
        	insert_into_log_file(ex->Message);
    		MessageBox (NULL,ex->Message.c_str(),"Ошибка!",MB_OK|MB_ICONERROR);
		}
    }
    else {
    	MessageBox (NULL,"Не выбран файл для записи в базу данных.\nПожалуйста выберите файл (Настройки->Выбрать файл с данными) и повторите","Ошибка выбора BIN файла",MB_OK|MB_ICONERROR);

    }
}
//---------------------------------------------------------------------------

void __fastcall TForm1::FormClose(TObject *Sender, TCloseAction &Action)
{
	if (ADOConnection1->Connected) {
    	ADOConnection1->Close();
	}
}
//---------------------------------------------------------------------------

void __fastcall TForm1::cxBtnReportClick(TObject *Sender)
{
	TFormReports *Reports = new TFormReports(this);
    Reports->ShowModal();
}
//---------------------------------------------------------------------------

