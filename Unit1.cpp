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
 *��������� ������ Str String(AnsiString) �� ������ �� ����������� Delimeter
*/
TStringList * ExplodeStr(String Str,String Delimeter) {
	TStringList *stringList = new TStringList;
    stringList->Clear();
	stringList->Delimiter = *Delimeter.c_str();
    stringList->DelimitedText=Str;
    return stringList;
}

// ������� ��������� ��-��������� ������������, maked by xAtom from CyberForum
void split(TStringList* lout, char* str, const char* separator) {
      //������� �������� - ������ ������� ��������
      for(char* tok = strtok(str, separator); tok; tok = strtok(NULL, separator)) {
          lout->Add(tok);
      }
      return;
}

//������� �����
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
		//�������� � ���� � ����������� ���������
     	String programm_name= ParamStr(0);
     	String CurrentDir = ExtractFilePath(programm_name);
     	//������ ��� INI ����� � ��Ҩ�
     	String INI_NAME;
     	INI_NAME=CurrentDir+"database.ini";
     	//������� ������ INWI
     	TIniFile *path_db = new TIniFile(INI_NAME);
     	//��� ������������� ������� �� �������� Undefined
     	//������ �� INI � ������ �������� ���������� DB_PATH==���� � ���� ������
     	path_to_mdb=path_db->ReadString("Configure","DB_PATH","Undefined");
     	//���� � �� �� ��������
     	if (AnsiCompareStr("Undefined",path_to_mdb)==0||path_to_mdb.IsEmpty()) {
     	MessageBox (NULL,"�� ������� ��� ����������� ��������� ��������� ����� database.ini.\n ��������� ����� ���������!\n ��������� ������ � ��������� ��������� ������!","������ ������������",MB_OK|MB_ICONERROR);
     		exit(1);
     	}
     	EditDBName->Text = path_to_mdb;
    }
    catch (Exception *ex) {
    	MessageBox (NULL,ex->Message.c_str(),"������!",MB_OK|MB_ICONERROR);
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
            	�������������� ����� ����� ��� ������,
                � ������ ����, ��� � ������ ������ 7 ��������:
                �����
   				���������
    			���
  				�����
   				������ ����������
  				�����
   				�������

            */
            int row_count = ceil(length/91);//������������ ����� ������ ��� ������ - 91 ������
            /*
            	rows[0][0] �����
                rows[0][1] ���������
                rows[0][2] ��
                rows[0][3] �����
                rows[0][4] ������ ����������
                rows[0][5] �����
                rows[0][6] �������
            */
            vector < vector<char*> > rows(row_count);
            //�� ��������� ������, �� ������� �
            int i_beg = 0;int i_end;
            /*
            	����� ������ ��� ������.
                �������� � �������.
                ��������������, ��� � ����� ������ 7 ��������� ��� ������.
                ��� ��� ��������� ��������� \0 (Null character)
            */
            int row_num = 0;
            int word_num = 0;//������� ����
            for (int i=0;i<length;i++) {
                if (buffer[i] == '\0'&&row_num<row_count) {
                	i_end = i;
                    char *bf = new char [i_end+1];
                    for (int j = 0;j<=i_end-i_beg;j++) {
                    	bf[j] = buffer[j+i_beg];
                    }
                    rows[row_num].push_back(bf);
					word_num++;
                    //��������� ������� ������, ���� ��������� - ������� ������
                    if ((i+1)<length) {
                    	while (buffer[i + 1] == '\0') {
                    		i++;
                    	}
                    }	    
                    i_beg = i+1;
                }
                //������ ���������� ����� � ��������� ����� ��� ������
                if ( i + 1 == length) {
                	i_end = i;
                    char *bf = new char [i_end+1];
                    for (int j = 0;j<=i_end-i_beg;j++) {
                    	bf[j] = buffer[j+i_beg];
                    }
                    rows[row_num].push_back(bf);
                }

                if (i!= 0) {
                	//��� ���������� ������� 7-��� �����
                	if (word_num == 7) {
                		//��������� �� ����� ������ ������
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
            				rows[0][0] �����
                            rows[0][1] ���������
                			rows[0][2] ���
                			rows[0][3] �����
                			rows[0][4] ������ ����������
                			rows[0][5] �����
                			rows[0][6] �������
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
                        	throw Exception("������ "+Department+" ��� � ����������� �������");
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
                       		throw Exception("��������� "+Position+" ��� � ����������� ����������");
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

                       (String)rows[i][4]=="��"?p_vehicle = "true":p_vehicle = "false";
                       (String)rows[i][5]=="��"?d_licence = "true":d_licence = "false";
                       String txt;
                       txt = "INSERT INTO [Employees]([FIO],[Department],[Position],[Salary],[Personal_vehicle],[Driver_licence],[Passport_s_n]) VALUES('"+(String)rows[i][2]+"',"+Department_int+","+Position_int+","+rows[i][3]+","+p_vehicle+","+d_licence+",'"+rows[i][6]+"');";
                       ADOQuery1->SQL->Text = txt;
                       ADOQuery1->ExecSQL();
                       ADOQuery1->Close();
                       ShowMessage("������ ������� ��������!");
                       }
                }
            }
            catch (Exception *ex) {
            	if (ADOQuery1->Active) {
                	ADOQuery1->Close();
                }
                insert_into_log_file(ex->Message);
            	//MessageBox (NULL,ex->Message.c_str(),"������!",MB_OK|MB_ICONERROR);
            }


        }
        catch (Exception *ex) {
        	insert_into_log_file(ex->Message);
    		MessageBox (NULL,ex->Message.c_str(),"������!",MB_OK|MB_ICONERROR);
		}
    }
    else {
    	MessageBox (NULL,"�� ������ ���� ��� ������ � ���� ������.\n���������� �������� ���� (���������->������� ���� � �������) � ���������","������ ������ BIN �����",MB_OK|MB_ICONERROR);

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

