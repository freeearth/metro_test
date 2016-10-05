object Form1: TForm1
  Left = 379
  Top = 145
  BorderStyle = bsToolWindow
  Caption = 'TST'
  ClientHeight = 354
  ClientWidth = 779
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -16
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  Menu = MainMenu1
  OldCreateOrder = False
  OnClose = FormClose
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 20
  object Label1: TLabel
    Left = 8
    Top = 40
    Width = 128
    Height = 13
    Caption = 'BIN '#1092#1072#1081#1083' '#1089' '#1076#1072#1085#1085#1099#1084#1080':'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object Label2: TLabel
    Left = 8
    Top = 0
    Width = 77
    Height = 13
    Caption = #1041#1072#1079#1072' '#1076#1072#1085#1085#1099#1093
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object EditDBFile: TEdit
    Left = 136
    Top = 40
    Width = 537
    Height = 21
    AutoSelect = False
    DragMode = dmAutomatic
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ImeMode = imSKata
    ParentFont = False
    ReadOnly = True
    TabOrder = 0
  end
  object EditDBName: TEdit
    Left = 136
    Top = 0
    Width = 537
    Height = 21
    AutoSelect = False
    DragMode = dmAutomatic
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ImeMode = imSKata
    ParentFont = False
    ReadOnly = True
    TabOrder = 1
  end
  object CxBtnWriteFromFileToDB: TcxButton
    Left = 8
    Top = 96
    Width = 233
    Height = 25
    Caption = #1047#1072#1087#1080#1089#1072#1090#1100' '#1076#1072#1085#1085#1099#1077' '#1080#1079' '#1092#1072#1081#1083#1072' '#1074' '#1041#1044
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 2
    OnClick = CxBtnWriteFromFileToDBClick
  end
  object cxBtnReport: TcxButton
    Left = 288
    Top = 96
    Width = 233
    Height = 25
    Caption = #1042#1099#1075#1088#1091#1079#1080#1090#1100' '#1086#1090#1095#1105#1090' '#1074' Excel'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 3
    OnClick = cxBtnReportClick
  end
  object MainMenu1: TMainMenu
    Top = 152
    object N1: TMenuItem
      Caption = #1053#1072#1089#1090#1088#1086#1081#1082#1080
      object NChooseBinFile: TMenuItem
        Caption = #1042#1099#1073#1088#1072#1090#1100' '#1092#1072#1081#1083' '#1089' '#1076#1072#1085#1085#1099#1084#1080
        OnClick = NChooseBinFileClick
      end
    end
  end
  object ChooseBinFile: TOpenDialog
    Filter = 'bin db file|*.bin'
    Left = 32
    Top = 152
  end
  object ADOConnection1: TADOConnection
    Left = 88
    Top = 152
  end
  object ADOQuery1: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      '')
    Left = 136
    Top = 152
  end
  object DataSource1: TDataSource
    DataSet = ADOQuery1
    Left = 184
    Top = 152
  end
end
