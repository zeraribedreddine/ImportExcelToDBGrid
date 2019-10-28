object FrmMain: TFrmMain
  Left = 215
  Top = 125
  BiDiMode = bdRightToLeft
  Caption = 'Import Data Excel to DbGRID'
  ClientHeight = 453
  ClientWidth = 680
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Times New Roman'
  Font.Style = []
  OldCreateOrder = False
  ParentBiDiMode = False
  Position = poDesktopCenter
  PixelsPerInch = 96
  TextHeight = 14
  object pnl1: TPanel
    Left = 0
    Top = 0
    Width = 680
    Height = 57
    Align = alTop
    TabOrder = 0
    object Label3: TLabel
      Left = 183
      Top = 12
      Width = 338
      Height = 39
      Caption = #1575#1587#1578#1610#1585#1575#1583' '#1575#1604#1576#1610#1575#1606#1575#1578' '#1605#1606' Excel '#1573#1604#1609' DBGrid'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -33
      Font.Name = 'Arabic Typesetting'
      Font.Style = [fsBold]
      ParentFont = False
    end
  end
  object Dbgrd_Excel: TDBGrid
    Left = 0
    Top = 137
    Width = 680
    Height = 269
    Align = alClient
    DataSource = ds_Excel
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -19
    Font.Name = 'Times New Roman'
    Font.Style = []
    ParentFont = False
    TabOrder = 1
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'Times New Roman'
    TitleFont.Style = []
  end
  object pnl2: TPanel
    Left = 0
    Top = 406
    Width = 680
    Height = 47
    Align = alBottom
    TabOrder = 2
    object Btn_Exit: TButton
      Left = 589
      Top = 6
      Width = 75
      Height = 35
      Caption = #1582#1585#1608#1580
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -24
      Font.Name = 'Times New Roman'
      Font.Style = []
      ParentFont = False
      TabOrder = 0
      OnClick = Btn_ExitClick
    end
  end
  object pnl3: TPanel
    Left = 0
    Top = 57
    Width = 680
    Height = 80
    Align = alTop
    TabOrder = 3
    object Label1: TLabel
      Left = 17
      Top = 17
      Width = 26
      Height = 24
      Caption = 'File'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -21
      Font.Name = 'Arabic Typesetting'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object Label2: TLabel
      Left = 17
      Top = 47
      Width = 76
      Height = 23
      BiDiMode = bdLeftToRight
      Caption = 'Excel Sheet'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -20
      Font.Name = 'Arabic Typesetting'
      Font.Style = [fsBold]
      ParentBiDiMode = False
      ParentFont = False
    end
    object Edt_ExcelFilePath: TEdit
      Left = 49
      Top = 14
      Width = 534
      Height = 22
      TabOrder = 0
    end
    object Btn_OpenFile: TButton
      Left = 589
      Top = 4
      Width = 75
      Height = 32
      Cursor = crHandPoint
      Caption = #1578#1581#1600#1605#1600#1610#1600#1604
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -29
      Font.Name = 'Arabic Typesetting'
      Font.Style = []
      ParentFont = False
      TabOrder = 1
      OnClick = Btn_OpenFileClick
    end
    object CBSheet: TComboBox
      Left = 96
      Top = 44
      Width = 487
      Height = 22
      Style = csDropDownList
      TabOrder = 2
    end
    object Btn_LoadExcel: TButton
      Left = 589
      Top = 42
      Width = 75
      Height = 32
      Cursor = crHandPoint
      Caption = #1593#1585#1590
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -29
      Font.Name = 'Arabic Typesetting'
      Font.Style = []
      ParentFont = False
      TabOrder = 3
      OnClick = Btn_LoadExcelClick
    end
  end
  object AdoCon_Excel: TADOConnection
    CursorLocation = clUseServer
    LoginPrompt = False
    Mode = cmShareDenyNone
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 120
    Top = 252
  end
  object AdoQuery_Excel: TADOQuery
    Connection = AdoCon_Excel
    Parameters = <>
    Left = 225
    Top = 178
  end
  object DlgOpen_Excel: TOpenDialog
    Filter = 'Excel files|*.XLS;*.XL'
    Left = 342
    Top = 172
  end
  object ds_Excel: TDataSource
    DataSet = AdoQuery_Excel
    Left = 117
    Top = 170
  end
end
