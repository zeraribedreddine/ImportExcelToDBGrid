unit UMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Grids, DBGrids, ExtCtrls, DB, ADODB;

type
  TFrmMain = class(TForm)
    pnl1: TPanel;
    Dbgrd_Excel: TDBGrid;
    pnl2: TPanel;
    pnl3: TPanel;
    Edt_ExcelFilePath: TEdit;
    Btn_OpenFile: TButton;
    CBSheet: TComboBox;
    Btn_LoadExcel: TButton;
    Label1: TLabel;
    Label2: TLabel;
    Btn_Exit: TButton;
    Label3: TLabel;
    AdoCon_Excel: TADOConnection;
    AdoQuery_Excel: TADOQuery;
    DlgOpen_Excel: TOpenDialog;
    ds_Excel: TDataSource;
    procedure Btn_ExitClick(Sender: TObject);
    procedure Btn_OpenFileClick(Sender: TObject);
    procedure Btn_LoadExcelClick(Sender: TObject);
  private
    { Private declarations }
    procedure Connect_To_Excel;
    procedure Get_Excel_Data;
  public
    { Public declarations }
  end;

var
  FrmMain: TFrmMain;

implementation
  uses
     Vcl.OleAuto, AdjustGrid, System.UITypes;
{$R *.dfm}

procedure TFrmMain.Connect_To_Excel;
var
  strConn:  widestring;
begin
  strConn:='Provider=Microsoft.Jet.OLEDB.4.0;' +
           'Data Source=' + Edt_ExcelFilePath.Text + ';' +
           'Extended Properties=Excel 8.0;';

  AdoCon_Excel.Connected:=False;
  AdoCon_Excel.ConnectionString := strConn;
  try
    AdoCon_Excel.Open;
    AdoCon_Excel.GetTableNames(CBSheet.Items,True);
  except
    MessageDlg('Failed To Connect to Excel File !!',mtWarning,[mbok],0);
  end;
end;

procedure TFrmMain.Get_Excel_Data;
begin
  if not AdoCon_Excel.Connected then Connect_To_Excel;
  AdoQuery_Excel.Close;
  AdoQuery_Excel.SQL.Text:='SELECT * FROM ['+CBSheet.Text+']';
  try
    AdoQuery_Excel.Open;
  except
    MessageDlg('Failed To Connect to Excel File !!',mtWarning,[mbok],0);
  end;
end;

procedure TFrmMain.Btn_ExitClick(Sender: TObject);
begin
  Close;
end;

procedure TFrmMain.Btn_OpenFileClick(Sender: TObject);
begin
  DlgOpen_Excel.InitialDir := ExtractFileDir(ParamStr(0));

  if DlgOpen_Excel.Execute then
  begin
    Edt_ExcelFilePath.Text := DlgOpen_Excel.FileName;
    Connect_To_Excel;
    CBSheet.ItemIndex := 0;
    Btn_LoadExcel.Enabled := True;
  end else ShowMessage('please choose a Simple Excel File to load it here');
end;

procedure TFrmMain.Btn_LoadExcelClick(Sender: TObject);
begin
  Get_Excel_Data; AdjustColumnWidths(Dbgrd_Excel);
end;

end.
