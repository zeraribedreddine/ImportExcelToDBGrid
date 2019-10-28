program ExcelToDBGrid;

uses
  Forms,
  UMain in 'UMain.pas' {FrmMain},
  AdjustGrid in 'AdjustGrid.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TFrmMain, FrmMain);
  Application.Run;
end.
