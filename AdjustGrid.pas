unit AdjustGrid;

interface

uses Windows, Forms, DBGrids, DB;

procedure AdjustColumnWidths(DBGrid: TDBGrid);

implementation

procedure AdjustColumnWidths(DBGrid: TDBGrid);
const
  C_Add = 3;
var
  ds: TDataSet;
  bm: TBookmark;
  i: Integer;
  w: Integer;
  a: array of Integer;
begin
  ds := DBGrid.DataSource.DataSet;

  if not Assigned(ds) then
    exit;

  if DBGrid.Columns.Count = 0 then
    exit;

  ds.DisableControls;
  bm := ds.GetBookmark;
  try
    ds.First;
    SetLength(a, DBGrid.Columns.Count);
    for i := 0 to DBGrid.Columns.Count - 1 do
      if Assigned(DBGrid.Columns[i].Field) then
        a[i] := DBGrid.Canvas.TextWidth(DBGrid.Columns[i].FieldName);

    while not ds.Eof do
    begin

      for i := 0 to DBGrid.Columns.Count - 1 do
      begin
        if not Assigned(DBGrid.Columns[i].Field) then
          continue;

        w := DBGrid.Canvas.TextWidth(ds.FieldByName(DBGrid.Columns[i].Field.FieldName).DisplayText);

        if a[i] < w then
          a[i] := w;
      end;
      ds.Next;
    end;

    w := 0;
    for i := 0 to DBGrid.Columns.Count - 1 do
    begin
      DBGrid.Columns[i].Width := a[i] + C_Add;
      inc(w, a[i] + C_Add);
    end;

    w := (DBGrid.ClientWidth - w - 20) div (DBGrid.Columns.Count);

    if w > 0 then
      for i := 0 to DBGrid.Columns.Count - 1 do
        DBGrid.Columns[i].Width := DBGrid.Columns[i].Width + w;


    ds.GotoBookmark(bm);
  finally
    ds.FreeBookmark(bm);
    ds.EnableControls;
  end;
end;

end.
