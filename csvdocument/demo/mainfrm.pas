unit mainfrm;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, LResources, Forms, Controls, Graphics, Dialogs,
  Grids, StdCtrls, ExtCtrls, CsvDocument;

type
  TfmCSVTest = class(TForm)
    btnSave: TButton;
    cbbDelimiter: TComboBox;
    Delimiter: TLabel;
    SaveDialog: TSaveDialog;
    sgView: TStringGrid;
    mmSource: TMemo;
    splTop: TSplitter;
    mmResult: TMemo;
    splBottom1: TSplitter;
    mmCellValue: TMemo;
    splBottom2: TSplitter;
    UpdateTimer: TIdleTimer;
    lblSource: TLabel;
    lblOutput: TLabel;
    lblCSVDoc: TLabel;
    lblCellContent: TLabel;
    pnButtons: TPanel;
    btnLoad: TButton;
    OpenDialog: TOpenDialog;
    procedure btnSaveClick(Sender: TObject);
    procedure cbbDelimiterChange(Sender: TObject);
    procedure mmSourceChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure sgViewSelectCell(Sender: TObject; aCol, aRow: Integer;
      var CanSelect: Boolean);
    procedure mmCellValueChange(Sender: TObject);
    procedure UpdateTimerTimer(Sender: TObject);
    procedure btnLoadClick(Sender: TObject);
  private
    FDoc: TCSVDocument;
    procedure UpdateView;
  public
    { public declarations }
  end; 

var
  fmCSVTest: TfmCSVTest;

implementation

{ TfmCSVTest }

procedure TfmCSVTest.mmSourceChange(Sender: TObject);
begin
  FDoc.CSVText := mmSource.Text;
  UpdateTimer.Enabled := True;
end;

procedure TfmCSVTest.btnSaveClick(Sender: TObject);
begin
  if SaveDialog.Execute then
    FDoc.SaveToFile(SaveDialog.FileName);
end;

procedure TfmCSVTest.cbbDelimiterChange(Sender: TObject);
begin
  FDoc.Delimiter := cbbDelimiter.Text[1];
  FDoc.CSVText := mmSource.Text;
  UpdateTimer.Enabled := True;
end;

procedure TfmCSVTest.FormCreate(Sender: TObject);
begin
  FDoc := TCSVDocument.Create;
  FDoc.Delimiter := ';';
end;

procedure TfmCSVTest.FormDestroy(Sender: TObject);
begin
  FreeAndNil(FDoc);
end;

procedure TfmCSVTest.sgViewSelectCell(Sender: TObject; aCol, aRow: Integer;
  var CanSelect: Boolean);
begin
  mmCellValue.OnChange := nil;
  mmCellValue.Text := FDoc.Cells[aCol, aRow];
  mmCellValue.OnChange := @mmCellValueChange;
end;

procedure TfmCSVTest.mmCellValueChange(Sender: TObject);
begin
  if not Assigned(FDoc) then Exit; // qt workaround
  FDoc.Cells[sgView.Col, sgView.Row] := mmCellValue.Text;
  UpdateTimer.Enabled := True;
end;

procedure TfmCSVTest.UpdateTimerTimer(Sender: TObject);
begin
  UpdateView;
  mmResult.Text := FDoc.CSVText;
  UpdateTimer.Enabled := False;
end;

procedure TfmCSVTest.btnLoadClick(Sender: TObject);
begin
  if OpenDialog.Execute then
    FDoc.LoadFromFile(OpenDialog.FileName);
end;

procedure TfmCSVTest.UpdateView;
var
  i, j: Integer;
begin
  sgView.BeginUpdate;
  try
    i := FDoc.RowCount;
    if sgView.RowCount <> i then
      sgView.RowCount := i;

    i := FDoc.MaxColCount;
    if sgView.ColCount <> i then
      sgView.ColCount := i;

    for i := 0 to FDoc.RowCount - 1 do
      for j := 0 to sgView.ColCount - 1 do
        sgView.Cells[j, i] := FDoc.Cells[j, i];
  finally
    sgView.EndUpdate;
  end;
end;

initialization
  {$I mainfrm.lrs}

end.
