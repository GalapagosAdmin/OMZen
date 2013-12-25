unit omz_main;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, ComCtrls,
  Buttons, ExtCtrls;

type

  { TForm1 }

  TForm1 = class(TForm)
    bbAddChildOU: TBitBtn;
    bbAddPosition: TBitBtn;
    bbAddEmployee1: TBitBtn;
    ImageList1: TImageList;
    LabeledEdit1: TLabeledEdit;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TreeView1: TTreeView;
    procedure bbAddChildOUClick(Sender: TObject);
    procedure bbAddEmployee1Click(Sender: TObject);
    procedure bbAddPositionClick(Sender: TObject);
    procedure LabeledEdit1DragOver(Sender, Source: TObject; X, Y: Integer;
      State: TDragState; var Accept: Boolean);
    procedure TreeView1Click(Sender: TObject);
  private
    { private declarations }
  public
    { public declarations }
  end;

var
  Form1: TForm1;

implementation

var
  HWMOUObjNum :LongInt;

{$R *.lfm}

{ TForm1 }

procedure TForm1.bbAddChildOUClick(Sender: TObject);
begin
   if TreeView1.Selected <> nil then
    begin
    with Treeview1.Items.AddChild(Treeview1.Selected ,'New OU') do
      begin
        ImageIndex := 0;
        SelectedIndex := 0;
      end;
    TreeView1.Selected.Expand(True);
//   TreeView1.Selected.StateIndex:=;
    end;
//   TreeView1.Selected.ImageIndex:=
//   TreeView1.Selected.SelectedIndex;
//   TreeView1.SaveToFile();
end;

Function GetNextOUObjNum:Integer;
  begin
     Result := HWMOUObjNum + 1;
     HWMOUObjNum := Result;
  end;

procedure TForm1.bbAddEmployee1Click(Sender: TObject);
begin
   if TreeView1.Selected <> nil then
    begin
    with Treeview1.Items.AddChildObject(Treeview1.Selected ,'New Employee', Pointer(GetNextOUObjNum)) do
      begin
        ImageIndex := 3;
        SelectedIndex := 3;
      end;
    TreeView1.Selected.Expand(True);
    end;
end;

procedure TForm1.bbAddPositionClick(Sender: TObject);
begin
   if TreeView1.Selected <> nil then
    begin
    with Treeview1.Items.AddChild(Treeview1.Selected ,'New Position') do
      begin
        ImageIndex := 1;
        SelectedIndex := 1;
      end;
    TreeView1.Selected.Expand(True);
    end;
end;


procedure TForm1.LabeledEdit1DragOver(Sender, Source: TObject; X, Y: Integer;
  State: TDragState; var Accept: Boolean);
begin
  Accept := true;
end;

procedure TForm1.TreeView1Click(Sender: TObject);
begin
   if TreeView1.Selected <> nil then
    begin
      labelededit1.text := IntToStr(TreeView1.Selected.ImageIndex);

      Case TreeView1.Selected.ImageIndex of
        0: begin  // Org. Unit
            bbAddChildOU.Enabled:=True;
            bbAddPosition.Enabled := True;
            bbAddEmployee1.Enabled:= False;
            labelededit1.text := 'Organizational Unit';
           end;
        1: begin // Normal Position
             bbAddChildOU.Enabled:=False;
             bbAddPosition.Enabled := False;
             bbAddEmployee1.Enabled:= True;
             labelededit1.text := 'Position';
           end;
        2: begin // Chief Position
             bbAddChildOU.Enabled:= False;
             bbAddPosition.Enabled := False;
             bbAddEmployee1.Enabled:= True;
             labelededit1.text := 'Chief Position (Manager)';
           end;
        3: begin // Employee
             bbAddChildOU.Enabled:= False;
             bbAddPosition.Enabled := False;
             bbAddEmployee1.Enabled:= False;
             labelededit1.text := 'Employee';
           end;
      end; // of case
    end;
end;

end.

