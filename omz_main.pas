unit omz_main;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, ComCtrls,
  Buttons, ExtCtrls, StdCtrls, Menus;

type

  { TForm1 }

  TForm1 = class(TForm)
    bbAddChildOU: TBitBtn;
    bbAddEmployee2: TBitBtn;
    bbAddPosition: TBitBtn;
    bbAddEmployee1: TBitBtn;
    CheckBox1: TCheckBox;
    ImageList1: TImageList;
    LabeledEdit1: TLabeledEdit;
    LabeledEdit2: TLabeledEdit;
    MainMenu1: TMainMenu;
    MenuItem1: TMenuItem;
    MenuItem2: TMenuItem;
    MenuItem3: TMenuItem;
    MenuItem4: TMenuItem;
    miProcess: TMenuItem;
    OpenDialog1: TOpenDialog;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    tvOrgChart: TTreeView;
    procedure bbAddChildOUClick(Sender: TObject);
    procedure bbAddEmployee1Click(Sender: TObject);
    procedure bbAddEmployee2Click(Sender: TObject);
    procedure bbAddPositionClick(Sender: TObject);
    procedure CheckBox1Change(Sender: TObject);
    procedure LabeledEdit1DragOver(Sender, Source: TObject; X, Y: Integer;
      State: TDragState; var Accept: Boolean);
    procedure MenuItem1Click(Sender: TObject);
    procedure MenuItem2Click(Sender: TObject);
    procedure MenuItem4Click(Sender: TObject);
    procedure miProcessClick(Sender: TObject);
    procedure tvOrgChartClick(Sender: TObject);
    procedure tvOrgChartDragOver(Sender, Source: TObject; X, Y: Integer;
      State: TDragState; var Accept: Boolean);
  private
    { private declarations }
  public
    { public declarations }
  end;

var
  Form1: TForm1;

implementation


uses
  omz_logic, dbugintf;

{$R *.lfm}

{ TForm1 }

procedure TForm1.bbAddChildOUClick(Sender: TObject);
begin
   if tvOrgChart.Selected <> nil then
    begin
     with tvOrgChart.Items.AddChildObject(tvOrgChart.Selected ,'New Org. Unit', Pointer(GetNextOUObjNum)) do
      begin
        ImageIndex := IDX_ORG_UNIT;
        SelectedIndex := IDX_ORG_UNIT;
      end;
    tvOrgChart.Selected.Expand(True);
//   tvOrgChart.Selected.StateIndex:=;
    end;
//   tvOrgChart.Selected.ImageIndex:=
//   tvOrgChart.Selected.SelectedIndex;
//   tvOrgChart.SaveToFile();
end;



procedure TForm1.bbAddEmployee1Click(Sender: TObject);
begin
   if tvOrgChart.Selected <> nil then
    begin
    with tvOrgChart.Items.AddChildObject(tvOrgChart.Selected ,'New Employee', Pointer(0)) do
      begin
        ImageIndex := IDX_EMPLOYEE;
        SelectedIndex := IDX_EMPLOYEE;
      end;
    tvOrgChart.Selected.Expand(True);
    end;
end;

procedure TForm1.bbAddEmployee2Click(Sender: TObject);

  Procedure DeleteNode(Node:TTreeNode);
    begin
      while Node.HasChildren do DeleteNode(node.GetLastChild);
      tvOrgChart.Items.Delete(Node) ;
    end;
begin
   if tvOrgChart.Selected <> nil then
     begin
       if messagedlg('Delete object and all children ?',mtConfirmation,
                 [mbYes,mbNo],0) <> mrYes then exit;
       tvOrgChart.Selected.Delete;
     end;
end;

procedure TForm1.bbAddPositionClick(Sender: TObject);
begin
   if tvOrgChart.Selected <> nil then
    begin
    with tvOrgChart.Items.AddChildObject(tvOrgChart.Selected ,'New Position', Pointer(GetNextPosObjNum)) do
      begin
        ImageIndex := IDX_POSITION;
        SelectedIndex := IDX_POSITION;
      end;
    tvOrgChart.Selected.Expand(True);
    end;
end;

procedure TForm1.CheckBox1Change(Sender: TObject);
begin
   if tvOrgChart.Selected <> nil then
     begin
      Case tvOrgChart.Selected.ImageIndex of
        IDX_ORG_UNIT: begin  // Org. Unit
           end;
        IDX_POSITION: begin // Normal Position
          If self.CheckBox1.Checked then
           tvOrgChart.Selected.ImageIndex:=IDX_CHIEF_POSITION;
        end;
        IDX_CHIEF_POSITION: begin // Chief Position
          If not self.CheckBox1.Checked then
           tvOrgChart.Selected.ImageIndex:=IDX_POSITION;
           end;
        IDX_EMPLOYEE: begin // Employee
           end;
      end; // of case
      // Sync the icons for selected and non-selected case
      with tvOrgChart.Selected do
        SelectedIndex:= ImageIndex;

     end;
end;


procedure TForm1.LabeledEdit1DragOver(Sender, Source: TObject; X, Y: Integer;
  State: TDragState; var Accept: Boolean);
begin
  Accept := true;
end;

procedure TForm1.MenuItem1Click(Sender: TObject);
begin

end;

procedure TForm1.MenuItem2Click(Sender: TObject);
begin
  if OpenDialog1.Execute then
    begin
      Import_ADP_HRP1001(OpenDialog1.FileName);
    end;
end;

procedure TForm1.MenuItem4Click(Sender: TObject);
begin
    if OpenDialog1.Execute then
    begin
      Import_ADP_HRP1000(OpenDialog1.FileName);
    end;
end;

procedure TForm1.miProcessClick(Sender: TObject);
var
  h, i,j:longint;
  ParentNode:TTreeNode;
begin
  tvOrgChart.items.clear;
  // This transfers parent ID relationship from IT1001 to IT1000 structure.
  UpdateHasParent;
 for h := 1 to 10 do // 10 rounds = max 10 levels org. chart heirarchy
   // Loop through all objects
   for i := low(ObjectList) to high(ObjectList) do
     // We are only interested in Org. Units for now, and ones we haven't updated yet.
    if (ObjectList[i].ObjType = 'O') and (ObjectList[i].Stale) then  // Org. Unit
    begin
     case ObjectList[i].HasParent of
      False: begin // Has no parent, this is a root OU
       with
       tvOrgChart.Items.AddChildObject(nil,               // Parent ID
                  ObjectList[i].LongText,                 // Long Text
                  Pointer(ObjectList[i].ObjNum)) do       // Object Number
                    begin
                      ImageIndex := IDX_ORG_UNIT;
                      SelectedIndex := IDX_ORG_UNIT;
                    end;
       ObjectList[i].Stale := False;
      end;
      True:begin   // Has a parent, we should put this under an existing OU
             // Search for the parent
             ParentNode := tvOrgChart.Items.FindNodeWithData(Pointer(ObjectList[i].ParentObjID));
             if ParentNode = nil then
               begin  // We couldn't find the parent on the Org. Chart
                SendDebug('Skipping OU with missing parent: ' + IntToStr(ObjectList[i].ObjNum));
               end
             else // Parent found
              begin   // We found the parent
             with
                tvOrgChart.Items.AddChildObject(ParentNode,       // Parent ID
                        ObjectList[i].LongText,                 // Long Text
                        Pointer(ObjectList[i].ObjNum)) do       // Object Number
                          begin
                            ImageIndex := IDX_ORG_UNIT;
                            SelectedIndex := IDX_ORG_UNIT;
                          end;
                ObjectList[i].Stale := False;
              end; // parent found
           end; // of TRUE
     end;
    end;

end;

procedure TForm1.tvOrgChartClick(Sender: TObject);
begin
   if tvOrgChart.Selected <> nil then
    begin
      labelededit1.text := IntToStr(tvOrgChart.Selected.ImageIndex);
      labelededit2.text := IntToStr(Integer(tvOrgChart.Selected.Data));

      Case tvOrgChart.Selected.ImageIndex of
        IDX_ORG_UNIT: begin  // Org. Unit
            bbAddChildOU.Enabled:=True;
            bbAddPosition.Enabled := True;
            bbAddEmployee1.Enabled:= False;
            labelededit1.text := 'Organizational Unit';
            CheckBox1.Checked:=False;
            CheckBox1.Enabled:=False;
           end;
        IDX_POSITION: begin // Normal Position
             bbAddChildOU.Enabled:=False;
             bbAddPosition.Enabled := False;
             bbAddEmployee1.Enabled:= True;
             labelededit1.text := 'Position';
             CheckBox1.Enabled:=True;
             CheckBox1.Checked:=False;
           end;
        IDX_CHIEF_POSITION: begin // Chief Position
             bbAddChildOU.Enabled:= False;
             bbAddPosition.Enabled := False;
             bbAddEmployee1.Enabled:= True;
             labelededit1.text := 'Chief Position (Manager)';
             CheckBox1.Enabled:=True;
             CheckBox1.Checked:=True;
           end;
        IDX_EMPLOYEE: begin // Employee
             bbAddChildOU.Enabled:= False;
             bbAddPosition.Enabled := False;
             bbAddEmployee1.Enabled:= False;
             labelededit1.text := 'Employee';
             CheckBox1.Enabled:=False;
             CheckBox1.Checked:=False;
           end;
      end; // of case
    end;
end;

procedure TForm1.tvOrgChartDragOver(Sender, Source: TObject; X, Y: Integer;
  State: TDragState; var Accept: Boolean);
begin
  Accept := True;
end;

end.

