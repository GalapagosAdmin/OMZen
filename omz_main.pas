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
    miImportSSL1001: TMenuItem;
    miExportHTML: TMenuItem;
    miExportTsv: TMenuItem;
    miImportSSL1000: TMenuItem;
    miExportDot: TMenuItem;
    miProcess: TMenuItem;
    OpenDialog1: TOpenDialog;
    PageControl1: TPageControl;
    SaveDialog1: TSaveDialog;
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
    procedure miImportSSL1001Click(Sender: TObject);
    procedure miExportHTMLClick(Sender: TObject);
    procedure miExportDotClick(Sender: TObject);
    procedure miExportTsvClick(Sender: TObject);
    procedure miImportSSL1000Click(Sender: TObject);
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

procedure TForm1.miImportSSL1001Click(Sender: TObject);
begin
  if OpenDialog1.Execute then
    begin
      Import_ADP_HRP1001(OpenDialog1.FileName);
      miProcess.enabled := True;
    end;
end;

procedure TForm1.miExportHTMLClick(Sender: TObject);
var
  Node: TTreeNode;
  Indent, s: Integer;
  Padding: utf8string;
  icon_file: utf8string;
  htmlFile:TextFile;
const
  LevelIndent = 4;
begin
  if SaveDialog1.Execute then
   begin
    AssignFile(htmlFile, SaveDialog1.FileName);
    Rewrite(htmlFile);
    Writeln(htmlFile, '<html><head><title>Org. Chart</title></head><body><pre>');

    Node := tvOrgChart.Items.GetFirstNode;
    while Node <> nil do
     begin
       if Node.Level = 0 then
        padding := ''
       else
        begin
//         Padding := StringOfChar(' ', Node.Level * LevelIndent-1);
//         s := 1;
         padding := '';
         for s := 1 to ((Node.Level-1) * LevelIndent) do //step 4 do
           begin
//         while s <= length(Padding) do begin
           if (s-1) mod 4 = 0              then
             padding := padding +  '│'
           else
             padding := padding + '&nbsp;'
//           Inc(s, 4);
         end;
         Padding := Padding + '└───';
        end;
      case Node.ImageIndex of
        IDX_ORG_UNIT:icon_file := 'ou_icon.png';
        IDX_POSITION:icon_file := 'normal_position_icon.png';
        IDX_CHIEF_POSITION:icon_file := 'chief_position_icon.png';
        IDX_EMPLOYEE:icon_file := 'employee_icon.png';
        else
          icon_file := '';
      end;
      Writeln(htmlFile, Padding + '<img src="' +
                                icon_file + '">&nbsp;'
                                + Node.Text + '');
//      + Node.Text + '<br>');
      Node := Node.GetNext;
     end;

    Writeln(htmlFile, '</pre></body></html>');
    System.close(htmlFile);
   end;
end;

procedure TForm1.miExportDotClick(Sender: TObject);
var
 dotFile:TextFile;
 i:LongInt;
 tmpStr:UTF8String;
begin
  if SaveDialog1.Execute then
   begin
    AssignFile(dotFile, SaveDialog1.FileName);
    Rewrite(dotFile);
//    Writeln(dotfile, 'digraph o { rankdir=BT ');
    Writeln(dotfile, 'digraph o { rankdir=RL; ratio=compress ');
    writeln(dotfile,'node [shape=box];');
    For i := low(ObjectList) to high(ObjectList) do
      if ObjectList[i].ObjType = OBJ_ORG_UNIT then
      begin
        case ObjectList[i].hasparent of

          False: begin
            tmpStr :=  IntToStr(ObjectList[i].objnum)
                          + ' [label = "' + ObjectList[i].LongText + '"]'
                          + ';';
            Writeln(dotFile, tmpStr);
            end;

          True:begin
            tmpStr :=  IntToStr(ObjectList[i].objnum)
            + ' [label = "' + DotStrip(ObjectList[i].LongText) + '"];';
//            + ' [label = "O ' + IntToStr(ObjectList[i].objnum) + '"];';
            Writeln(dotFile, tmpStr);

//                          + ' -> ' + IntToStr(ObjectList[i].ParentObjID) + ';';
//            Writeln(dotFile, tmpStr);
            Writeln(dotFile, IntToStr(ObjectList[i].objnum)+ ' -> ' + IntToStr(ObjectList[i].ParentObjID) + ';' );
               end;
          end; // of CASE
      end; // of FOR
    Writeln(dotfile, '}');
    System.close(dotfile);
   end;
end;

procedure TForm1.miExportTsvClick(Sender: TObject);
begin
  if OpenDialog1.Execute then
    tvOrgChart.SaveToFile(SaveDialog1.FileName);
end;

procedure TForm1.miImportSSL1000Click(Sender: TObject);
begin
    if OpenDialog1.Execute then
    begin
      Import_ADP_HRP1000(OpenDialog1.FileName);
      miImportSSL1001.enabled := true;
    end;
end;

procedure TForm1.miProcessClick(Sender: TObject);

Procedure Add_ou_no_parent(const h,i,j:longint);
 begin // Has no parent, this is a root OU
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

Procedure Add_ou_with_parent(const h,i,j:Longint);
  var
    ParentNode:TTreeNode;
       begin   // Has a parent, we should put this under an existing OU
              // Search for the parent
              ParentNode := tvOrgChart.Items.FindNodeWithData(Pointer(ObjectList[i].ParentObjID));
              if ParentNode = nil then
                begin  // We couldn't find the parent on the Org. Chart
                 SendDebug('Level:' + IntToStr(H) + ' Skipping OU with missing parent: ' + IntToStr(ObjectList[i].ObjNum) + ObjectList[i].LongText);
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
                 ParentNode.Expand(True);
                 ObjectList[i].Stale := False;
               end; // parent found
            end; // of TRUE

Procedure Add_Position(const h,i,j:LongInt);
 var
   ParentNode:TTreeNode;
   begin
    ParentNode := tvOrgChart.Items.FindNodeWithData(Pointer(ObjectList[i].ParentObjID));
    if ParentNode = nil then
      begin  // We couldn't find the parent on the Org. Chart
       SendDebug('Level:' + IntToStr(H) + ' Skipping Position with missing parent: ' + IntToStr(ObjectList[i].ObjNum) + ObjectList[i].LongText);
      end
    else // Parent found
     begin   // We found the parent
    with
       tvOrgChart.Items.AddChildObject(ParentNode,       // Parent ID
               ObjectList[i].LongText,                 // Long Text
               Pointer(ObjectList[i].ObjNum)) do       // Object Number
                 begin
                   case ObjectList[i].Chief of
                     false:ImageIndex := IDX_POSITION;
                     true:ImageIndex :=  IDX_CHIEF_POSITION;
                   end;
                   SelectedIndex := ImageIndex;
                 end;
       ParentNode.Expand(True);
       ObjectList[i].Stale := False;
     end; // parent found

   end; // Add Position

const
  MAX_DEPTH = 20;
var
  h, i,j:longint;
  ParentNode:TTreeNode;
begin
  tvOrgChart.items.clear;
  // This transfers parent ID relationship from IT1001 to IT1000 structure.
  UpdateHasParent;
 for h := 1 to MAX_DEPTH do // 10 rounds = max 10 levels org. chart heirarchy
   // Loop through all objects
   for i := low(ObjectList) to high(ObjectList) do
     if (ObjectList[i].Stale) then
     // We are only interested in Org. Units for now, and ones we haven't updated yet.
    case ObjectList[i].ObjType of  // Org. Unit
     'O': case ObjectList[i].HasParent of
            False: Add_ou_no_parent(h,i,j);
            True: Add_ou_with_parent(h,i,j);
          end;
     'S': Add_position(h,i,j);
     else
       SendDebug('Skipping Unhandled Object Type: ' + ObjectList[i].ObjType );
    end; // Case ObjType
  miExportHTML.enabled := True;
  miExportDot.enabled := True;
  miExportTsv.enabled := True;
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
