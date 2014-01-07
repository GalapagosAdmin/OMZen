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
    leObjTextL: TEdit;
    leKostlTextL: TEdit;
    leJobTextL: TEdit;
    leObjText: TEdit;
    leKostlText: TEdit;
    leJobText: TEdit;
    ImageList1: TImageList;
    LabeledEdit1: TLabeledEdit;
    LabeledEdit2: TLabeledEdit;
    leCstCtr: TLabeledEdit;
    leJobCd: TLabeledEdit;
    MainMenu1: TMainMenu;
    MenuItem1: TMenuItem;
    MenuItem2: TMenuItem;
    miEEFileExport: TMenuItem;
    miImportSSLCCRMD: TMenuItem;
    miImportSSLNHIRE: TMenuItem;
    miAbout: TMenuItem;
    miImportSSL1001: TMenuItem;
    miExportHTML: TMenuItem;
    miExportTsv: TMenuItem;
    miImportSSL1000: TMenuItem;
    miExportDot: TMenuItem;
    miProcess: TMenuItem;
    OpenDialog1: TOpenDialog;
    PageControl1: TPageControl;
    Panel1: TPanel;
    SaveDialog1: TSaveDialog;
    StatusBar1: TStatusBar;
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
    procedure miEEFileExportClick(Sender: TObject);
    procedure miImportSSLCCRMDClick(Sender: TObject);
    procedure miAboutClick(Sender: TObject);
    procedure miImportSSL1001Click(Sender: TObject);
    procedure miExportHTMLClick(Sender: TObject);
    procedure miExportDotClick(Sender: TObject);
    procedure miExportTsvClick(Sender: TObject);
    procedure miImportSSL1000Click(Sender: TObject);
    procedure miImportSSLNHIREClick(Sender: TObject);
    procedure miProcessClick(Sender: TObject);
    procedure tvOrgChartClick(Sender: TObject);
    procedure tvOrgChartDragDrop(Sender, Source: TObject; X, Y: Integer);
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
  omz_logic, dbugintf, omz_rsrc;

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

end;

procedure TForm1.miEEFileExportClick(Sender: TObject);
Const
  sep:char=#9;
var
  EEfile:TextFile;
  i:LongInt;
  outline:UTF8String;
  EEObj:TObjectEntry;
  PosObj:TObjectEntry;
  OU1:TObjectEntry;
  OU2:TObjectEntry;
  OU3:TObjectEntry;
  kostl:LongInt;
  ChiefFlag:UTF8String;
begin
if SaveDialog1.Execute then
 begin
  AssignFile(EEFile, UTF8ToANSI(SaveDialog1.FileName));
  Rewrite(EEFile);
  Outline := 'EE No.'
             + sep + 'Kanji Name'
             + sep + 'Romaji Name'
             + sep + 'Position No.'
             + Sep + 'Pos. Text(S)'
             + Sep + 'Pos. Text(L)'
             + Sep + 'Job No.'
             + sep + 'Job Text'
             + sep + 'Manager'
             + sep + 'OU3 No.'
             + sep + 'OU3 Text'
             + sep + 'OU2 No.'
             + sep + 'OU2 Text'
             + sep + 'OU1 No.'
             + sep + 'OU1 Text'
             + sep + 'Eff. CC No.'
             + sep + 'Eff. CC Text'

             ;
  Writeln(EEFile, outline);

  for i := low(ObjectList) to high(ObjectList) do
    if ObjectList[i].ObjType = OBJ_EMPLOYEE then
    begin
     FillChar(eeObj, SizeOf(eeObj), #0);;
     FillChar(PosObj, SizeOf(PosObj), #0);;
     FillChar(OU1, SizeOf(ou1), #0);;
     FillChar(OU2, SizeOf(ou2), #0);;
     FillChar(OU3, SizeOf(ou3), #0);;
     eeObj := ObjectList[i];
     kostl := 0;
     PosObj := GetObjectById(eeObj.ParentObjID);
     OU1 := GetObjectById(PosObj.ParentObjID);
     OU2 := GetObjectById(OU1.ParentObjID);
     OU3 := GetObjectById(OU2.ParentObjID);
     // calculate effective cost center
     if ou3.CostCtr <> 0 then
      Kostl := ou3.CostCtr;
     if ou2.CostCtr <> 0 then
      Kostl := ou2.CostCtr;
     if ou1.CostCtr <> 0 then
      Kostl := ou1.CostCtr;
     if PosObj.CostCtr <> 0 then
      Kostl := PosObj.CostCtr;

     If PosObj.Chief then ChiefFlag := 'Yes' else ChiefFlag := 'No';

     Outline := IntToStr(eeObj.ObjNum)
                 + sep
                 + eeObj.ShortText
                 + sep
                 + eeObj.LongText
                 + sep
                 + IntToStr(eeObj.ParentObjID)
                 + Sep + PosObj.LongText
                 + Sep + IntToStr(PosObj.Job)
                 + sep + GetObjectShortTextByID(PosObj.Job)
                 + sep + GetObjectLongTextByID(PosObj.Job)
                 + sep + ChiefFlag
                 + sep + IntToStr(ou3.ObjNum)
                 + sep + ou3.LongText
                 + sep + IntToStr(ou2.ObjNum)
                 + sep + ou2.LongText
                 + sep + IntToStr(ou1.ObjNum)
                 + sep + ou1.LongText
                 + sep + IntToStr(kostl)
                 + sep + GetObjectLongTextByID(kostl)
                 ;
      Writeln(EEFile, outline);
    end;
  System.Close(EEFile);
  end;
end;
procedure TForm1.miImportSSLCCRMDClick(Sender: TObject);
begin
   if OpenDialog1.Execute then
     begin
       Import_ADP_CCRMD(OpenDialog1.FileName);
       miProcess.enabled := True;
       miImportSSLCCRMD.Checked := True;
       // Don't allow the user to load it more than once.
       miImportSSLCCRMD.Enabled := False;
//       miImportSSLNHIRE.Enabled:=True;
       StatusBar1.SimpleText := rsImportSuccCCRMD;
     end;
end;

procedure TForm1.miAboutClick(Sender: TObject);
begin
  showmessage(rsAbout);
end;

procedure TForm1.miImportSSL1001Click(Sender: TObject);
begin
  if OpenDialog1.Execute then
    begin
      Import_ADP_HRP1001(OpenDialog1.FileName);
      miProcess.enabled := True;
      miImportSSL1001.Checked := True;
      miImportSSLNHIRE.Enabled:=True;
      StatusBar1.SimpleText := rsImportSucc1001;
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
    AssignFile(htmlFile, UTF8ToANSI(SaveDialog1.FileName));
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
    StatusBar1.SimpleText := rsHTMLExportComplete;
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
    AssignFile(dotFile, UTF8ToANSI(SaveDialog1.FileName));
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
    StatusBar1.SimpleText:=rsDotExportComplete;
   end;
end;

procedure TForm1.miExportTsvClick(Sender: TObject);
begin
  if OpenDialog1.Execute then
    begin
      tvOrgChart.SaveToFile(SaveDialog1.FileName);
      StatusBar1.SimpleText:= rsTSVExportComplete;
    end;
end;

procedure TForm1.miImportSSL1000Click(Sender: TObject);
begin
    if OpenDialog1.Execute then
    begin
      Import_ADP_HRP1000(OpenDialog1.FileName);
      miImportSSL1001.Enabled := True;
      miImportSSLCCRMD.Enabled := True;
      miImportSSL1000.Checked := True;
      miImportSSLCCRMD.Checked := False;
      miImportSSLNhire.Checked:=False;
      StatusBar1.SimpleText := rsImportSucc1000;
    end;
end;

procedure TForm1.miImportSSLNHIREClick(Sender: TObject);
begin
    if OpenDialog1.Execute then
    begin
      Import_ADP_NHIRE(OpenDialog1.FileName);
      miImportSSLNhire.Checked:=True;
      // We can't allow them to load it more than once
      miImportSSLNhire.Enabled:=False;
      miEEFileExport.enabled := True;
      StatusBar1.SimpleText := rsImportSuccNHIRE;
    end;
end;

procedure TForm1.miProcessClick(Sender: TObject);

Procedure Add_ou_no_parent(const h,i,j:longint);
  var
   LongText:UTF8String;
 begin // Has no parent, this is a root OU
  LongText := IntToStr(ObjectList[i].ObjNum) + ' ' + ObjectList[i].LongText;
  with
       tvOrgChart.Items.AddChildObject(nil,               // Parent ID
                  LongText,                 // Long Text
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
    LongText:UTF8String;
       begin   // Has a parent, we should put this under an existing OU
              // Search for the parent
              ParentNode := tvOrgChart.Items.FindNodeWithData(Pointer(ObjectList[i].ParentObjID));
              if ParentNode = nil then
                begin  // We couldn't find the parent on the Org. Chart
                 SendDebug('Level:' + IntToStr(H) + ' Skipping OU with missing parent: ' + IntToStr(ObjectList[i].ObjNum) + ObjectList[i].LongText);
                end
              else // Parent found
               begin   // We found the parent
               // Set up item description on the tree
               LongText := IntToStr(ObjectList[i].ObjNum) + ' '
                           +  ObjectList[i].LongText;
               // Add Cost Center Text to description if we have it
               if ObjectList[i].CostCtr <> 0 then
                 LongText := LongText + ' [CC:' + IntToStr(ObjectList[i].CostCtr)
                             + ' ' + GetObjectLongTextByID(ObjectList[i].CostCtr) + ']';
               with
                 tvOrgChart.Items.AddChildObject(ParentNode,       // Parent ID
                         LongText,                                 // Long Text
                         Pointer(ObjectList[i].ObjNum)) do         // Object Number
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
    LongText:UTF8String;
   begin
    ParentNode := tvOrgChart.Items.FindNodeWithData(Pointer(ObjectList[i].ParentObjID));
    if ParentNode = nil then
      begin  // We couldn't find the parent on the Org. Chart
       SendDebug('Level:' + IntToStr(H) + ' Skipping Position with missing parent: ' + IntToStr(ObjectList[i].ObjNum) + ObjectList[i].LongText);
      end
    else // Parent found
     begin   // We found the parent
     // Generate Long Text
     LongText :=  IntToStr(ObjectList[i].ObjNum)
                   + ' ' + ObjectList[i].LongText;
     if ObjectList[i].CostCtr <> 0 then
       LongText := LongText + ' [CC:' + IntToStr(ObjectList[i].CostCtr)
           + ' ' + GetObjectLongTextByID(ObjectList[i].CostCtr) + ']';
     if ObjectList[i].Job <> 0 then
       LongText := LongText + ' [Job:' + IntToStr(ObjectList[i].Job)
           + ' ' + GetObjectLongTextByID(ObjectList[i].Job) + ']';
      with
       tvOrgChart.Items.AddChildObject(ParentNode,       // Parent ID
               LongText,                   // Long Text
               Pointer(ObjectList[i].ObjNum)) do         // Object Number
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

Procedure Add_Employee(const h,i,j:LongInt);
 var
   ParentNode:TTreeNode;
   LongText:UTF8String;
 begin
  ParentNode := tvOrgChart.Items.FindNodeWithData(Pointer(ObjectList[i].ParentObjID));
  if ParentNode = nil then
    begin  // We couldn't find the parent on the Org. Chart
     SendDebug('Level:' + IntToStr(H) + ' Skipping Employee with missing parent: ' + IntToStr(ObjectList[i].ObjNum) + ObjectList[i].LongText);
    end
  else // Parent (Position) found
   begin   // We found the parent (Position)
   // Generate Long Text
   LongText :=  IntToStr(ObjectList[i].ObjNum)
     + ' ' + ObjectList[i].LongText
     + ' (' + ObjectList[i].ShortText + ')';
  with
     tvOrgChart.Items.AddChildObject(ParentNode,       // Parent ID
             LongText,                                 // Long Text
             Pointer(ObjectList[i].ObjNum)) do         // Object Number
               begin
                 ImageIndex := IDX_EMPLOYEE;
                 SelectedIndex := ImageIndex;
               end;
     ParentNode.Expand(True);
     ObjectList[i].Stale := False;
   end; // parent found
 end; // Add Employee

const
  MAX_DEPTH = 20;
var
  h, i,j:longint;
  ParentNode:TTreeNode;
  ChangedThisRound:Boolean;
begin
  tvOrgChart.items.clear;
  // This transfers parent ID relationship from IT1001 to IT1000 structure.
  UpdateHasParent;
  UpdateCstCtr;
  UpdateJob;
  UpdateEE;
 for h := 1 to MAX_DEPTH do // 10 rounds = max 10 levels org. chart heirarchy
   begin
    ChangedThisRound := False;
   // Loop through all objects
   for i := low(ObjectList) to high(ObjectList) do
     begin
     if (ObjectList[i].Stale) then
       begin
        ChangedThisRound := True; // We found an outstanding item
     // We are only interested in Org. Units for now, and ones we haven't updated yet.
        case ObjectList[i].ObjType of  // Org. Unit
         OBJ_ORG_UNIT: case ObjectList[i].HasParent of
                         False: Add_ou_no_parent(h,i,j);
                         True: Add_ou_with_parent(h,i,j);
                       end;
         OBJ_POSITION: Add_Position(h,i,j);
         OBJ_EMPLOYEE: Add_Employee(h,i,j);
         else
           SendDebug('Skipping Unhandled Object Type: ' + ObjectList[i].ObjType );
        end; // Case ObjType
       end; // Stale Check
     end; // Object list Loop
       // Leave loop early if we havn't found any remaining stale items
//       If not ChangedThisRound then break;
     end; // FOR... Max_Depth
  miExportHTML.enabled := True;
  miExportDot.enabled := True;
  miExportTsv.enabled := True;
  miProcess.Enabled:=False;
  StatusBar1.SimpleText := rsProcessingComplete;
end;

procedure TForm1.tvOrgChartClick(Sender: TObject);
var
  ObjID:LongInt;
  ObjectEntry:TObjectEntry;
begin
   if tvOrgChart.Selected <> nil then
    begin
      ObjID:= Integer(tvOrgChart.Selected.Data);
      labelededit1.text := IntToStr(tvOrgChart.Selected.ImageIndex);
      labelededit2.text := IntToStr(ObjID);
      ObjectEntry := GetObjectById(ObjID);
      leObjText.text := ObjectEntry.ShortText;
      leObjTextL.text := ObjectEntry.LongText;
      leCstCtr.text := '';
      leCstCtr.text := IntToStr(ObjectEntry.CostCtr);
      leKostlText.text := GetObjectShortTextByID(ObjectEntry.CostCtr);
      leKostlTextL.text := GetObjectLongTextByID(ObjectEntry.CostCtr);
      leJobCd.text := '';
      Case tvOrgChart.Selected.ImageIndex of
        IDX_ORG_UNIT: begin  // Org. Unit
            bbAddChildOU.Enabled:=True;
            bbAddPosition.Enabled := True;
            bbAddEmployee1.Enabled:= False;
            labelededit1.text := rsOrgUnit;
            CheckBox1.Checked:=False;
            CheckBox1.Enabled:=False;
            leJobCd.text := 'N/A';
            leJobText.text := '';
            leJobTextL.text := '';
           end;
        IDX_POSITION: begin // Normal Position
             bbAddChildOU.Enabled:=False;
             bbAddPosition.Enabled := False;
             bbAddEmployee1.Enabled:= True;
             labelededit1.text := rsPosition;
             CheckBox1.Enabled:=True;
             CheckBox1.Checked:=False;
             leJobCd.text := IntToStr(ObjectEntry.Job);
             leJobText.text := GetObjectShortTextByID(ObjectEntry.Job);
             leJobTextL.text := GetObjectLongTextByID(ObjectEntry.Job);
           end;
        IDX_CHIEF_POSITION: begin // Chief Position
             bbAddChildOU.Enabled:= False;
             bbAddPosition.Enabled := False;
             bbAddEmployee1.Enabled:= True;
             labelededit1.text := rsChiefPosition;
             CheckBox1.Enabled:=True;
             CheckBox1.Checked:=True;
             leJobCd.text := IntToStr(ObjectEntry.Job);
             leJobText.text := GetObjectShortTextByID(ObjectEntry.Job);
             leJobTextL.text := GetObjectLongTextByID(ObjectEntry.Job);
           end;
        IDX_EMPLOYEE: begin // Employee
             bbAddChildOU.Enabled:= False;
             bbAddPosition.Enabled := False;
             bbAddEmployee1.Enabled:= False;
             labelededit1.text := rsEmployee;
             CheckBox1.Enabled:=False;
             CheckBox1.Checked:=False;
             leJobCd.text := 'N/A';
             leJobText.text := '';
             leJobTextL.text := '';
           end;
      end; // of case
    end;
end;

procedure TForm1.tvOrgChartDragDrop(Sender, Source: TObject; X, Y: Integer);
var
  tv     : TTreeView;
  iNode  : TTreeNode;
begin
  tv := TTreeView(Sender);      { Sender is TreeView where the data is being dropped  }
  iNode := tv.GetNodeAt(x,y);   { x,y are drop coordinates (relative to the Sender)   }
                                {   sinse Sender is TreeView we can evaluate          }
                                {   a tree at the X,Y coordinates                     }

  { TreeView can also be a Source! So we must make sure }
  { that Source is TEdit, before getting its text }
  if Source = Sender then begin         { drop is happening within a TreeView   }
    if Assigned(tv.Selected) and             {  check if any node has been selected  }
      (iNode <> tv.Selected) then            {   and we're droping to another node   }
    begin
      if iNode <> nil then
        tv.Selected.MoveTo(iNode, naAddChild) { complete the drop operation, by moving the selectede node }
      else
        tv.Selected.MoveTo(iNode, naAdd); { complete the drop operation, by moving in root of a TreeView }
    end;
  end;
end;

procedure TForm1.tvOrgChartDragOver(Sender, Source: TObject; X, Y: Integer;
  State: TDragState; var Accept: Boolean);
begin
  Accept := True;
end;

end.

