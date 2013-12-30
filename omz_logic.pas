unit omz_logic;

{$mode objfpc}{$H+}

{$modeswitch advanced_records}
interface

uses
  Classes, SysUtils, CsvDocument;

const
 IDX_ORG_UNIT = 0;
 IDX_POSITION = 1;
 IDX_CHIEF_POSITION = 2;
 IDX_EMPLOYEE = 3;
 OBJ_ORG_UNIT = 'O';
 OBJ_POSITION = 'S';
 OBJ_JOB = 'C';
 OBJ_COSTCTR = 'K';
 OBJ_EMPLOYEE = 'P';

type
  TRelationshipEntry=record
    // Standard SAP attributes
    SrcObjType:Char;
    SrcObjNum:LongInt;
    Relationship:String[4];
    BeginDate:String[10];
    EndDate:String[10];
    DestObjType:Char;
    DestObjNum:LongInt;
    // Extra attributes for internal use
   // Stale:Boolean;
  end;

  TObjectEntry=record
    ObjType:Char;
    ObjNum:LongInt;
    BeginDate:String[10];
    EndDate:String[10];
    ShortText:String[24];
    LongText:String[80];
    Dummy:String;
    LangCode:String[2];
    // Extra attributes for internal use
    HasParent:Boolean;
    ParentObjID:Longint;
    Stale:Boolean;
    Chief:Boolean;
    Job:LongInt;
    CostCtr:LongInt;
  end;

  //  TRelationshipList=specialize TFPGList<TRelationshipEntry>;
  TRelationshipList = Array of TRelationshipEntry;
  TObjectList = Array of TObjectEntry;


var
   // Lists
  RelationshipList:TRelationshipList;
  ObjectList:TObjectList;

Function GetNextOUObjNum:LongInt;
Function GetNextPosObjNum:LongInt;
Procedure Import_ADP_HRP1000(Const UTF16FileName:UTF8String);
// Imports the H1000 sheet of an OM SSL file, which has been exported from Excel
// by using Save As... > Unicode  (This produces a tab delimited UTF16LE text file)
Procedure Import_ADP_HRP1001(Const UTF16FileName:UTF8String);
// Imports the H1001 sheet of an OM SSL file, which has been exported from Excel
// by using Save As... > Unicode  (This produces a tab delimited UTF16LE text file)

Procedure UpdateHasParent;
Procedure UpdateCstCtr;
Procedure UpdateJob;
Function GetObjectById(Const ObjID:LongInt):TObjectEntry;

Function DotStrip(const Original:UTF8String):UTF8String;


implementation

uses FileUtil,lazutf8, fgl, CharEncStreams, dbugintf, strutils;

//class operator TRelationshipEntry.=(aLeft, aRight: TRelationshipEntry): Boolean;
//begin
//   result := aLeft.SrcObjID = aright.SrcObjID;
//end;


var
  HWMOUObjNum :LongInt;  // High Water Mark
  HWMPosObjNum :LongInt;

Function DotStrip(const Original:UTF8String):UTF8String;
  begin
    Result := Original;
    DelChars(Result, '/');
    DelChars(Result, '(');
    DelChars(Result, ')');
    DelChars(Result, '&');
  end;

Procedure Init_number_ranges;
// Set various high-water marks to number ranges
  begin
   HWMOUObjNum := 12200000; // Should be read from settings
   HWMPosObjNum := 22200000; // Should be read from settings
  end;

Procedure Init_Lists;
  begin
   SetLength(ObjectList,0);
   SetLength(RelationshipList,0);
  end;

Function GetNextOUObjNum:LongInt;
  begin
     Result := HWMOUObjNum + 1;
     HWMOUObjNum := Result;
  end;

Function GetNextPosObjNum:LongInt;
  begin
     Result := HWMPosObjNum + 1;
     HWMPosObjNum := Result;
  end;

// Flatten IT1001 into IT1000
Procedure UpdateHasParent;
  var
    i,j:LongInt;
  begin
   for i := low(ObjectList) to high(ObjectList) do
     begin
     // Look for Org. Unit objects with no parent
     If (not ObjectList[i].HasParent) and (ObjectList[i].ObjType = OBJ_ORG_UNIT) then
       begin
         // Search for the parent
         for j := low(RelationshipList) to high(RelationshipList) do
           // Obj to Obj relationship
           If (RelationshipList[j].SrcObjType = OBJ_ORG_UNIT)
             and (RelationshipList[j].DestObjType = OBJ_ORG_UNIT)
             and (RelationshipList[j].Relationship = 'A002')
             // Where the source is the current ObjectID
             and (RelationshipList[j].SrcObjNum = ObjectList[i].ObjNum )
             // If so, then we have a parent ID, so we set HasParent to True.
             then
               with ObjectList[i] do begin
                 HasParent := True;
                 ParentObjID := RelationshipList[j].DestObjNum;
               end;
         end; // if
     // Search for Positions w/o assignment to OU
     If (not ObjectList[i].HasParent)
         and (ObjectList[i].ObjType = OBJ_POSITION) then
       begin
         // Search for the parent
         for j := low(RelationshipList) to high(RelationshipList) do
           // Obj to Obj relationship
           If (RelationshipList[j].SrcObjType = OBJ_POSITION)
             and (RelationshipList[j].DestObjType = OBJ_ORG_UNIT)
             and ( (RelationshipList[j].Relationship = 'A003')   // Normal Reporting
                     or (RelationshipList[j].Relationship = 'A012') )  // Chief (Manager)
             // Where the source is the current ObjectID
             and (RelationshipList[j].SrcObjNum = ObjectList[i].ObjNum )
             // If so, then we have a parent ID, so we set HasParent to True.
             then
               with ObjectList[i] do begin
                 HasParent := True;
                 ParentObjID := RelationshipList[j].DestObjNum;
                 if RelationshipList[j].Relationship = 'A012' then
                   Chief := True;
               end;
       end;
    end; // for
  end; // of Procedure

Function GetObjectById(Const ObjID:LongInt):TObjectEntry;
  var
    i:longint;
  begin
   for i := low(ObjectList) to high(ObjectList) do
     if  ObjectList[i].ObjNum = ObjID then
     begin
       result := ObjectList[i];
       exit;
     end;
    result.costctr := 0;
  end;

Procedure UpdateCstCtr;
  var
    i,j:LongInt;
  begin
   for i := low(ObjectList) to high(ObjectList) do
     begin
      // put filter criteria here
       begin
         // Search for the parent
         for j := low(RelationshipList) to high(RelationshipList) do
           // Obj to Obj relationship
           If (RelationshipList[j].SrcObjType <> OBJ_COSTCTR)
            and (RelationshipList[j].DestObjType = OBJ_COSTCTR)
//             and (RelationshipList[j].Relationship = 'A002')
             // Where the source is the current ObjectID
             and (RelationshipList[j].SrcObjNum = ObjectList[i].ObjNum )
             // If so, then we have a parent ID, so we set HasParent to True.
             then
               with ObjectList[i] do begin
         //        HasParent := True;
                 CostCtr := RelationshipList[j].DestObjNum;
               end;
         end; // if

    end; // for
  end; // of Procedure


Procedure UpdateJob;
  var
    i,j:LongInt;
  begin
   for i := low(ObjectList) to high(ObjectList) do
     begin
      // put filter criteria here
       If (ObjectList[i].ObjType = OBJ_POSITION) then
       begin
         // Search for the parent
         for j := low(RelationshipList) to high(RelationshipList) do
           // Obj to Obj relationship
           If (RelationshipList[j].SrcObjType = OBJ_POSITION)
            and (RelationshipList[j].DestObjType = OBJ_JOB)
//             and (RelationshipList[j].Relationship = 'A002')
             // Where the source is the current ObjectID
             and (RelationshipList[j].SrcObjNum = ObjectList[i].ObjNum )
             // If so, then we have a parent ID, so we set HasParent to True.
             then
               with ObjectList[i] do begin
         //        HasParent := True;
                 Job := RelationshipList[j].DestObjNum;
               end;
         end; // if

    end; // for
  end; // of Procedure

// Convert Excel UTF16LE TSV .txt file to UTF8 for easier processing
Function UTF16FiletoUTF8File(const InFileName:UTF8String):UTF8String;
  var
    OutFileName:UTF8String;
    Infile, Outfile: TextFile;
    buffer:UnicodeString;
    OutBuffer:UTF8String;
    fCES:TCharEncStream;
    StringList:TStringList;
  begin
   if not(FileExistsUTF8(InFileName)) then exit;
   OutFileName := ExtractFileNameWithoutExt(InFileName) + '-UTF8.txt';
   SendDebug('Converting ' + InFileName + ' to '+ OutFileName);
(*   AssignFile(Infile, UTF8toANSI(InFileName));
   system.Reset(Infile);
   system.assign(outfile, UTF8ToANSI(OutFileName));
   system.rewrite(outfile);
   while not eof(infile) do
     begin
       Readln(Infile, buffer);
       OutBuffer := UTF16toUTF8(buffer);
       Writeln(outfile, OutBuffer);
     end;
   system.close(Infile);
   system.close(outfile); *)
   fCES := TCharEncStream.Create;
   fCES.LoadFromFile(UTF8ToANSI(InFileName));
   StringList := TStringList.Create;
   StringList.text := fCES.UTF8Text;
   StringList.SaveToFile(UTF8ToANSI(OutFileName));
   fCES.free;
   Result := OutFileName;
  end;

Procedure Import_ADP_HRP1000(Const UTF16FileName:UTF8String);
  var
    UTF8FileName:UTF8String;
    FileStream: TFileStream;
    Parser: TCSVParser;
    ADelimiter:Char=#9; // Tab?
    LineBuffer:TObjectEntry;
    tmpString :String;
    cRow, cCol:LongInt;
    RecordsLoaded:LongInt;
  begin
   SetLength(ObjectList,0);
   RecordsLoaded := 0;
   UTF8FileName := UTF16FiletoUTF8File(UTF16FileName);
   if not(FileExistsUTF8(UTF8FileName)) then exit;
   Parser:=TCSVParser.Create;
   FileStream := TFileStream.Create(UTF8ToANSI(UTF8Filename), fmOpenRead+fmShareDenyWrite);
   try
    Parser.Delimiter:=ADelimiter;
    Parser.SetSource(FileStream);
     while Parser.ParseNextCell do
       begin
         cRow := Parser.CurrentRow;
         cCol := Parser.CurrentCol;
        Case Parser.CurrentRow of
         0,1,2:; // Skip header rows
         else // Data starts from row 3 (zero based)
           begin
            tmpString := Parser.CurrentCellText;

             case Parser.CurrentCol of
               0:begin // Object Type
                   If length(Parser.CurrentCellText) > 0 then
                    LineBuffer.ObjType:=Copy(Parser.CurrentCellText,1,1)[1];
                 end;
               1:begin // Object ID
                   LineBuffer.ObjNum:=StrToInt(Parser.CurrentCellText);
                   // Update high water mark if necessary
                   case LineBuffer.ObjType of
                     OBJ_ORG_UNIT: if LineBuffer.ObjNum > HWMOUObjNum then
                                       HWMOUObjNum := LineBuffer.ObjNum;
                     OBJ_POSITION: if LineBuffer.ObjNum > HWMPosObjNum then
                                       HWMPosObjNum := LineBuffer.ObjNum
                   end;  // of CASE
                 end;
               2:begin // Begin Date
                   LineBuffer.BeginDate:=Parser.CurrentCellText;
                 end;
               3:begin // End Date
                   LineBuffer.EndDate:=Parser.CurrentCellText;
                 end;
               4:begin // Short Text Description
                    LineBuffer.ShortText := UTF8Copy(Parser.CurrentCellText,1,12);
                 end;
               5:begin // Long Text Description
                   LineBuffer.LongText := UTF8Copy(Parser.CurrentCellText,1,40);
               end;
               6:Begin        // deleimit date?
                 end;
               7:Begin
                 LineBuffer.LangCode := UTF8Copy(Parser.CurrentCellText,1,2);
                 LineBuffer.HasParent := False;
                 LineBuffer.Stale := True;
                 LineBuffer.Chief := False;
                 LineBuffer.ParentObjID := 0;
                 Inc(RecordsLoaded);
                 SetLength(ObjectList,RecordsLoaded);
                 SendDebug(LineBuffer.ObjType
                           + IntToStr(LineBuffer.ObjNum)
                           + LineBuffer.BeginDate
                           + LineBuffer.EndDate
                           + LineBuffer.ShortText
                           + LineBuffer.LongText
                           + LineBuffer.LangCode
                           );
                 ObjectList[RecordsLoaded-1] := LineBuffer;

                 end;
             end; // of CASE CurrentCol
           end;// row >= 3
        end; // of CASE CurrentRow
      // Set Array text
      //:=Parser.CurrentCellText;
      end;
     SendDebug('Records Loaded:' + IntToStr(RecordsLoaded) );


   finally
     Parser.Free;
     FileStream.Free;
     If not DeleteFileUTF8(UTF8FileName) then
       SendDebug('Warning: Unable to delete temporary UTF8 file.');
   end;

  end;



Procedure Import_ADP_HRP1001(Const UTF16FileName:UTF8String);
  var
    UTF8FileName:UTF8String;
    FileStream: TFileStream;
    Parser: TCSVParser;
    ADelimiter:Char=#9; // Tab?
    LineBuffer:TRelationshipEntry;
    tmpString :String;
    cRow, cCol:LongInt;
    RecordsLoaded:LongInt;
  begin
   // Clear out high water marks
   Init_number_ranges;
   // Erase current items (we don't append)
   SetLength(RelationshipList,0);
   RecordsLoaded := 0;
   UTF8FileName := UTF16FiletoUTF8File(UTF16FileName);
   if not(FileExistsUTF8(UTF8FileName)) then exit;
   Parser:=TCSVParser.Create;
   FileStream := TFileStream.Create(UTF8ToANSI(UTF8Filename), fmOpenRead+fmShareDenyWrite);
   try
    Parser.Delimiter:=ADelimiter;
    Parser.SetSource(FileStream);
     while Parser.ParseNextCell do
       begin
         cRow := Parser.CurrentRow;
         cCol := Parser.CurrentCol;
        Case Parser.CurrentRow of
         0,1,2:; // Skip header rows
         else // Data starts from row 3 (zero based)
           begin
            tmpString := Parser.CurrentCellText;

             case Parser.CurrentCol of
               0:begin // Source Object Type
                   If length(Parser.CurrentCellText) > 0 then
                    LineBuffer.SrcObjType:=Copy(Parser.CurrentCellText,1,1)[1];
                 end;
               1:begin // Source Object ID

                   LineBuffer.SrcObjNum:=StrToInt(Parser.CurrentCellText);
                 end;
               2:begin // Relationship Direction + code
                   LineBuffer.Relationship:=Parser.CurrentCellText;
                 end;
               3:begin // Begin Date
                   LineBuffer.BeginDate:=Parser.CurrentCellText;
                 end;
               4:begin // End Date
                   LineBuffer.EndDate:=Parser.CurrentCellText;
                 end;
               5:begin // Destination Object Type
                   If length(Parser.CurrentCellText) > 0 then
                    LineBuffer.DestObjType := Parser.CurrentCellText[1];
                 end;
               6:begin // Destination Object Number
                   LineBuffer.DestObjNum := StrToInt(Parser.CurrentCellText);
                   Inc(RecordsLoaded);
                   SetLength(RelationshipList,RecordsLoaded);
                   SendDebug(LineBuffer.SrcObjType
                             + IntToStr(LineBuffer.SrcObjNum)
                             + LineBuffer.Relationship
                             + LineBuffer.BeginDate
                             + LineBuffer.EndDate
                             + LineBuffer.DestObjType
                             + IntToStr(LineBuffer.DestObjNum)
                             );
                   RelationshipList[RecordsLoaded-1] := LineBuffer;
               end;
             end; // of CASE CurrentCol
           end;// row >= 3
        end; // of CASE CurrentRow
      // Set Array text
      //:=Parser.CurrentCellText;
      end;
     SendDebug('Records Loaded:' + IntToStr(RecordsLoaded) );


   finally
     Parser.Free;
     FileStream.Free;
     If not DeleteFileUTF8(UTF8FileName) then
       SendDebug('Warning: Unable to delete temporary UTF8 file.');
   end;

  end;



initialization
  Init_Number_Ranges;
  Init_Lists;

end.

