unit omz_logic;
// Reads input files, manages array
{$mode objfpc}{$H+}

{$modeswitch advanced_records}
interface

uses
  Classes, SysUtils, CsvDocument, INIFiles;

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
// Column number constants to convert from spreadsheet view to CSVfile column number.
 COL_A = 0;
 COL_B = 1;
 COL_C = 2;
 COL_D = 3;
 COL_E = 4;
 COL_F = 5;
 COL_G = 6;
 COL_H = 7;
 COL_I = 8;
 COL_J = 9;
 COL_K = 10;
 COL_L = 11;
 COL_M = 12;
 COL_N = 13;
 COL_O = 14;
 COL_P = 15;
 COL_Q = 16;
 COL_R = 17;
 COL_S = 18;
 COL_T = 19;
 COL_U = 20;
 COL_V = 21;
 COL_W = 22;
 COL_X = 23;
 COL_Y = 24;
 COL_Z = 25;
 COL_AA = 26;
 COL_AB = 27;
 COL_AW = 48;



type
  // We need 64 bit to cover some long cost centers.
  TObjID=int64;

  TRelationshipEntry=record
    // Standard SAP attributes
    SrcObjType:Char;
    SrcObjNum:TObjID;
    Relationship:String[4];
    BeginDate:String[10];
    EndDate:String[10];
    DestObjType:Char;
    DestObjNum:TObjID;
    // Extra attributes for internal use
   // Stale:Boolean;
    Used:Boolean;
    SrcObjIdx:LongInt;  // Index into Object Array
    DestObjIdx:LongInt; // Index into Object Array
  end;

  TObjectEntry=record
    ObjType:Char;
    ObjNum:TObjID;
    BeginDate:String[10];
    EndDate:String[10];
    ShortText:String[24];
    LongText:String[80];
    Dummy:String;
    LangCode:String[2];
    // Extra attributes for internal use
    HasParent:Boolean;
    ParentObjID:TObjID;
    Stale:Boolean;
    Chief:Boolean;
    Job:TObjID;  // Job Code for Position
    CostCtr:TObjID; // Cost Center for Position or Org. Unit
    Used:Boolean;
  end;

  //  TRelationshipList=specialize TFPGList<TRelationshipEntry>;
  TRelationshipList = Array of TRelationshipEntry;
  TObjectList = Array of TObjectEntry;


var
   // Lists
  RelationshipList:TRelationshipList;
  ObjectList:TObjectList;
  INI:TINIFile;
   HWMOUObjMin,
 HWMOUObjMin2,
 HWMPosObjMin,
 HWMEEObjMin,
 HWMCCObjMin,
 HWMOUObjMax,
 HWMOUObjMax2,
 HWMPosObjMax,
 HWMEEObjMax,
 HWMCCObjMax:LongInt;

Function EmptyObj:TObjectEntry;
Function GetNextOUObjNum:LongInt;
Function GetNextPosObjNum:LongInt;
Procedure Import_ADP_HRP1000(Const UTF16FileName:UTF8String);
// Imports the H1000 sheet of an OM SSL file, which has been exported from Excel
// by using Save As... > Unicode  (This produces a tab delimited UTF16LE text file)
Procedure Import_ADP_HRP1001(Const UTF16FileName:UTF8String);
// Imports the H1001 sheet of an OM SSL file, which has been exported from Excel
// by using Save As... > Unicode  (This produces a tab delimited UTF16LE text file)

Procedure Import_ADP_NHIRE(Const UTF16FileName:UTF8String);
// Imports NHIRE (Hire Action) Sheet, which has been exported from Excel
// by using Save As... > Unicode  (This produces a tab delimited UTF16LE text file)

Procedure Import_ADP_CCRMD(Const UTF16FileName:UTF8String);
// Imports CCRMD (Cost Center Master Data) Sheet, which has been exported from Excel
// by using Save As... > Unicode  (This produces a tab delimited UTF16LE text file)


Procedure UpdateHasParent;
Procedure UpdateCstCtr;
Procedure UpdateJob;
Procedure UpdateEE; // Update after loading NHIRE

Function GetObjectById(Const ObjID:TObjID):TObjectEntry;
Function GetObjectIdxById(Const ObjID:TObjID):LongInt;
Function GetObjectLongTextByID(Const ObjID:TObjID):UTF8String;
Function GetObjectShortTextByID(Const ObjID:TObjID):UTF8String;

Function DotStrip(const Original:UTF8String):UTF8String;
Procedure UnusedObjectReport(StringList:TStrings);
Procedure NumberRangeReport(StringList:TStrings);

Function ObjIDToStr(const OBJID:TObjID):UTF8String;


// This builds an index to for the objext list array for faster access when
// using relationships.
// This should be called after all files are loaded, before any other processing.
Procedure IndexRelations;

implementation

uses FileUtil,lazutf8, fgl, CharEncStreams, dbugintf, strutils;

//class operator TRelationshipEntry.=(aLeft, aRight: TRelationshipEntry): Boolean;
//begin
//   result := aLeft.SrcObjID = aright.SrcObjID;
//end;


var
  HWMOUObjNum  :LongInt;  // High Water Mark
  HWMPosObjNum :LongInt;
  HWMEEObjNum  :LongInt;
  HWMCCObjNum  :LongInt;

Function ObjIDToStr(const ObjID:TObjID):UTF8String;
  begin
    if ObjID = 0 then
      result := '' // This just means it wasn't set, so use blank
    else
      begin
        Result := IntToStr(ObjID);
      end;
  end;

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
    HWMOUObjMin := INI.ReadInteger('NumberRange', 'HWMOUObjMin', 12200000);
    HWMOUObjMin2 :=  HWMOUObjMin;
    HWMPosObjMin := INI.ReadInteger('NumberRange', 'HWMPosObjMin', 22200000);
    HWMEEObjMin :=  INI.ReadInteger('NumberRange', 'HWMEEObjMin', 22600000);
    HWMCCObjMin :=  INI.ReadInteger('NumberRange', 'HWMCCObjMin', 9990000000);
    HWMOUObjMax := INI.ReadInteger('NumberRange', 'HWMOUObjMax', 12299999);
    HWMOUObjMax2 :=  HWMOUObjMax;
    HWMPosObjMax := INI.ReadInteger('NumberRange', 'HWMPosObjMax', 22299999);
    HWMEEObjMax := INI.ReadInteger('NumberRange', 'HWMEEObjMax', 22699999);
    HWMCCObjMax := INI.ReadInteger('NumberRange', 'HWMCCObjMax', 9999999999);

    HWMOUObjNum :=  HWMOUObjMin;
    HWMPosObjNum := HWMPosObjMin;
    HWMEEObjNum :=  HWMEEObjMin;
    HWMCCObjNum :=  HWMCCObjMin;
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

Function GetNextEEObjNum:LongInt;
  begin
     Result := HWMEEObjNum + 1;
     HWMEEObjNum := Result;
  end;

Procedure UpdateEE;
    var
      i,j:LongInt;
  begin
     // ParentObjID is loaded when the NHIRE file is loaded.
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
                 Used := True;  // what is the parent object doesn't exist?
                 ParentObjID := RelationshipList[j].DestObjNum;
                 ObjectList[RelationshipList[j].DestObjIdx].Used := True;
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
                 Used := True;  // what is the parent object doesn't exist?
                 ObjectList[RelationshipList[j].DestObjIdx].Used := True;
                 RelationshipList[j].Used := True;
                 HasParent := True;
                 ParentObjID := RelationshipList[j].DestObjNum;
                 if RelationshipList[j].Relationship = 'A012' then
                   Chief := True;
               end;
       end;
    end; // for
  end; // of Procedure

Procedure IndexRelations;
  var
    i,j:LongInt;
  begin
    for j := low(RelationshipList) to high(RelationshipList) do
        with RelationshipList[j] do
          begin
            // Look up index in Object array for source object ID
            SrcObjIdx := GetObjectIdxById(SrcObjNum);
            // Look up index in Object array for destination object ID
            DestObjIdx := GetObjectIdxById(DestObjNum);
          end;
  end;

Procedure UnusedObjectReport(StringList:TStrings);
  var
    i:LongInt;
  begin
     StringList.Append('Unused Objects');
     for i := low(ObjectList) to high(ObjectList) do
       case ObjectList[i].ObjType of
       // Look for Unused objects
       OBJ_EMPLOYEE,   // perhaps skip employees with Position 99999999
       OBJ_ORG_UNIT,
       OBJ_POSITION:
       If (not ObjectList[i].Used) and (ObjectList[i].ParentObjID <> 99999999 )
         then
         with ObjectList[i] do
           StringList.Append(ObjType + ' '+ IntToStr(ObjNum) + IntToStr(ParentObjID)
                             + ' ' + ShortText + ' ' + LongText);
       end;
  end;

Procedure NumberRangeReport(StringList:TStrings);
  var
    i:LongInt;
  begin
     StringList.Append('Number Range check:');
     for i := low(ObjectList) to high(ObjectList) do
       case ObjectList[i].ObjType of
       // Look for Unused objects
       OBJ_EMPLOYEE:begin
                     if ((ObjectList[i].ObjNum <  HWMEEObjMin )
                      or
                      (ObjectList[i].ObjNum > HWMEEObjMax)) then
                        with ObjectList[i] do
                          StringList.Append(ObjType + ' '+ IntToStr(ObjNum)
                                            + ' ' + ShortText + ' ' + LongText);
                    end;
       OBJ_ORG_UNIT:begin
                     if ((ObjectList[i].ObjNum <   HWMOUObjMin)
                      or
                      (ObjectList[i].ObjNum > HWMOUObjMax)) then
                        with ObjectList[i] do
                          StringList.Append(ObjType + ' '+ IntToStr(ObjNum)
                                            + ' ' + ShortText + ' ' + LongText);

                    end;
       OBJ_POSITION:begin
                     if ((ObjectList[i].ObjNum <   HWMPosObjMin)
                      or
                      (ObjectList[i].ObjNum > HWMPosObjMax)) then
         with ObjectList[i] do
           StringList.Append(ObjType + ' '+ IntToStr(ObjNum)
                             + ' ' + ShortText + ' ' + LongText);
       end;
       end;
  end;

Function GetObjectById(Const ObjID:TObjId):TObjectEntry;
  var
    ThisObject:longint;
  begin
   Result := EmptyObj;
//   result.costctr := 0;
//   result.ParentObjID := 0;
 //  FillChar(Result, Result, #0);
   for ThisObject := low(ObjectList) to high(ObjectList) do
     if  ObjectList[ThisObject].ObjNum = ObjID then
     begin
       result := ObjectList[ThisObject];
       exit;
     end;
  end;

Function GetObjectIdxById(Const ObjID:TObjID):LongInt;
var
  i:longint;
begin
  result := 0;
  for i := low(ObjectList) to high(ObjectList) do
    if  ObjectList[i].ObjNum = ObjID then
      begin
        result := i;
        exit;
      end;
end;


Function GetObjectLongTextByID(Const ObjID:TObjID):UTF8String;
  var
    Obj:TObjectEntry;
  begin
     result := '';
     Obj := GetObjectById(ObjID);
     Result := Obj.LongText;
  end;

Function GetObjectShortTextByID(Const ObjID:TObjID):UTF8String;
  var
    Obj:TObjectEntry;
  begin
     result := '';
     Obj := GetObjectById(ObjID);
     Result := Obj.ShortText;
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
                 RelationshipList[j].Used := True;
         //        HasCostCtr := True;
                 CostCtr := RelationshipList[j].DestObjNum;
                 Used := True; // What if the cost center doesn't exist?
                 ObjectList[RelationshipList[j].DestObjIdx].Used := True;
                 ObjectList[RelationshipList[j].SrcObjIdx].Used := True;
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
         //        HasJob := True;
                 RelationshipList[j].Used := True;
                 Job := RelationshipList[j].DestObjNum;
                 Used := True; // What if destination Job doesn't exist?
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
               COL_A:begin // Object Type
                   FillChar(LineBuffer, SizeOf(LineBuffer), #0 );
                   //LineBuffer := Default(TObjectEntry);
                   If length(Parser.CurrentCellText) > 0 then
                    LineBuffer.ObjType:=Copy(Parser.CurrentCellText,1,1)[1];
                 end;
               COL_B:begin // Object ID
                   LineBuffer.ObjNum:=StrToInt64(Parser.CurrentCellText);
                   // Update high water mark if necessary
                   case LineBuffer.ObjType of
                     OBJ_ORG_UNIT: if LineBuffer.ObjNum > HWMOUObjNum then
                                       HWMOUObjNum := LineBuffer.ObjNum;
                     OBJ_POSITION: if LineBuffer.ObjNum > HWMPosObjNum then
                                       HWMPosObjNum := LineBuffer.ObjNum
                   end;  // of CASE
                 end;
               COL_C:begin // Begin Date
                   LineBuffer.BeginDate:=Parser.CurrentCellText;
                 end;
               COL_D:begin // End Date
                   LineBuffer.EndDate:=Parser.CurrentCellText;
                 end;
               COL_E:begin // Short Text Description
                    LineBuffer.ShortText := UTF8Copy(Parser.CurrentCellText,1,12);
                 end;
               COL_F:begin // Long Text Description
                   LineBuffer.LongText := UTF8Copy(Parser.CurrentCellText,1,40);
               end;
               COL_G:Begin        // deleimit date?
                 end;
               COL_H:Begin
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
               COL_A:begin // Source Object Type
                   FillChar(LineBuffer, SizeOf(LineBuffer), #0 );
                   If length(Parser.CurrentCellText) > 0 then
                    LineBuffer.SrcObjType:=Copy(Parser.CurrentCellText,1,1)[1];
                 end;
               COL_B:begin // Source Object ID
                   LineBuffer.SrcObjNum:=StrToInt64(Parser.CurrentCellText);
                 end;
               COL_C:begin // Relationship Direction + code
                   LineBuffer.Relationship:=Parser.CurrentCellText;
                 end;
               COL_D:begin // Begin Date
                   LineBuffer.BeginDate:=Parser.CurrentCellText;
                 end;
               COL_E:begin // End Date
                   LineBuffer.EndDate:=Parser.CurrentCellText;
                 end;
               COL_F:begin // Destination Object Type
                   If length(Parser.CurrentCellText) > 0 then
                    LineBuffer.DestObjType := Parser.CurrentCellText[1];
                 end;
               COL_G:begin // Destination Object Number
                   LineBuffer.DestObjNum := StrToInt64(Parser.CurrentCellText);
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

  end;    // PROCEDURE Import_ADP_HRP1001


Procedure Import_ADP_NHIRE(Const UTF16FileName:UTF8String);
// Important columns
// 16 Position
// 26 - 27 Romaji name
// 22-23 Kanji Name
// 70 last column
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
//     SetLength(ObjectList,0);
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
                // 0 Employee ID
                 COL_A:begin // Object Type / Object ID
                    FillChar(LineBuffer, SizeOf(LineBuffer), #0 );
                    If length(Parser.CurrentCellText) > 0 then
                       begin
                         LineBuffer.ObjType:=OBJ_EMPLOYEE;
                         LineBuffer.ObjNum :=StrToInt64(Parser.CurrentCellText);
                         // Update high water mark if necessary
                         if LineBuffer.ObjNum > HWMEEObjNum then
                           HWMEEObjNum := LineBuffer.ObjNum;
                       end;
                     end; // Case Col 0
                 COL_B:begin // Begin Date
                     LineBuffer.BeginDate:=Parser.CurrentCellText;
                   end;
                 COL_C:begin // End Date
                     LineBuffer.EndDate:=Parser.CurrentCellText;
                   end;
                 COL_Q:begin // Position
                     LineBuffer.ParentObjID :=StrToInt64(Parser.CurrentCellText);
                   end;
                 COL_W:begin // Kanji Family Name
                      LineBuffer.ShortText := UTF8Copy(Parser.CurrentCellText,1,12);
                    end;
                 COL_X:Begin // Kanji given name
                      LineBuffer.ShortText :=  LineBuffer.ShortText + ' '+ Parser.CurrentCellText;
                      LineBuffer.ShortText := UTF8Copy(LineBuffer.ShortText,1,12);
                    end;
                 COL_AA:begin // Romaji name
                      LineBuffer.LongText := UTF8Copy(Parser.CurrentCellText,1,40);
                   end;
                 COL_AB:begin // Romaji Name
                   LineBuffer.LongText := LineBuffer.LongText + ' ' + Parser.CurrentCellText;
                   LineBuffer.LongText := UTF8Copy(LineBuffer.LongText,1,40);
                 end;
                 COL_AW:Begin  // IT0041 Date 01 - last used field
//                   LineBuffer.LangCode := UTF8Copy(Parser.CurrentCellText,1,2);
                   LineBuffer.HasParent :=  (LineBuffer.ParentObjID <> 99999999);
                   LineBuffer.Stale := True;
                   LineBuffer.Chief := False;
//                   LineBuffer.ParentObjID := 0;
                   Inc(RecordsLoaded);
                   SetLength(ObjectList,Length(ObjectList)+1);
                   SendDebug(LineBuffer.ObjType
                             + IntToStr(LineBuffer.ObjNum)
                             + LineBuffer.BeginDate
                             + LineBuffer.EndDate
                             + LineBuffer.ShortText
                             + LineBuffer.LongText
                             + LineBuffer.LangCode
                             );
                   ObjectList[High(ObjectList)] := LineBuffer;

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


// Cost Center SSL file
Procedure Import_ADP_CCRMD(Const UTF16FileName:UTF8String);
// Important columns
// 0 language
// 1 controlling area
// 2 Cost Center
// 3 Company Code
// 4 Short Text
// 5 Long Text
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
//     SetLength(ObjectList,0);
     RecordsLoaded := 0;
     UTF8FileName := UTF16FiletoUTF8File(UTF16FileName);
     if not(FileExistsUTF8(UTF8FileName)) then exit;
     Parser:=TCSVParser.Create;
     FileStream := TFileStream.Create(UTF8ToANSI(UTF8Filename),
                                      fmOpenRead + fmShareDenyWrite);
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
                // 0 Employee ID
                 COL_A:begin // Language Code
                     FillChar(LineBuffer, SizeOf(LineBuffer), #0 );
                     LineBuffer.LangCode  :=Parser.CurrentCellText;
                   end;
                 COL_B:begin // Controlling Area
                   end;
                 COL_C:begin // Object Type / Object ID (Cost Center Code)
                     If length(Parser.CurrentCellText) > 0 then
                       begin
                         LineBuffer.ObjType:=OBJ_COSTCTR;
                         LineBuffer.ObjNum :=StrToInt64(Parser.CurrentCellText);
                         // Update high water mark if necessary
                         if LineBuffer.ObjNum > HWMCCObjNum then
                           HWMCCObjNum := LineBuffer.ObjNum;
                       end;
                     end; // Case Col 0
                 COL_D:begin // Short Text Description
                      LineBuffer.ShortText := UTF8Copy(Parser.CurrentCellText,1,12);
                   end;
                 COL_E:Begin  // IT0041 Date 01 - last used field
                   LineBuffer.LongText := UTF8Copy(Parser.CurrentCellText,1,40);
//                   LineBuffer.LangCode := UTF8Copy(Parser.CurrentCellText,1,2);
                   LineBuffer.HasParent :=  false; // No parents for CC
                   LineBuffer.Stale := False; // Cost Centers not shown in GUI
                   LineBuffer.Chief := False; // Chief is only relevant to positions
//                   LineBuffer.ParentObjID := 0;
                   Inc(RecordsLoaded);
                   SetLength(ObjectList,Length(ObjectList)+1);
                   SendDebug(LineBuffer.ObjType
                             + IntToStr(LineBuffer.ObjNum)
                             + LineBuffer.ShortText
                             + LineBuffer.LongText
                             + LineBuffer.LangCode
                             );
                   ObjectList[High(ObjectList)] := LineBuffer;

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

Function EmptyObj:TObjectEntry;
  begin
    Result.ObjType := ' ';
    Result.ObjNum := 0;
    Result.BeginDate := '';
    Result.EndDate := '';
    Result.ShortText:='';
    Result.LongText:='';
    Result.Dummy:= '';
    Result.LangCode:='';
    // Extra attributes for internal use
    Result.HasParent:=False;
    Result.ParentObjID:=0;
    Result.Stale:=True;
    Result.Chief:=False;
    Result.Job:=0;  // Job Code for Position
    Result.CostCtr:=0; // Cost Center for Position or Org. Unit
    Result.Used:=False;
  end;

initialization
  INI := TINIFile.Create('omz.ini');
  Init_Number_Ranges;
  Init_Lists;

Finalization
  Ini.Free;

end.

