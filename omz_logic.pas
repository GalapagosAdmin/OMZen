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
    Job:LongInt;  // Job Code for Position
    CostCtr:LongInt; // Cost Center for Position or Org. Unit
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

Procedure Import_ADP_NHIRE(Const UTF16FileName:UTF8String);
// Imports NHIRE (Hire Action) Sheet, which has been exported from Excel
// by using Save As... > Unicode  (This produces a tab delimited UTF16LE text file)

Procedure Import_ADP_CCRMD(Const UTF16FileName:UTF8String);
// Imports NHIRE (Hire Action) Sheet, which has been exported from Excel
// by using Save As... > Unicode  (This produces a tab delimited UTF16LE text file)


Procedure UpdateHasParent;
Procedure UpdateCstCtr;
Procedure UpdateJob;
Procedure UpdateEE; // Update after loading NHIRE

Function GetObjectById(Const ObjID:LongInt):TObjectEntry;
Function GetObjectLongTextByID(Const ObjID:LongInt):UTF8String;
Function GetObjectShortTextByID(Const ObjID:LongInt):UTF8String;

Function DotStrip(const Original:UTF8String):UTF8String;


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
   HWMOUObjNum :=  12200000; // Should be read from settings
   HWMPosObjNum := 22200000; // Should be read from settings
   HWMEEObjNum :=  22600000; // Should be read from settings
   HWMCCObjNum :=  9990000000; // Should be read from settings
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

Function GetObjectLongTextByID(Const ObjID:LongInt):UTF8String;
  var
    Obj:TObjectEntry;
  begin
     result := '';
     Obj := GetObjectById(ObjID);
     Result := Obj.LongText;
  end;

Function GetObjectShortTextByID(Const ObjID:LongInt):UTF8String;
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
               COL_A:begin // Object Type
                   FillChar(LineBuffer, SizeOf(LineBuffer), #0 );
                   //LineBuffer := Default(TObjectEntry);
                   If length(Parser.CurrentCellText) > 0 then
                    LineBuffer.ObjType:=Copy(Parser.CurrentCellText,1,1)[1];
                 end;
               COL_B:begin // Object ID
                   LineBuffer.ObjNum:=StrToInt(Parser.CurrentCellText);
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

                   LineBuffer.SrcObjNum:=StrToInt(Parser.CurrentCellText);
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
                         LineBuffer.ObjNum :=StrToInt(Parser.CurrentCellText);
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
                     LineBuffer.ParentObjID :=StrToInt(Parser.CurrentCellText);
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
                         LineBuffer.ObjNum :=StrToInt(Parser.CurrentCellText);
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


initialization
  Init_Number_Ranges;
  Init_Lists;

end.

