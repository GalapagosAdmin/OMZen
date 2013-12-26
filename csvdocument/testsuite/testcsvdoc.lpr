program testcsvdoc;

{$mode objfpc}{$H+}

uses
  SysUtils, Classes, CsvDocument, DateUtils;

function ReadStringFromFile(AFileName: string): string;
var
  FileStream: TFileStream;
  Size: Integer;
begin
  Result := '';
  if not FileExists(AFileName) then
    Exit;
  FileStream := TFileStream.Create(AFileName, fmOpenRead);
  Size := FileStream.Size;
  if Size > 0 then
  begin
    SetLength(Result, Size);
    FileStream.ReadBuffer(Result[1], Size);
  end;
  FreeAndNil(FileStream);
end;

procedure FindTestFiles(AFileList: TStringList; const ASpec: string);
var
  SearchRec: TSearchRec;
  TestFilesPath: String;
begin
  AFileList.Clear;
  TestFilesPath := IncludeTrailingPathDelimiter(GetCurrentDir)
    + 'tests' + DirectorySeparator + ASpec + DirectorySeparator;
  if FindFirst(TestFilesPath + '*.csv', faAnyFile, SearchRec) = 0 then
    repeat
      AFileList.Add(TestFilesPath + SearchRec.Name);
    until FindNext(SearchRec) <> 0;
  FindClose(SearchRec);
end;

function TestCsvString(ADocument: TCSVDocument; const AnInput: String;
  out AParseTime, ABuildTime: Int64): String;
var
  Start: TDateTime;
begin
  ADocument.Clear;
  Start := Now;
  ADocument.CSVText := AnInput;
  AParseTime := MilliSecondsBetween(Start, Now);
  Start := Now;
  Result := ADocument.CSVText;
  ABuildTime := MilliSecondsBetween(Start, Now);
end;

procedure TestCsvFile(ADocument: TCSVDocument; const AnInputFilename,
  ASampleOutputFilename: String);
var
  InBuffer, OutBuffer: String;
  SampleBuffer: String;
  ParseTime: Int64;
  BuildTime: Int64;
begin
  InBuffer := ReadStringFromFile(AnInputFilename);
  SampleBuffer := ReadStringFromFile(ASampleOutputFilename);
  if SampleBuffer = '' then
    SampleBuffer := InBuffer;

  OutBuffer := TestCsvString(ADocument, InBuffer, ParseTime, BuildTime);

  Write(ExtractFileName(AnInputFilename));
  if OutBuffer = SampleBuffer then
  begin
    Write(': ok');
    WriteLn('   (parsed in ', BuildTime, ' ms)');
  end else
  begin
    WriteLn(': FAILED');
    WriteLn('--- Expected: ---');
    SampleBuffer := StringReplace(SampleBuffer, #13, '<CR>', [rfReplaceAll]);
    SampleBuffer := StringReplace(SampleBuffer, #10, '<LF>', [rfReplaceAll]);
    WriteLn(SampleBuffer);
    WriteLn('--- Got: --------');
    OutBuffer := StringReplace(OutBuffer, #13, '<CR>', [rfReplaceAll]);
    OutBuffer := StringReplace(OutBuffer, #10, '<LF>', [rfReplaceAll]);
    WriteLn(OutBuffer);
    WriteLn('-----------------');
  end;
end;

procedure ExecTests(ADocument: TCSVDocument; const ASpec: String);
var
  I: Integer;
  TestFiles: TStringList;
  CurrentTestFile: String;
begin
  WriteLn('== Format: ', ASpec, ' ==');
  TestFiles := TStringList.Create;
  FindTestFiles(TestFiles, ASpec);
  for I := 0 to TestFiles.Count - 1 do
  begin
    CurrentTestFile := TestFiles[I];
    TestCsvFile(ADocument,
      CurrentTestFile,
      ChangeFileExt(CurrentTestFile, '.sample'));
  end;
  FreeAndNil(TestFiles);
  WriteLn();
end;

procedure ExecPerformanceTest(ADocument: TCSVDocument; const AMinSizeKB: Integer);
var
  I, MaxRows: Integer;
  Seq: String;
  SeqLen: Integer;
  RealSize: Integer;
  InBuffer: String;
  ParseTime: Int64;
  BuildTime: Int64;
begin
  WriteLn('== Performance test: ==');

  WriteLn('Preparing the test...');
  Seq := '"abcd efg";"hij""k;l;m;n";opqrstuvw;    xyz12     ;"3456'#13#10'7890'#13#10;
  SeqLen := Length(Seq);
  MaxRows := ((AMinSizeKB * 1024) div SeqLen) + 1;

  InBuffer := '';
  RealSize := 0;
  for I := 0 to MaxRows do
  begin
    InBuffer := InBuffer + Seq;
    Inc(RealSize, SeqLen);
  end;

  WriteLn('Testing with ', RealSize div 1024, ' KB of CSV data...');
  TestCsvString(ADocument, InBuffer, ParseTime, BuildTime);
  WriteLn('Parsed in ', ParseTime, ' ms');
  WriteLn('Built in ', BuildTime, ' ms');
end;

procedure SetupUnofficialCsvHandler(AHandler: TCSVHandler);
begin
  AHandler.Delimiter := ',';
  AHandler.QuoteChar := '"';
  AHandler.LineEnding := #13#10;
  AHandler.IgnoreOuterWhitespace := True;
  AHandler.QuoteOuterWhitespace := True;
  AHandler.EqualColCountPerRow := False;
end;

var
  CsvDoc: TCSVDocument;
  Param1: String;
begin
  Param1 := ParamStr(1);
  WriteLn('Testing CSVDocument');
  WriteLn('-------------------');
  CsvDoc := TCSVDocument.Create;

  if Pos(Param1, 'performance') = 1 then
  begin
    CsvDoc.Delimiter := ';';
    CsvDoc.QuoteChar := '"';
    CsvDoc.LineEnding := #13#10;
    CsvDoc.EqualColCountPerRow := False;
    CsvDoc.IgnoreOuterWhitespace := False;
    CsvDoc.QuoteOuterWhitespace := True;
    ExecPerformanceTest(CsvDoc, StrToIntDef(ParamStr(2), 1));
  end else
  begin
    // no setup needed, rfc4180 supported out-of-the-box
    ExecTests(CsvDoc, 'rfc4180');

    // setup for unofficial Creativyst spec
    SetupUnofficialCsvHandler(CsvDoc);
    ExecTests(CsvDoc, 'unofficial');

    // setup for MS Excel files
    // (todo)
    ExecTests(CsvDoc, 'msexcel');

    // setup for OOo Calc files
    // (todo)
    ExecTests(CsvDoc, 'oocalc');
  end;

  FreeAndNil(CsvDoc);
  WriteLn('------------------');
  WriteLn('All tests complete');
end.

