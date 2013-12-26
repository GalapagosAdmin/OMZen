program csvdemo;

{$mode objfpc}{$H+}

uses
  {$IFDEF UNIX}{$IFDEF UseCThreads}
  cthreads,
  {$ENDIF}{$ENDIF}
  Interfaces, // this includes the LCL widgetset
  Forms, mainfrm;

{$R csvdemo.res}

begin
  Application.Title := 'CsvDemo';
  Application.Initialize;
  Application.CreateForm(TfmCSVTest, fmCSVTest);
  Application.Run;
end.

