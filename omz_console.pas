unit omz_console;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, StdCtrls;

type

  { TfrmConsole }

  TfrmConsole = class(TForm)
    mmConsole: TMemo;
  private
    { private declarations }
  public
    { public declarations }
  end;

var
  frmConsole: TfrmConsole;

implementation

{$R *.lfm}

end.

