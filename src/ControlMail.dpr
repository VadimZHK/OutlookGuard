program ControlMail;

uses
  Vcl.Forms,
  uMain in 'uMain.pas' {frmMain},
  uOutLookCommander in 'uOutLookCommander.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
