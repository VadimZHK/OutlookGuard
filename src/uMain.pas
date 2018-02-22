unit uMain;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls,
  uOutLookCommander, cxGraphics, cxControls, cxLookAndFeels,
  cxLookAndFeelPainters, dxBarBuiltInMenu, cxPC, cxContainer, cxEdit, cxLabel,
  Vcl.Menus, Vcl.StdCtrls, cxButtons;

type
  TfrmMain = class(TForm)
    Timer1: TTimer;
    pcMain: TcxPageControl;
    tsNewMail: TcxTabSheet;
    tsProgrammClose: TcxTabSheet;
    cxLabel1: TcxLabel;
    cxLabel2: TcxLabel;
    cxButton1: TcxButton;
    tsEmpty: TcxTabSheet;
    procedure Timer1Timer(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure cxButton1Click(Sender: TObject);
  private
    { Private declarations }
    olEventListener : TOutlookEventListener;
    procedure WndProc(var Msg: TMessage); override;
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;


implementation

{$R *.dfm}


procedure TfrmMain.cxButton1Click(Sender: TObject);
begin
  WindowState := wsMinimized;
end;

procedure TfrmMain.FormCreate(Sender: TObject);
begin
  pcMain.HideTabs := true;
end;

procedure TfrmMain.Timer1Timer(Sender: TObject);
begin
    if not TOutlookEventListener.IsOutlookRunning then exit;

    olEventListener := TOutlookEventListener.Create();
    TTimer(Sender).Enabled := false;
    olEventListener.NotifyHandle := self.Handle;
end;

procedure TfrmMain.WndProc(var Msg: TMessage);
begin
if Msg.WParam = SC_MINIMIZE then
  begin
    pcMain.ActivePage := tsEmpty;
//  MinimizeToSysTray();
  end;

  if Msg.Msg = WM_OUTLOOK_EVENT_CLOSE then
    begin
      Timer1.Enabled := true;
      pcMain.ActivePage := tsProgrammClose;
      WindowState := wsMinimized;
    end
  else
  if Msg.Msg = WM_OUTLOOK_EVENT_NEWMAIL then
    begin
      pcMain.ActivePage := tsNewMail;
      WindowState := wsMaximized;
      FormStyle := fsStayOnTop;
      Show;
    end
  else
    inherited;

end;

end.
