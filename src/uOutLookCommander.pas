unit uOutLookCommander;

interface

uses Outlook_TLB, Windows, ActiveX, Messages;

type TOutlookEventListener = class(TInterfacedObject, IDispatch)
  strict private
    outlook : _Application;
    connectionPoint : IConnectionPoint;
    cookie : integer;


    function GetTypeInfoCount(out Count: Integer): HResult; stdcall;
    function GetTypeInfo(Index, LocaleID: Integer; out TypeInfo): HResult; stdcall;
    function GetIDsOfNames(const IID: TGUID; Names: Pointer;
                           NameCount, LocaleID: Integer; DispIDs: Pointer): HResult; stdcall;
    function Invoke(DispID: Integer; const IID: TGUID; LocaleID: Integer;
                    Flags: Word; var Params; VarResult, ExcepInfo, ArgErr: Pointer): HResult; stdcall;

  public
    notifyHandle : HWND;
    constructor Create();

    procedure newMailEventAction();
    procedure CloseEventAction();

    class function IsOutlookRunning():boolean;
end;

const WM_OUTLOOK_EVENT_CLOSE   = WM_USER + 1;
      WM_OUTLOOK_EVENT_NEWMAIL = WM_USER + 2;

implementation

uses ComObj;

{ TOutlookEventListener }

procedure TOutlookEventListener.CloseEventAction;
begin
    try
        connectionPoint.Unadvise(cookie);
        connectionPoint := nil;

        PostMessage(notifyHandle, WM_OUTLOOK_EVENT_CLOSE, 0, 0);
    except
        {$IFDEF DEBUG}
      //  MessageDlg('Santa error! cant unadvise connection point', mtError, [mbOk],0);
        {$ENDIF}
    end;
end;

constructor TOutlookEventListener.Create;
var cpc : IConnectionPointContainer;
    ol : IDispatch;
begin
    inherited;

    ol := GetActiveOleObject('Outlook.Application');

    cpc := ol as IConnectionPointContainer;
    cpc.FindConnectionPoint(DIID_ApplicationEvents, connectionPoint);
    connectionPoint.Advise(self, cookie);end;

function TOutlookEventListener.GetIDsOfNames(const IID: TGUID; Names: Pointer;
  NameCount, LocaleID: Integer; DispIDs: Pointer): HResult;
begin
   result := E_NOTIMPL;
end;

function TOutlookEventListener.GetTypeInfo(Index, LocaleID: Integer;
  out TypeInfo): HResult;
begin
   pointer(TypeInfo) := nil;
    result := E_NOTIMPL;
end;

function TOutlookEventListener.GetTypeInfoCount(out Count: Integer): HResult;
begin
   count := 1;
    result := E_NOTIMPL;
end;

function TOutlookEventListener.Invoke(DispID: Integer; const IID: TGUID;
  LocaleID: Integer; Flags: Word; var Params; VarResult, ExcepInfo,
  ArgErr: Pointer): HResult;
begin
    result := S_OK;

    case DispId of
        61442 : ; // ItemSend(const Item: IDispatch; var Cancel: WordBool);
        61443 : newMailEventAction();
        61444 : ; // Reminder(const Item: IDispatch);
        61445 : ; // OptionsPagesAdd(const Pages: PropertyPages);
        61446 : ; // Startup;
        61447 : CloseEventAction();
        else
            result := E_INVALIDARG;
    end;
end;

class function TOutlookEventListener.IsOutlookRunning: boolean;
begin
    result := true;
    try
        getActiveOleObject('Outlook.Application');
    except
        result := false;
    end;

end;

procedure TOutlookEventListener.newMailEventAction;
begin
  PostMessage(notifyHandle, WM_OUTLOOK_EVENT_NEWMAIL, 0, 0);
end;

end.
