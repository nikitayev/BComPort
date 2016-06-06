/////////////////////////////////////////////////////////////////////////////
//                                                                         //
//              TBComPort  ver.2.10  -  24.11.2005      freeware           //
//  ---------------------------------------------------------------------  //
//  Компонент для обмена данными с внешними устройствами через интерфейс   //
//  RS-232 в асинхронном или синхронном режиме. Разработан на основе       //
//  библиотеки ComPort Library от Dejan Crnila.                            //
//            Работает с Delphi 2..7 под Windows 9X/ME/NT4/2K/XP.          //
//  ---------------------------------------------------------------------  //
//                 (c) 2005 Брусникин И.В.    majar@nm.ru                  //
//                                                                         //
/////////////////////////////////////////////////////////////////////////////

unit BCPort;

interface

uses
  Windows, Messages, SysUtils, Classes;

{$B-,H+,X+}

{$IFDEF VER140}
  {$DEFINE D6UP}
{$ENDIF}

{$IFDEF VER150}
  {$DEFINE D6UP}
{$ENDIF}

type
  TBaudRate = (br110, br300, br600, br1200, br2400, br4800, br9600, br14400,
    br19200, br38400, br56000, br57600, br115200, br128000, br256000);
  TByteSize = (bs5, bs6, bs7, bs8);
  TComErrors = set of (ceFrame, ceRxParity, ceOverrun, ceBreak, ceIO, ceMode,
                       ceRxOver, ceTxFull);
  TComEvents = set of (evRxChar, evTxEmpty, evRing, evCTS, evDSR, evRLSD,
                       evError, evRx80Full);
  TComSignals = set of (csCTS, csDSR, csRing, csRLSD);
  TParity = (paNone, paOdd, paEven, paMark, paSpace);
  TStopBits = (sb1, sb1_5, sb2);
  TSyncMethod = (smThreadSync, smWindowSync, smNone);
  TComSignalEvent = procedure(Sender: TObject; State: Boolean) of object;
  TComErrorEvent = procedure(Sender: TObject; Errors: TComErrors) of object;
  TRxCharEvent = procedure(Sender: TObject; Count: Integer) of object;

  TOperationKind = (okWrite, okRead);
  TAsync = record
    Overlapped: TOverlapped;
    Kind: TOperationKind;
    Data: Pointer;
    Size: Integer;
  end;
  PAsync = ^TAsync;

  TBComPort = class;

  TComThread = class(TThread)
  private
    FComPort: TBComPort;
    FEvents: TComEvents;
    FStopEvent: THandle;
  protected
    procedure DoEvents;
    procedure Execute; override;
    procedure SendEvents;
    procedure Stop;
  public
    constructor Create(AComPort: TBComPort);
    destructor Destroy; override;
  end;

  TComTimeouts = class(TPersistent)
  private
    FComPort: TBComPort;
    FReadInterval: Integer;
    FReadTotalM: Integer;
    FReadTotalC: Integer;
    FWriteTotalM: Integer;
    FWriteTotalC: Integer;
    procedure SetComPort(const AComPort: TBComPort);
    procedure SetReadInterval(const Value: Integer);
    procedure SetReadTotalM(const Value: Integer);
    procedure SetReadTotalC(const Value: Integer);
    procedure SetWriteTotalM(const Value: Integer);
    procedure SetWriteTotalC(const Value: Integer);
  protected
    procedure AssignTo(Dest: TPersistent); override;
  public
    constructor Create;
    property ComPort: TBComPort read FComPort;
  published
    property ReadInterval: Integer read FReadInterval write SetReadInterval;
    property ReadTotalMultiplier: Integer read FReadTotalM write SetReadTotalM;
    property ReadTotalConstant: Integer read FReadTotalC write SetReadTotalC;
    property WriteTotalMultiplier: Integer
      read FWriteTotalM write SetWriteTotalM;
    property WriteTotalConstant: Integer
      read FWriteTotalC write SetWriteTotalC;
  end;

  TBComPort = class(TComponent)
  private
    FBaudRate: TBaudRate;
    FByteSize: TByteSize;
    FConnected: Boolean;
    FCTPriority: TThreadPriority;
    FEvents: TComEvents;
    FEventThread: TComThread;
    FHandle: THandle;
    FInBufSize: Integer;
    FOutBufSize: Integer;
    FParity: TParity;
    FPort: ansistring;
    FStopBits: TStopBits;
    FSyncMethod: TSyncMethod;
    FTimeouts: TComTimeouts;
    FUpdate: Boolean;
    FWindow: THandle;
    FOnCTSChange: TComSignalEvent;
    FOnDSRChange: TComSignalEvent;
    FOnError: TComErrorEvent;
    FOnRing: TNotifyEvent;
    FOnRLSDChange: TComSignalEvent;
    FOnRx80Full: TNotifyEvent;
    FOnRxChar: TRxCharEvent;
    FOnTxEmpty: TNotifyEvent;
    procedure CallCTSChange;
    procedure CallDSRChange;
    procedure CallError;
    procedure CallRing;
    procedure CallRLSDChange;
    procedure CallRx80Full;
    procedure CallRxChar;
    procedure CallTxEmpty;
    procedure SetBaudRate(const Value: TBaudRate);
    procedure SetByteSize(const Value: TByteSize);
    procedure SetCTPriority(const Value: TThreadPriority);
    procedure SetInBufSize(const Value: Integer);
    procedure SetOutBufSize(const Value: Integer);
    procedure SetParity(const Value: TParity);
    procedure SetPort(const Value: ansistring);
    procedure SetStopBits(const Value: TStopBits);
    procedure SetSyncMethod(const Value: TSyncMethod);
    procedure SetTimeouts(const Value: TComTimeouts);
    procedure WindowMethod(var Message: TMessage);
  protected
    procedure ApplyBuffer;
    procedure ApplyDCB;
    procedure ApplyTimeouts;
    procedure CreateHandle;
    procedure DestroyHandle;
    procedure SetupComPort;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure AbortAllAsync;
    procedure BeginUpdate;
    procedure ClearBuffer(Input, Output: Boolean);
    function Close: Boolean;
    procedure EndUpdate;
    function InBufCount: Integer;
    function IsAsyncCompleted(AsyncPtr: PAsync): Boolean;
    function Open: Boolean;
    function OutBufCount: Integer;
    function Read(var Buffer; Count: Integer): Integer;
    function ReadAsync(var Buffer; Count: Integer; var AsyncPtr: PAsync): Integer;
    function ReadStr(var Str: ansistring; Count: Integer): Integer;
    function ReadStrAsync(var Str: ansistring; Count: Integer; var AsyncPtr: PAsync): Integer;
    procedure SetDTR(State: Boolean);
    procedure SetRTS(State: Boolean);
    function Signals: TComSignals;
    function WaitForAsync(var AsyncPtr: PAsync): Integer;
    function Write(const Buffer; Count: Integer): Integer;
    function WriteAsync(const Buffer; Count: Integer; var AsyncPtr: PAsync): Integer;
    function WriteStr(const Str: ansistring): Integer;
    function WriteStrAsync(const Str: ansistring; var AsyncPtr: PAsync): Integer;
    property Connected: Boolean read FConnected;
    property CTPriority: TThreadPriority read FCTPriority write SetCTPriority;
  published
    property BaudRate: TBaudRate read FBaudRate write SetBaudRate;
    property ByteSize: TByteSize read FByteSize write SetByteSize;
    property InBufSize: Integer read FInBufSize write SetInBufSize;
    property OutBufSize: Integer read FOutBufSize write SetOutBufSize;
    property Parity: TParity read FParity write SetParity;
    property Port: ansistring read FPort write SetPort;
    property SyncMethod: TSyncMethod read FSyncMethod write SetSyncMethod;
    property StopBits: TStopBits read FStopBits write SetStopBits;
    property Timeouts: TComTimeouts read FTimeouts write SetTimeouts;
    property OnCTSChange: TComSignalEvent read FOnCTSChange write FOnCTSChange;
    property OnDSRChange: TComSignalEvent read FOnDSRChange write FOnDSRChange;
    property OnError: TComErrorEvent read FOnError write FOnError;
    property OnRing: TNotifyEvent read FOnRing write FOnRing;
    property OnRLSDChange: TComSignalEvent read FOnRLSDChange write FOnRLSDChange;
    property OnRx80Full: TNotifyEvent read FOnRx80Full write FOnRx80Full;
    property OnRxChar: TRxCharEvent read FOnRxChar write FOnRxChar;
    property OnTxEmpty: TNotifyEvent read FOnTxEmpty write FOnTxEmpty;
  end;

  EComPort = class(Exception);

procedure InitAsync(var AsyncPtr: PAsync);
procedure DoneAsync(var AsyncPtr: PAsync);
procedure EnumComPorts(Ports: TStrings);

procedure Register;

implementation

uses
  Forms;

const
  CM_COMPORT = WM_USER + 1;

  CEMess: array[1..15] of ansistring =
    ('Попытка открыть несуществующий COM-порт',
     'Порт занят другой программой',
     'Ошибка записи в порт',
     'Ошибка чтения из порта',
     'Недопустимый параметр Async',
     'Ошибка функции PurgeComm',
     'Не удалось получить статус асинхронной операции',
     'Ошибка функции SetCommState',
     'Ошибка функции SetCommTimeouts',
     'Ошибка функции SetupComm',
     'Ошибка функции ClearCommError',
     'Ошибка функции GetCommModemStatus',
     'Ошибка функции EscapeCommFunction',
     'Нельзя менять свойство, пока порт открыт',
     'Ошибка обращения к системному реестру');

function EventsToInt(const Events: TComEvents): Integer;
begin
  Result := 0;
  if evRxChar in Events then Result := Result or EV_RXCHAR;
  if evTxEmpty in Events then Result := Result or EV_TXEMPTY;
  if evRing in Events then Result := Result or EV_RING;
  if evCTS in Events then Result := Result or EV_CTS;
  if evDSR in Events then Result := Result or EV_DSR;
  if evRLSD in Events then Result := Result or EV_RLSD;
  if evError in Events then Result := Result or EV_ERR;
  if evRx80Full in Events then Result := Result or EV_RX80FULL;
end;

function IntToEvents(Mask: Integer): TComEvents;
begin
  Result := [];
  if (EV_RXCHAR and Mask) <> 0 then Result := Result + [evRxChar];
  if (EV_TXEMPTY and Mask) <> 0 then Result := Result + [evTxEmpty];
  if (EV_RING and Mask) <> 0 then Result := Result + [evRing];
  if (EV_CTS and Mask) <> 0 then Result := Result + [evCTS];
  if (EV_DSR and Mask) <> 0 then Result := Result + [evDSR];
  if (EV_RLSD and Mask) <> 0 then Result := Result + [evRLSD];
  if (EV_ERR and Mask) <> 0 then Result := Result + [evError];
  if (EV_RX80FULL and Mask) <> 0 then Result := Result + [evRx80Full];
end;

{ TComThread }

constructor TComThread.Create(AComPort: TBComPort);
begin
  inherited Create(True);
  FStopEvent := CreateEvent(nil, True, False, nil);
  FComPort := AComPort;
  Priority := FComPort.CTPriority;
  SetCommMask(FComPort.FHandle, EventsToInt(FComPort.FEvents));
  Resume;
end;

destructor TComThread.Destroy;
begin
  Stop;
  inherited Destroy;
end;

procedure TComThread.Execute;
var
  EventHandles: array[0..1] of THandle;
  Overlapped: TOverlapped;
  Signaled, BytesTrans, Mask: DWORD;
begin
  FillChar(Overlapped, SizeOf(Overlapped), 0);
  Overlapped.hEvent := CreateEvent(nil, True, True, nil);
  EventHandles[0] := FStopEvent;
  EventHandles[1] := Overlapped.hEvent;
  repeat
    WaitCommEvent(FComPort.FHandle, Mask, @Overlapped);
    Signaled := WaitForMultipleObjects(2, @EventHandles, False, INFINITE);
    if (Signaled = WAIT_OBJECT_0 + 1) and
      GetOverlappedResult(FComPort.FHandle, Overlapped, BytesTrans, False) then
    begin
      FEvents := IntToEvents(Mask);
      case FComPort.SyncMethod of
        smThreadSync: Synchronize(DoEvents);
        smWindowSync: SendEvents;
        smNone      : DoEvents;
      end;
    end;
  until Signaled <> (WAIT_OBJECT_0 + 1);
  SetCommMask(FComPort.FHandle, 0);
  PurgeComm(FComPort.FHandle, PURGE_TXCLEAR or PURGE_RXCLEAR);
  CloseHandle(Overlapped.hEvent);
  CloseHandle(FStopEvent);
end;

procedure TComThread.Stop;
begin
  SetEvent(FStopEvent);
  Sleep(0);
end;

procedure TComThread.SendEvents;
begin
  if evError in FEvents then
    SendMessage(FComPort.FWindow, CM_COMPORT, EV_ERR, 0);
  if evRxChar in FEvents then
    SendMessage(FComPort.FWindow, CM_COMPORT, EV_RXCHAR, 0);
  if evTxEmpty in FEvents then
    SendMessage(FComPort.FWindow, CM_COMPORT, EV_TXEMPTY, 0);
  if evRing in FEvents then
    SendMessage(FComPort.FWindow, CM_COMPORT, EV_RING, 0);
  if evCTS in FEvents then
    SendMessage(FComPort.FWindow, CM_COMPORT, EV_CTS, 0);
  if evDSR in FEvents then
    SendMessage(FComPort.FWindow, CM_COMPORT, EV_DSR, 0);
  if evRing in FEvents then
    SendMessage(FComPort.FWindow, CM_COMPORT, EV_RLSD, 0);
  if evRx80Full in FEvents then
    SendMessage(FComPort.FWindow, CM_COMPORT, EV_RX80FULL, 0);
end;

procedure TComThread.DoEvents;
begin
  if evError in FEvents then FComPort.CallError;
  if evRxChar in FEvents then FComPort.CallRxChar;
  if evTxEmpty in FEvents then FComPort.CallTxEmpty;
  if evRing in FEvents then FComPort.CallRing;
  if evCTS in FEvents then FComPort.CallCTSChange;
  if evDSR in FEvents then FComPort.CallDSRChange;
  if evRLSD in FEvents then FComPort.CallRLSDChange;
  if evRx80Full in FEvents then FComPort.CallRx80Full;
end;

{ TComTimeouts }

constructor TComTimeouts.Create;
begin
  inherited Create;
  FReadInterval := -1;
  FWriteTotalM := 100;
  FWriteTotalC := 1000;
end;

procedure TComTimeouts.AssignTo(Dest: TPersistent);
begin
  if Dest is TComTimeouts then
  begin
    with TComTimeouts(Dest) do
    begin
      FReadInterval := Self.ReadInterval;
      FReadTotalM   := Self.ReadTotalMultiplier;
      FReadTotalC   := Self.ReadTotalConstant;
      FWriteTotalM  := Self.WriteTotalMultiplier;
      FWriteTotalC  := Self.WriteTotalConstant;
    end;
  end
  else
    inherited AssignTo(Dest);
end;

procedure TComTimeouts.SetComPort(const AComPort: TBComPort);
begin
  FComPort := AComPort;
end;

procedure TComTimeouts.SetReadInterval(const Value: Integer);
begin
  if Value <> FReadInterval then
  begin
    FReadInterval := Value;
    FComPort.ApplyTimeouts;
  end;
end;

procedure TComTimeouts.SetReadTotalC(const Value: Integer);
begin
  if Value <> FReadTotalC then
  begin
    FReadTotalC := Value;
    FComPort.ApplyTimeouts;
  end;
end;

procedure TComTimeouts.SetReadTotalM(const Value: Integer);
begin
  if Value <> FReadTotalM then
  begin
    FReadTotalM := Value;
    FComPort.ApplyTimeouts;
  end;
end;

procedure TComTimeouts.SetWriteTotalC(const Value: Integer);
begin
  if Value <> FWriteTotalC then
  begin
    FWriteTotalC := Value;
    FComPort.ApplyTimeouts;
  end;
end;

procedure TComTimeouts.SetWriteTotalM(const Value: Integer);
begin
  if Value <> FWriteTotalM then
  begin
    FWriteTotalM := Value;
    FComPort.ApplyTimeouts;
  end;
end;

{ TBComPort }

constructor TBComPort.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FComponentStyle := FComponentStyle - [csInheritable];
  FBaudRate := br9600;
  FByteSize := bs8;
  FConnected := False;
  FCTPriority := tpNormal;
  FEvents := [evRxChar, evTxEmpty, evRing, evCTS, evDSR, evRLSD, evError,
              evRx80Full];
  FHandle := INVALID_HANDLE_VALUE;
  FInBufSize := 2048;
  FOutBufSize := 2048;
  FParity := paNone;
  FPort := 'COM2';
  FStopBits := sb1;
  FSyncMethod := smThreadSync;
  FTimeouts := TComTimeouts.Create;
  FTimeouts.SetComPort(Self);
  FUpdate := True;
end;

destructor TBComPort.Destroy;
begin
  Close;
  FTimeouts.Free;
  inherited Destroy;
end;

procedure TBComPort.CreateHandle;
var
  zPortName: widestring;
begin
  zPortName := '\\.\' + FPort;
  FHandle := CreateFile(PWideChar(zPortName), GENERIC_READ or GENERIC_WRITE,
                        0, nil, OPEN_EXISTING, FILE_FLAG_OVERLAPPED, 0);
  if FHandle = INVALID_HANDLE_VALUE then
  begin
    if GetLastError = ERROR_FILE_NOT_FOUND then
      raise EComPort.Create(CEMess[1])
    else if GetLastError = ERROR_ACCESS_DENIED then
      raise EComPort.Create(CEMess[2]);
  end;
end;

procedure TBComPort.DestroyHandle;
begin
  if FHandle <> INVALID_HANDLE_VALUE then CloseHandle(FHandle);
end;

procedure TBComPort.WindowMethod(var Message: TMessage);
begin
  with Message do
    if Msg = CM_COMPORT then
      try
        if InSendMessage then ReplyMessage(0);
        if FConnected then
          case wParam of
            EV_CTS:      CallCTSChange;
            EV_DSR:      CallDSRChange;
            EV_RING:     CallRing;
            EV_RLSD:     CallRLSDChange;
            EV_RX80FULL: CallRx80Full;
            EV_RXCHAR:   CallRxChar;
            EV_ERR:      CallError;
            EV_TXEMPTY:  CallTxEmpty;
          end;
      except
        Application.HandleException(Self);
      end
    else
      Result := DefWindowProc(FWindow, Msg, wParam, lParam);
end;

procedure TBComPort.BeginUpdate;
begin
  FUpdate := False;
end;

procedure TBComPort.EndUpdate;
begin
  if not FUpdate then FUpdate := True;
  SetupComPort;
end;

function TBComPort.Open: Boolean;
begin
  if not FConnected then
  begin
    CreateHandle;
    FConnected := True;
    try
      SetupComPort;
    except
      DestroyHandle;
      FConnected := False;
      raise;
    end;
    if (FSyncMethod = smWindowSync) then
      {$IFDEF D6UP}
        {$WARN SYMBOL_DEPRECATED OFF}
      {$ENDIF}
      FWindow := AllocateHWnd(WindowMethod);
      {$IFDEF D6UP}
        {$WARN SYMBOL_DEPRECATED ON}
      {$ENDIF}
    FEventThread := TComThread.Create(Self);
  end;
  Result := FConnected;
end;

function TBComPort.Close: Boolean;
begin
  if FConnected then
  begin
    SetDTR(False);
    SetRTS(False);
    AbortAllAsync;
    FEventThread.Free;
    if FSyncMethod = smWindowSync then
      {$IFDEF D6UP}
        {$WARN SYMBOL_DEPRECATED OFF}
      {$ENDIF}
      DeallocateHWnd(FWindow);
      {$IFDEF D6UP}
        {$WARN SYMBOL_DEPRECATED ON}
      {$ENDIF}
    DestroyHandle;
    FConnected := False;
  end;
  Result := not FConnected;
end;

procedure TBComPort.ApplyDCB;
const
  CBaudRate: array[TBaudRate] of Integer = (CBR_110, CBR_300, CBR_600,
    CBR_1200, CBR_2400, CBR_4800, CBR_9600, CBR_14400, CBR_19200, CBR_38400,
    CBR_56000, CBR_57600, CBR_115200, CBR_128000, CBR_256000);
var
  DCB: TDCB;
begin
  if FConnected and FUpdate then
  begin
    FillChar(DCB, SizeOf(TDCB), 0);
    DCB.DCBlength := SizeOf(TDCB);
    DCB.BaudRate := CBaudRate[FBaudRate];
    DCB.ByteSize := Ord(TByteSize(FByteSize)) + 5;
    DCB.Flags := 1 or ($30 and (DTR_CONTROL_ENABLE shl 4))
                   or ($3000 and (RTS_CONTROL_ENABLE shl 12));
    if FParity <> paNone then
      DCB.Flags := DCB.Flags or 2;
    DCB.Parity   := Ord(TParity(FParity));
    DCB.StopBits := Ord(TStopBits(FStopBits));
    DCB.XonChar  := #17;
    DCB.XoffChar := #19;
    if not SetCommState(FHandle, DCB) then
      raise EComPort.Create(CEMess[8]);
  end;
end;

procedure TBComPort.ApplyTimeouts;
var
  Timeouts: TCommTimeouts;

  function MValue(const Value: Integer): DWORD;
  begin
    if Value < 0 then Result := MAXDWORD else Result := Value;
  end;

begin
  if FConnected and FUpdate then
  begin
    Timeouts.ReadIntervalTimeout := MValue(FTimeouts.ReadInterval);
    Timeouts.ReadTotalTimeoutMultiplier := MValue(FTimeouts.ReadTotalMultiplier);
    Timeouts.ReadTotalTimeoutConstant := MValue(FTimeouts.ReadTotalConstant);
    Timeouts.WriteTotalTimeoutMultiplier := MValue(FTimeouts.WriteTotalMultiplier);
    Timeouts.WriteTotalTimeoutConstant := MValue(FTimeouts.WriteTotalConstant);
    if not SetCommTimeouts(FHandle, Timeouts) then
      raise EComPort.Create(CEMess[9]);
  end;
end;

procedure TBComPort.ApplyBuffer;
begin
  if FConnected and FUpdate then
    if not SetupComm(FHandle, FInBufSize, FOutBufSize) then
      raise EComPort.Create(CEMess[10]);
end;

procedure TBComPort.SetupComPort;
begin
  ApplyBuffer;
  ApplyDCB;
  ApplyTimeouts;
end;

function TBComPort.InBufCount: Integer;
var
  Errors: DWORD;
  ComStat: TComStat;
begin
  if not ClearCommError(FHandle, Errors, @ComStat) then
    raise EComPort.Create(CEMess[11]);
  Result := ComStat.cbInQue;
end;

function TBComPort.OutBufCount: Integer;
var
  Errors: DWORD;
  ComStat: TComStat;
begin
  if not ClearCommError(FHandle, Errors, @ComStat) then
    raise EComPort.Create(CEMess[11]);
  Result := ComStat.cbOutQue;
end;

function TBComPort.Signals: TComSignals;
var
  Status: DWORD;
begin
  if not GetCommModemStatus(FHandle, Status) then
    raise EComPort.Create(CEMess[12]);
  Result := [];
  if (MS_CTS_ON and Status) <> 0 then Result := Result + [csCTS];
  if (MS_DSR_ON and Status) <> 0 then Result := Result + [csDSR];
  if (MS_RING_ON and Status) <> 0 then Result := Result + [csRing];
  if (MS_RLSD_ON and Status) <> 0 then Result := Result + [csRLSD];
end;

procedure TBComPort.SetDTR(State: Boolean);
var
  Act: DWORD;
begin
  if State then Act := Windows.SETDTR else Act := Windows.CLRDTR;
  if not EscapeCommFunction(FHandle, Act) then
    raise EComPort.Create(CEMess[13]);
end;

procedure TBComPort.SetRTS(State: Boolean);
var
  Act: DWORD;
begin
  if State then Act := Windows.SETRTS else Act := Windows.CLRRTS;
  if not EscapeCommFunction(FHandle, Act) then
    raise EComPort.Create(CEMess[13]);
end;

procedure TBComPort.ClearBuffer(Input, Output: Boolean);
var
  Flag: DWORD;
begin
  Flag := 0;
  if Input then
    Flag := PURGE_RXCLEAR;
  if Output then
    Flag := Flag or PURGE_TXCLEAR;
  if not PurgeComm(FHandle, Flag) then
    raise EComPort.Create(CEMess[6]);
end;

procedure PrepareAsync(AKind: TOperationKind; const Buffer;
  Count: Integer; AsyncPtr: PAsync);
begin
  with AsyncPtr^ do
  begin
    Kind := AKind;
    if Data <> nil then FreeMem(Data);
    GetMem(Data, Count);
    Move(Buffer, Data^, Count);
    Size := Count;
  end;
end;

function TBComPort.WriteAsync(const Buffer; Count: Integer; var AsyncPtr: PAsync): Integer;
var
  Success: Boolean;
  BytesTrans: DWORD;
begin
  if AsyncPtr = nil then
    raise EComPort.Create(CEMess[5]);
  PrepareAsync(okWrite, Buffer, Count, AsyncPtr);
  Success := WriteFile(FHandle, Buffer, Count, BytesTrans, @AsyncPtr^.Overlapped)
    or (GetLastError = ERROR_IO_PENDING);
  if not Success then
    raise EComPort.Create(CEMess[3]);
  Result := BytesTrans;
end;

function TBComPort.Write(const Buffer; Count: Integer): Integer;
var
  AsyncPtr: PAsync;
begin
  InitAsync(AsyncPtr);
  try
    WriteAsync(Buffer, Count, AsyncPtr);
    Result := WaitForAsync(AsyncPtr);
  finally
    DoneAsync(AsyncPtr);
  end;
end;

function TBComPort.WriteStrAsync(const Str: ansistring; var AsyncPtr: PAsync): Integer;
begin
  if Length(Str) > 0 then
    Result := WriteAsync(Str[1], Length(Str), AsyncPtr)
  else
    Result := 0;
end;

function TBComPort.WriteStr(const Str: ansistring): Integer;
var
  AsyncPtr: PAsync;
begin
  InitAsync(AsyncPtr);
  try
    WriteStrAsync(Str, AsyncPtr);
    Result := WaitForAsync(AsyncPtr);
  finally
    DoneAsync(AsyncPtr);
  end;
end;

function TBComPort.ReadAsync(var Buffer; Count: Integer; var AsyncPtr: PAsync): Integer;
var
  Success: Boolean;
  BytesTrans: DWORD;
begin
  if AsyncPtr = nil then
    raise EComPort.Create(CEMess[5]);
  AsyncPtr^.Kind := okRead;
  Success := ReadFile(FHandle, Buffer, Count, BytesTrans, @AsyncPtr^.Overlapped)
    or (GetLastError = ERROR_IO_PENDING);
  if not Success then
    raise EComPort.Create(CEMess[4]);
  Result := BytesTrans;
end;

function TBComPort.Read(var Buffer; Count: Integer): Integer;
var
  AsyncPtr: PAsync;
begin
  InitAsync(AsyncPtr);
  try
    ReadAsync(Buffer, Count, AsyncPtr);
    Result := WaitForAsync(AsyncPtr);
  finally
    DoneAsync(AsyncPtr);
  end;
end;

function TBComPort.ReadStrAsync(var Str: ansistring; Count: Integer; var AsyncPtr: PAsync): Integer;
begin
  SetLength(Str, Count);
  if Count > 0 then
    Result := ReadAsync(Str[1], Count, AsyncPtr)
  else
    Result := 0;
end;

function TBComPort.ReadStr(var Str: ansistring; Count: Integer): Integer;
var
  AsyncPtr: PAsync;
begin
  InitAsync(AsyncPtr);
  try
    ReadStrAsync(Str, Count, AsyncPtr);
    Result := WaitForAsync(AsyncPtr);
    SetLength(Str, Result);
  finally
    DoneAsync(AsyncPtr);
  end;
end;

function ErrorCode(AsyncPtr: PAsync): Integer;
begin
  if AsyncPtr^.Kind = okRead then Result := 4 else Result := 3;
end;

function TBComPort.WaitForAsync(var AsyncPtr: PAsync): Integer;
var
  BytesTrans, Signaled: DWORD;
  Success: Boolean;
begin
  if AsyncPtr = nil then
    raise EComPort.Create(CEMess[5]);
  Signaled := WaitForSingleObject(AsyncPtr^.Overlapped.hEvent, INFINITE);
  Success := (Signaled = WAIT_OBJECT_0) and
      (GetOverlappedResult(FHandle, AsyncPtr^.Overlapped, BytesTrans, False));
  if not Success then
    raise EComPort.Create(CEMess[ErrorCode(AsyncPtr)]);
  Result := BytesTrans;
end;

procedure TBComPort.AbortAllAsync;
begin
  if not PurgeComm(FHandle, PURGE_TXABORT or PURGE_RXABORT) then
    raise EComPort.Create(CEMess[6]);
end;

function TBComPort.IsAsyncCompleted(AsyncPtr: PAsync): Boolean;
var
  BytesTrans: DWORD;
begin
  if AsyncPtr = nil then
    raise EComPort.Create(CEMess[5]);
  Result := GetOverlappedResult(FHandle, AsyncPtr^.Overlapped, BytesTrans, False);
  if not Result then
    if (GetLastError <> ERROR_IO_PENDING) and (GetLastError <> ERROR_IO_INCOMPLETE) then
      raise EComPort.Create(CEMess[7]);
end;

procedure TBComPort.CallCTSChange;
begin
  if Assigned(FOnCTSChange) then FOnCTSChange(Self, csCTS in Signals);
end;

procedure TBComPort.CallDSRChange;
begin
  if Assigned(FOnDSRChange) then FOnDSRChange(Self, csDSR in Signals);
end;

procedure TBComPort.CallRLSDChange;
begin
  if Assigned(FOnRLSDChange) then FOnRLSDChange(Self, csRLSD in Signals);
end;

procedure TBComPort.CallError;
var
  Errs: TComErrors;
  Errors: DWORD;
  ComStat: TComStat;
begin
  if not ClearCommError(FHandle, Errors, @ComStat) then
    raise EComPort.Create(CEMess[11]);
  Errs := [];
  if (CE_FRAME and Errors) <> 0 then Errs := Errs + [ceFrame];
  if ((CE_RXPARITY and Errors) <> 0) and (FParity <> paNone) then
    Errs := Errs + [ceRxParity];
  if (CE_OVERRUN and Errors) <> 0 then Errs := Errs + [ceOverrun];
  if (CE_RXOVER and Errors) <> 0 then Errs := Errs + [ceRxOver];
  if (CE_TXFULL and Errors) <> 0 then Errs := Errs + [ceTxFull];
  if (CE_BREAK and Errors) <> 0 then Errs := Errs + [ceBreak];
  if (CE_IOE and Errors) <> 0 then Errs := Errs + [ceIO];
  if (CE_MODE and Errors) <> 0 then Errs := Errs + [ceMode];
  if (Errs <> []) and Assigned(FOnError) then FOnError(Self, Errs);
end;

procedure TBComPort.CallRing;
begin
  if Assigned(FOnRing) then FOnRing(Self);
end;

procedure TBComPort.CallRx80Full;
begin
  if Assigned(FOnRx80Full) then FOnRx80Full(Self);
end;

procedure TBComPort.CallRxChar;
var
  Count: Integer;
begin
  Count := InBufCount;
  if (Count > 0) and Assigned(FOnRxChar) then FOnRxChar(Self, Count);
end;

procedure TBComPort.CallTxEmpty;
begin
  if Assigned(FOnTxEmpty) then FOnTxEmpty(Self);
end;

procedure TBComPort.SetBaudRate(const Value: TBaudRate);
begin
  if Value <> FBaudRate then
  begin
    FBaudRate := Value;
    ApplyDCB;
  end;
end;

procedure TBComPort.SetByteSize(const Value: TByteSize);
begin
  if Value <> FByteSize then
  begin
    FByteSize := Value;
    ApplyDCB;
  end;
end;

procedure TBComPort.SetParity(const Value: TParity);
begin
  if Value <> FParity then
  begin
    FParity := Value;
    ApplyDCB;
  end;
end;

procedure TBComPort.SetPort(const Value: ansistring);
begin
  if FConnected then
    raise EComPort.Create(CEMess[14])
  else
    if Value <> FPort then FPort := Value;
end;

procedure TBComPort.SetStopBits(const Value: TStopBits);
begin
  if Value <> FStopBits then
  begin
    FStopBits := Value;
    ApplyDCB;
  end;
end;

procedure TBComPort.SetSyncMethod(const Value: TSyncMethod);
begin
  if Value <> FSyncMethod then
  begin
    if FConnected then
      raise EComPort.Create(CEMess[14])
    else
      FSyncMethod := Value;
  end;
end;

procedure TBComPort.SetCTPriority(const Value: TThreadPriority);
begin
  if Value <> FCTPriority then
  begin
    if FConnected then
      raise EComPort.Create(CEMess[14])
    else
      FCTPriority := Value;
  end;
end;

procedure TBComPort.SetInBufSize(const Value: Integer);
begin
  if Value <> FInBufSize then
  begin
    FInBufSize := Value;
    if (FInBufSize mod 2) = 1 then Dec(FInBufSize);
    ApplyBuffer;
  end;
end;

procedure TBComPort.SetOutBufSize(const Value: Integer);
begin
  if Value <> FOutBufSize then
  begin
    FOutBufSize := Value;
    if (FOutBufSize mod 2) = 1 then Dec(FOutBufSize);
    ApplyBuffer;
  end;
end;

procedure TBComPort.SetTimeouts(const Value: TComTimeouts);
begin
  FTimeouts.Assign(Value);
  ApplyTimeouts;
end;

procedure InitAsync(var AsyncPtr: PAsync);
begin
  New(AsyncPtr);
  with AsyncPtr^ do
  begin
    FillChar(Overlapped, SizeOf(TOverlapped), 0);
    Overlapped.hEvent := CreateEvent(nil, True, True, nil);
    Data := nil;
    Size := 0;
  end;
end;

procedure DoneAsync(var AsyncPtr: PAsync);
begin
  with AsyncPtr^ do
  begin
    CloseHandle(Overlapped.hEvent);
    if Data <> nil then FreeMem(Data);
  end;
  Dispose(AsyncPtr);
  AsyncPtr := nil;
end;

procedure EnumComPorts(Ports: TStrings);
var
  KeyHandle: HKEY;
  ErrCode, Index: Integer;
  ValueName : widestring;
  Data: ansistring;
  ValueLen, DataLen, ValueType: DWORD;
  TmpPorts: TStringList;
begin
  ErrCode := RegOpenKeyEx(HKEY_LOCAL_MACHINE, 'HARDWARE\DEVICEMAP\SERIALCOMM',
                          0, KEY_READ, KeyHandle);
  if ErrCode <> ERROR_SUCCESS then
    raise EComPort.Create(CEMess[15]);
  TmpPorts := TStringList.Create;
  try
    Index := 0;
    repeat
      ValueLen := 256;
      DataLen := 256;
      SetLength(ValueName, ValueLen);
      SetLength(Data, DataLen);
      ErrCode := RegEnumValue(KeyHandle, Index, PWideChar(ValueName),
                              {$IFDEF VER120}
                              Cardinal(ValueLen),
                              {$ELSE}
                              ValueLen,
                              {$ENDIF}
                              nil, @ValueType, PByte(PAnsiChar(Data)), @DataLen);
      if ErrCode = ERROR_SUCCESS then
      begin
        SetLength(Data, DataLen);
        TmpPorts.Add(Data);
        Inc(Index);
      end
      else
        if ErrCode <> ERROR_NO_MORE_ITEMS then
          raise EComPort.Create(CEMess[15]);
    until (ErrCode <> ERROR_SUCCESS) ;
    TmpPorts.Sort;
    Ports.Assign(TmpPorts);
  finally
    RegCloseKey(KeyHandle);
    TmpPorts.Free;
  end;
end;

procedure Register;
begin
  RegisterComponents('Samples', [TBComPort]);
end;

end.
