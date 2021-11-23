
unit main;

interface

uses
  Windows,
  Messages,
  SysUtils,
  Classes,
  Graphics,
  Controls,
  SvcMgr,
  Dialogs,
  USBDeviceTree,
  ExtCtrls;

type
  TEjectDeviceLevel = (edNone = 0, edApprovedDevicesByGroup = 3);

type

  TUsbNinjaSvc = class(TService)
    CheckDevices: TTimer;
    TimerClearTNetLogLists: TTimer;
    TimerClearEventLogLists: TTimer;

    procedure ServiceExecute(Sender: TService);
    procedure ServiceStart(Sender: TService; var Started: boolean);
    procedure ServiceStop(Sender: TService; var Stopped: boolean);
    function IsExplicitlyApprovedDevice(sVid: string; sPid: string): boolean;
    function IsExplicitlyDeniedDevice(sVid: string; sPid: string): boolean;
    procedure CheckDevicesTimer(Sender: TObject);
    procedure OnLaunchCheckDevices();
    procedure StoreAuditMode_FloppyAndCdBurning_Status();

    procedure LogAuditToEventLog(EventType: DWord; Category, ID: integer;
      sMessage: string = ''; sUserName: string = ''; sUserSid: string = '';
      sDescr: string = ''; sDriveLetter: string = ''; sSerialNum: string = '';
      sVid: string = ''; sPid: string = ''; sIsMassStorage: string = '';
      sLocationInfo: string = '');

    function GetUsbLocationFromRegistry(var sBackupSerialNumber: string;
      var sLocationInformation: string; sWindowsHardwareID: string): boolean;
    function GetInteractiveUserString(): string;
    function GetEjectDeviceLevel(): TEjectDeviceLevel;
    procedure USBWorker(Dev: TUSBDevice; sUserName: string; sUserSid: string);
    procedure GetUsersDomainGroup_Sids(sUser: string;
      bForceQuery: boolean = False);
    procedure GetApprovedUsbDevicesForUser(sUserName: string;
      EjectDeviceLevel: TEjectDeviceLevel; bForceQuery: boolean = False);
    procedure GetDeniedUsbDevices(bForceQuery: boolean = False);
    function GetSecondsSincelastQuery(lastquerytime: TDateTime): int64;
    function IsTheSameUserLoggedOnAsBefore(): boolean;
    procedure TimerClearTNetLogListsTimer(Sender: TObject);
    function EnableTCPDebug: boolean;
    procedure TimerClearEventLogListsTimer(Sender: TObject);
    procedure USBDeviceTree1DeviceChange(Sender: TObject);
    function IsThisASID(sPossibleSid: string): boolean;
    function DeviceApprovedBeforeLogon(sVid, sPid: string): boolean;
    function VerboseLoggingEnabled(): boolean;
    function Check_AndOr_Set_DisableSelectiveSuspendKey(): boolean;
    procedure RegReadDelimitedUsbDevices(const RootKey: HKEY;
      const Key, Name: string; const List: TStrings);
    function GetApplicationRegistryLocation(var regPath: string;
      sKeyName: string): boolean;
  private
    UsersDomainGroup_Sids: TStringList;
    ApprovedUsbVidAndPid_ByGroup: TStringList;
    DeniedUsbVidAndPidForUserOrComputer: TStringList;
    RecentlyLoggedDevice: TStringList;
    AlreadyLogged_WindowsEventLog_ApprovedDevices: TStringList;
    procedure Logamessage(text: string; level: word);
    procedure SaveLogFile;
  public
    function GetServiceController: TServiceController; override;
    
  published
    destructor Destroy; override;
  end;

const
  sGpoRegistryLocation: string = 'SOFTWARE\Policies\UsbNinja';
  sLocalRegistryLocation: string = 'SOFTWARE\UsbNinja';
  iRefreshQuerySeconds: integer = 600; 

var
  UsbNinjaSvc: TUsbNinjaSvc;

implementation

uses
  StrUtils,
  Cfg,
  CfgMgr32,
  USB100,
  USBDesc,
  JclRegistry,
  uEjectionThread,
  SyncObjs,
  System.RegularExpressionsCore,
  uDriveHelpers,
  uHelpers,
  JvLogFile,
  
  JvLogClasses,
  JclSecurity,
  JclSysInfo,
  DateUtils,
  PsApi,
  SetupApi;

var
  dtTimeDomainGroupMembershipLastQueried: TDateTime;
  dtTimeApprovedUsbVidAndPidFromRegistryLastQueried: TDateTime;
  dtTimeExcludedUsbVidAndPidFromRegistryLastQueried: TDateTime;
  edLevel: TEjectDeviceLevel = edApprovedDevicesByGroup; 
  sPreviousUserQueried: string;
  USBDeviceTree1: TUSBDeviceTree;
  JvLogFile1: TJvLogFile;
  LogFilePath: string;
  
  bAuditOnlyModeEnabled: boolean = False;

{$R *.dfm}

destructor TUsbNinjaSvc.Destroy;
begin
  if assigned(UsersDomainGroup_Sids) then
    FreeAndNil(UsersDomainGroup_Sids);
  if assigned(ApprovedUsbVidAndPid_ByGroup) then
    FreeAndNil(ApprovedUsbVidAndPid_ByGroup);
  if assigned(DeniedUsbVidAndPidForUserOrComputer) then
    FreeAndNil(DeniedUsbVidAndPidForUserOrComputer);
  if assigned(RecentlyLoggedDevice) then
    FreeAndNil(RecentlyLoggedDevice);
  if assigned(AlreadyLogged_WindowsEventLog_ApprovedDevices) then
    FreeAndNil(AlreadyLogged_WindowsEventLog_ApprovedDevices);

  inherited; 
end;

procedure TUsbNinjaSvc.Logamessage(text: string; level: word);
var
  sev: TJvLogEventSeverity;
begin

  case level of
    1:
      sev := TJvLogEventSeverity.lesError;
    2:
      sev := TJvLogEventSeverity.lesWarning;
    3:
      sev := TJvLogEventSeverity.lesInformation;
  end;

  if VerboseLoggingEnabled then
  begin
    OutputDebugString(PChar(IntToStr(level) + ' => ' + text));
    JvLogFile1.Add(
      IntToStr(level) + ' => ' + text,
      sev);
  end;
end;

function TUsbNinjaSvc.GetApplicationRegistryLocation(var regPath: string;
  sKeyName: string): boolean;
var
  bSuccess: boolean;
begin
  bSuccess := False;

  if not JclRegistry.RegValueExists(HKLM, sGpoRegistryLocation, sKeyName) then
  begin
    
    if not JclRegistry.RegValueExists(HKLM, sLocalRegistryLocation, sKeyName)
    then
    begin
      
    end
    else
    begin
      regPath := sLocalRegistryLocation;
      bSuccess := True;
    end;
  end
  else
  begin
    regPath := sGpoRegistryLocation;
    bSuccess := True;
  end;
  
  Result := bSuccess;
end;

function TUsbNinjaSvc.GetSecondsSincelastQuery(lastquerytime
  : TDateTime): int64;
var
  CurTime: TDateTime;
  seconds: int64;
begin
  CurTime := SysUtils.Now();
  seconds := Round((CurTime - lastquerytime) * 24.0 * 60.0 * 60.0);
  
  Result := seconds;
end;

function TUsbNinjaSvc.IsThisASID(sPossibleSid: string): boolean;
var
  bResult: boolean;
  Regex: TPerlRegex;
begin

  Regex := TPerlRegex.Create();
  
  Regex.Regex := '\b[A-Z]-[0-9]{1,2}-[0-5]-[0-9].*?';
  Regex.Options := [preSingleLine, preCaseless, preMultiLine];

  try
    begin
      Regex.Subject := sPossibleSid;
      bResult := Regex.Match;
    end;
  finally
    if assigned(Regex) then
      FreeAndNil(Regex);
  end;

  Result := bResult;
end;

function TUsbNinjaSvc.GetInteractiveUserString(): string;
var
  bIsUserLoggedIn: boolean;
  sResult: string;
begin
  bIsUserLoggedIn := UserIsLoggedOn();
  
  if not bIsUserLoggedIn then
  begin
    sResult := ''; 
  end
  else
  begin
    sResult := JclSecurity.GetInteractiveUserName;
    
  end;

  Result := AnsiUpperCase(sResult);
end;

function TUsbNinjaSvc.IsTheSameUserLoggedOnAsBefore(): boolean;
var
  bResult: boolean;
  sInteractiveUser: string;
begin
  Logamessage(
    'checking same user',
    3);
  bResult := False; 

  if (not UserIsLoggedOn()) then
  begin
    bResult := False; 
    
    Result := bResult;
    Exit;
  end;

  sInteractiveUser := AnsiUpperCase(GetInteractiveUserString());
  
  sPreviousUserQueried := AnsiUpperCase(sInteractiveUser);

  if (Length(sPreviousUserQueried) > 0) and
    (sInteractiveUser <> sPreviousUserQueried) then
  begin
    Logamessage(
      'Not same interactive user, clearing int user:' + sInteractiveUser,
      3);
    UsersDomainGroup_Sids.Clear;
    
    ApprovedUsbVidAndPid_ByGroup.Clear;
    GetApprovedUsbDevicesForUser(
      sInteractiveUser,
      GetEjectDeviceLevel,
      True);

    bResult := False; 
  end;

  if (sInteractiveUser = sPreviousUserQueried) then
  begin
    bResult := True; 
    Logamessage(
      'Is Same interactive user ' + sPreviousUserQueried,
      3);
  end
  else if Length(sPreviousUserQueried) = 0 then
  begin
    Logamessage(
      'No previous user set yet',
      3);
    bResult := True;
  end;

  Result := bResult;
end;

procedure TUsbNinjaSvc.GetUsersDomainGroup_Sids(sUser: string;
  bForceQuery: boolean = False);
var
  seconds: int64;
  bSameUser: boolean;
begin
  if (not UserIsLoggedOn()) then
  begin
    
    Logamessage(
      'No user is logged on in get groupsids',
      3);
    Exit;
  end;

  if dtTimeDomainGroupMembershipLastQueried < 1 then
  begin
    
    dtTimeDomainGroupMembershipLastQueried :=
      dtTimeDomainGroupMembershipLastQueried - 1; 
    
  end;

  seconds := GetSecondsSincelastQuery(dtTimeDomainGroupMembershipLastQueried);
  if (Length(sPreviousUserQueried) < 1) and
    (dtTimeDomainGroupMembershipLastQueried < 1) then
  begin
    
    bForceQuery := True;
  end;

  bSameUser := IsTheSameUserLoggedOnAsBefore();

  if (seconds > iRefreshQuerySeconds) or (not bSameUser) or (bForceQuery = True)
  then
  begin
    
    if assigned(UsersDomainGroup_Sids) then
      UsersDomainGroup_Sids.Clear;

    GetGroupMembershipSids(UsersDomainGroup_Sids);
    dtTimeDomainGroupMembershipLastQueried := Now; 
    
  end;
end;

procedure ServiceController(CtrlCode: DWord); stdcall;
begin
  UsbNinjaSvc.Controller(CtrlCode);
end;

procedure TUsbNinjaSvc.CheckDevicesTimer(Sender: TObject);
begin
  try
    CheckDevices.Enabled := False;
    
    Check_AndOr_Set_DisableSelectiveSuspendKey();
    StoreAuditMode_FloppyAndCdBurning_Status();
    USBDeviceTree1DeviceChange(Self);
    SaveLogFile();
  finally
    CheckDevices.Enabled := True;
  end;
end;

function TUsbNinjaSvc.GetServiceController: TServiceController;
begin
  Result := ServiceController;
end;

function TUsbNinjaSvc.IsExplicitlyDeniedDevice(sVid: string;
  sPid: string): boolean;
var
  sVidAndPid, sEjectIfNLO: string;
  sCurVid, sCurPid: string;
  bIsDeniedDevice: boolean;
  sCurDevice: string;
begin
  bIsDeniedDevice := False; 
  
  for sCurDevice in DeniedUsbVidAndPidForUserOrComputer do
  begin
    SDUSplitString(
      sCurDevice,
      sVidAndPid,
      sEjectIfNLO,
      '|');
    
    SDUSplitString(
      sVidAndPid,
      sCurVid,
      sCurPid,
      '=');

    sCurVid := Trim(sCurVid);
    sCurPid := Trim(sCurPid);
    sEjectIfNLO := Trim(sEjectIfNLO);
    Logamessage(
      'Checking denied devices vid' + sCurVid + ' pid ' + sCurPid,
      3);
    if AnsiUpperCase(sCurVid) = AnsiUpperCase(sVid) then 
    begin
      if AnsiUpperCase(sCurPid) = 'ALL' then
      begin
        bIsDeniedDevice := True; 
        break;
      end
      else if AnsiUpperCase(sCurPid) = AnsiUpperCase(sPid) then
      begin
        bIsDeniedDevice := True; 
        break;
      end;
    end;
  end;

  Result := bIsDeniedDevice;
end;

function TUsbNinjaSvc.GetEjectDeviceLevel(): TEjectDeviceLevel;
var
  regPath: string;
  dEjectDeviceLevel: DWord;
begin
  Result := edNone;

  if not JclRegistry.RegValueExists(HKLM, sLocalRegistryLocation,
    'EjectDeviceLevel') then
  begin
    JclRegistry.RegWriteDWORD(
      HKLM,
      sLocalRegistryLocation,
      'EjectDeviceLevel',
      3);
    
  end;
  
  if not JclRegistry.RegValueExists(HKLM, sGpoRegistryLocation,
    'EjectDeviceLevel') then
  begin
    if not JclRegistry.RegValueExists(HKLM, sLocalRegistryLocation,
      'EjectDeviceLevel') then
    begin
      Exit;
    end
    else
    begin
      regPath := sLocalRegistryLocation;
    end;
  end
  else
  begin
    regPath := sGpoRegistryLocation;
  end;

  dEjectDeviceLevel := JclRegistry.RegReadDWORD(
    HKLM,
    regPath,
    'EjectDeviceLevel');

  case dEjectDeviceLevel of
    0:
      Result := edNone;
    
    3:
      Result := edApprovedDevicesByGroup;
  end;
end;

function TUsbNinjaSvc.VerboseLoggingEnabled(): boolean;
var
  dVerboseLog: DWord;
  bResult: boolean;
  regPath: string;
begin
  bResult := False; 

  if not JclRegistry.RegValueExists(HKLM, sLocalRegistryLocation,
    'ExtremelyVerboseLogging') then
  begin
    JclRegistry.RegWriteDWORD(
      HKLM,
      sLocalRegistryLocation,
      'ExtremelyVerboseLogging',
      0);
    
  end;

  if (GetApplicationRegistryLocation(regPath, 'ExtremelyVerboseLogging')) then
  
  begin
    
    if JclRegistry.RegValueExists(HKLM, regPath, 'ExtremelyVerboseLogging') then
    begin
      dVerboseLog := JclRegistry.RegReadDWORD(
        HKLM,
        regPath,
        'ExtremelyVerboseLogging');
      
      if dVerboseLog = 1 then
      begin
        bResult := True;
      end
      else
      begin
        bResult := False;
      end;
    end;
  end
  else
  begin
    bResult := False;
  end;

  Result := bResult;
end;

function TUsbNinjaSvc.Check_AndOr_Set_DisableSelectiveSuspendKey(): boolean;
var
  sRegPathToSystemUsbSuspendKey: string;
  regPath: string;
  bSetSelectiveSuspendKey: boolean;
  dSetKey: DWord;
begin
  
  sRegPathToSystemUsbSuspendKey := 'System\CurrentControlSet\Services\USB';

  if not JclRegistry.RegValueExists(HKLM, sLocalRegistryLocation,
    'SetDisableSelectiveSuspendKey') then
  begin
    JclRegistry.RegWriteDWORD(
      HKLM,
      sLocalRegistryLocation,
      'SetDisableSelectiveSuspendKey',
      1); 
  end;

  if (GetApplicationRegistryLocation(regPath, 'SetDisableSelectiveSuspendKey'))
  then
  
  begin
    bSetSelectiveSuspendKey := True;

    try
      begin
        if JclRegistry.RegValueExists(HKLM, regPath,
          'SetDisableSelectiveSuspendKey') then
        begin
          dSetKey := JclRegistry.RegReadDWORD(
            HKLM,
            regPath,
            'SetDisableSelectiveSuspendKey');
          if dSetKey = 0 then
          begin
            bSetSelectiveSuspendKey := False; 
            
            JclRegistry.RegWriteDWORD(
              HKLM,
              sRegPathToSystemUsbSuspendKey,
              'DisableSelectiveSuspend',
              0);
          end
          else if dSetKey = 1 then
          begin
            bSetSelectiveSuspendKey := True;
            JclRegistry.RegWriteDWORD(
              HKLM,
              sRegPathToSystemUsbSuspendKey,
              'DisableSelectiveSuspend',
              1);
          end;
        end;
      end;
    except
      
    end;
  end
  else
  begin
    bSetSelectiveSuspendKey := False;
  end;

  Result := bSetSelectiveSuspendKey;

end;

procedure TUsbNinjaSvc.StoreAuditMode_FloppyAndCdBurning_Status();
var
  regPath: string;
  dAuditMode: DWord;
  dDisableFloppy: DWord;
  dDisableCdWriting: DWord;
begin
  
  if ((not JclRegistry.RegKeyExists(HKLM, sGpoRegistryLocation)) and
    (not JclRegistry.RegKeyExists(HKLM, sLocalRegistryLocation))) then
    Exit;

  try
    begin

      if not JclRegistry.RegValueExists(HKLM, sLocalRegistryLocation,
        'AuditModeEnabled') then
      begin
        JclRegistry.RegWriteDWORD(
          HKLM,
          sLocalRegistryLocation,
          'AuditModeEnabled',
          1);
        
      end;

      if (GetApplicationRegistryLocation(regPath, 'AuditModeEnabled')) then
      begin
        if JclRegistry.RegValueExists(HKLM, regPath, 'AuditModeEnabled') then
        begin
          dAuditMode := JclRegistry.RegReadDWORD(
            HKLM,
            regPath,
            'AuditModeEnabled');
          if dAuditMode = 0 then
          begin
            bAuditOnlyModeEnabled := False; 
          end
          else if dAuditMode = 1 then
          begin
            bAuditOnlyModeEnabled := True;
          end;
        end;
      end;

      if not JclRegistry.RegValueExists(HKLM, sLocalRegistryLocation,
        'DisableFloppyDrive') then
      begin
        JclRegistry.RegWriteDWORD(
          HKLM,
          sLocalRegistryLocation,
          'DisableFloppyDrive',
          2);
        
      end;

      if (GetApplicationRegistryLocation(regPath, 'DisableFloppyDrive')) then
      begin
        if JclRegistry.RegValueExists(HKLM, regPath, 'DisableFloppyDrive') then
        begin
          dDisableFloppy := JclRegistry.RegReadDWORD(
            HKLM,
            regPath,
            'DisableFloppyDrive');
          if dDisableFloppy = 0 then
          begin
            
            ModifyFloppy(False);
          end
          else if dDisableFloppy = 1 then
          begin
            
            ModifyFloppy(True);
          end;
        end;
      end;

      if not JclRegistry.RegValueExists(HKLM, sLocalRegistryLocation,
        'DisableWindowsCdWriting') then
      begin
        JclRegistry.RegWriteDWORD(
          HKLM,
          sLocalRegistryLocation,
          'DisableWindowsCdWriting',
          2); 
      end;

      if (GetApplicationRegistryLocation(regPath, 'DisableWindowsCdWriting'))
      then
      begin
        if JclRegistry.RegValueExists(HKLM, regPath, 'DisableWindowsCdWriting')
        then
        begin
          dDisableCdWriting := JclRegistry.RegReadDWORD(
            HKLM,
            regPath,
            'DisableWindowsCdWriting');
          if dDisableCdWriting = 0 then
          begin
            
            ModifyWindowsCdWriting(False);
          end
          else if dDisableCdWriting = 1 then
          begin
            
            ModifyWindowsCdWriting(True);
          end;
        end;
      end;
    end;
  except
    
  end;
end;

procedure TUsbNinjaSvc.TimerClearEventLogListsTimer(Sender: TObject);
begin
  
  if (AlreadyLogged_WindowsEventLog_ApprovedDevices.Count > 0) then
  begin
    AlreadyLogged_WindowsEventLog_ApprovedDevices.Clear;
  end;
  
end;

procedure TUsbNinjaSvc.SaveLogFile();
begin
  if assigned(JvLogFile1) then
    try
      JvLogFile1.SaveToFile(LogFilePath);
      Logamessage(
        'UsbNinja service is shutting down' + DateTimeToStr(Now),
        2);
    except
      on E: Exception do
      begin
        OutputDebugString(PChar(E.Message));
      end;
    end;
end;

procedure TUsbNinjaSvc.TimerClearTNetLogListsTimer(Sender: TObject);
begin
  if RecentlyLoggedDevice.Count > 0 then
  begin
    RecentlyLoggedDevice.Clear; 
  end;

  SaveLogFile();
end;

procedure TUsbNinjaSvc.OnLaunchCheckDevices();
begin
  
  StoreAuditMode_FloppyAndCdBurning_Status();
  USBDeviceTree1DeviceChange(Self);
end;

function TUsbNinjaSvc.IsExplicitlyApprovedDevice(sVid: string;
  sPid: string): boolean;
var
  bIsApproved: boolean;
  sCurDevice: string;
  sVidAndPid, sEjectIfNLO: string;
  sCurVid, sCurPid: string;
begin
  
  sEjectIfNLO := '';
  bIsApproved := False;

  for sCurDevice in ApprovedUsbVidAndPid_ByGroup do
  begin
    SDUSplitString(
      sCurDevice,
      sVidAndPid,
      sEjectIfNLO,
      '|');
    
    SDUSplitString(
      sVidAndPid,
      sCurVid,
      sCurPid,
      '=');
    
    if AnsiUpperCase(sCurVid) = AnsiUpperCase(sVid) then 
    begin
      if AnsiUpperCase(sCurPid) = 'ALL' then
      begin
        Logamessage(
          'All devices allowed for vendor ' + sCurVid,
          2);
        bIsApproved := True; 
        break;
      end
      else if AnsiUpperCase(sCurPid) = AnsiUpperCase(sPid) then
      begin
        Logamessage(
          'Device ' + sCurPid + ' is allowed',
          2);
        bIsApproved := True; 
        break;
      end;
    end
    else if (AnsiUpperCase(sCurVid) = 'ALL') and (AnsiUpperCase(sCurPid) = 'ALL')
    then
    begin
      Logamessage(
        'devices is allowed, in all-all group',
        2);
      bIsApproved := True; 
      break;
    end
    else
      Logamessage(
        'Vid no match sCurvid ' + sCurVid + ' svid ' + sVid + ' scurpid '
        + sCurPid,
        3);
  end;

  Result := bIsApproved;
end;

function GetDriveFromSerial(sKnownUsbSerial: string;
  var sDriveLetter: string): boolean;
var
  DriveMountPoints: TStringList;
  i: integer;
  serial, drive: string;
  bResult: boolean;
begin
  bResult := False;
  sDriveLetter := '';

  LoadSetupApi;
  LoadConfigManagerApi;
  DriveMountPoints := TStringList.Create;
  try
    
    DriveMountPoints.Sorted := True;
    
    uDriveHelpers.FillDriveList(DriveMountPoints);

    for i := 0 to DriveMountPoints.Count - 1 do
    begin
      
      drive := DriveMountPoints.ValueFromIndex[i];
      serial := DriveMountPoints.Names[i];

      if AnsiUpperCase(serial) = AnsiUpperCase(sKnownUsbSerial) then
      begin
        bResult := True;
        sDriveLetter := drive[1];
      end;
    end;

    Result := bResult;
  finally
    UnloadConfigManagerApi;
    UnloadSetupApi;
    if assigned(DriveMountPoints) then
      FreeAndNil(DriveMountPoints);
  end;
end;

procedure TUsbNinjaSvc.LogAuditToEventLog(EventType: DWord;
  Category, ID: integer; sMessage: string = ''; sUserName: string = '';
  sUserSid: string = ''; sDescr: string = ''; sDriveLetter: string = '';
  sSerialNum: string = ''; sVid: string = ''; sPid: string = '';
  sIsMassStorage: string = ''; sLocationInfo: string = '');
var
  curLine: string;
  sLoggedTodayLine: string;
  bLogTodayEvent: boolean;
  bLogEvent: boolean;
  sUniqueIdentifier: string;
  bAlreadyLoggedToTNetLog: boolean;
  bVerboseLoggingEnabled: boolean;
begin
  sUniqueIdentifier := sMessage + sUserSid + sVid + sPid + sSerialNum +
    sDriveLetter;
  bAlreadyLoggedToTNetLog := False; 
  bVerboseLoggingEnabled := VerboseLoggingEnabled();

  bLogTodayEvent := True; 
  bLogEvent := True;
  
  if RecentlyLoggedDevice.Count > 0 then
  begin
    
    for curLine in RecentlyLoggedDevice do
    begin
      if ((AnsiStartsStr(AnsiUpperCase(curLine),
        AnsiUpperCase(sUniqueIdentifier))) or
        (AnsiStartsStr(AnsiUpperCase(sUniqueIdentifier),
        AnsiUpperCase(curLine)))) then
      begin
        bLogEvent := False; 
      end;
    end;
  end;

  if AnsiContainsStr(sMessage, '*** APPROVED DEVICE ***') or
    AnsiContainsStr(sMessage, '*** AUDIT ONLY') then
  begin
    
    if ((AlreadyLogged_WindowsEventLog_ApprovedDevices.Count > 0)) then
    begin
      
      for sLoggedTodayLine in AlreadyLogged_WindowsEventLog_ApprovedDevices do
      begin
        if ((AnsiStartsStr(AnsiUpperCase(sLoggedTodayLine),
          AnsiUpperCase(sUniqueIdentifier))) or
          (AnsiStartsStr(AnsiUpperCase(sUniqueIdentifier),
          AnsiUpperCase(sLoggedTodayLine)))) then
        begin
          bLogTodayEvent := False; 
        end;
      end;
    end;

    if bLogTodayEvent = True then
    begin
      
      AlreadyLogged_WindowsEventLog_ApprovedDevices.Add(sUniqueIdentifier);
      
    end;
    
    RecentlyLoggedDevice.Add(sUniqueIdentifier); 
    
  end
  else
  begin
    if bLogEvent = True then
    begin
      
      RecentlyLoggedDevice.Add(sUniqueIdentifier); 
      
      AlreadyLogged_WindowsEventLog_ApprovedDevices.Add(sUniqueIdentifier);
      
    end;
  end;

  if ((bLogEvent = True) and (bVerboseLoggingEnabled)) then
  begin
    bAlreadyLoggedToTNetLog := True;

    Logamessage(
      ',' + sMessage + ',' + sUserName + ',' + sUserSid + ',' + sVid + ',' +
      sPid + ',' + sDescr + ',' + '[' + sDriveLetter + ':],' + sSerialNum + ','
      + sIsMassStorage + ',' + sLocationInfo + ',' + IntToStr(Category) + ',' +
      DateTimeToStr(Now),
      2); 
    
  end;

  if (bLogTodayEvent = True) then
  begin
    LogMessage(
      #13 + #10 
      + sMessage + #13 + #10 + 'User Name: ' + sUserName + #13 + #10 +
      'User SID: ' + sUserSid + #13 + #10 + 'Vendor ID: ' + sVid + #13 + #10 +
      'Product ID: ' + sPid + #13 + #10 + 'Description: ' + sDescr + #13 + #10 +
      'Drive: [' + sDriveLetter + ':]' + #13 + #10 + 'Serial Number: ' +
      sSerialNum + #13 + #10 + 'Is Mass Storage?: ' + sIsMassStorage + #13 + #10
      + 'Location: ' + sLocationInfo + #13 + #10 + 'Event Category: ' +
      IntToStr(Category) + #13 + #10 + 'Time of event: ' + DateTimeToStr(Now),
      EventType,
      Category,
      ID);

    Logamessage(
      ',' + sMessage + ',' + sUserName + ',' + sUserSid + ',' + sVid + ',' +
      sPid + ',' + sDescr + ',' + '[' + sDriveLetter + ':],' + sSerialNum + ','
      + sIsMassStorage + ',' + sLocationInfo + ',' + IntToStr(Category) + ',' +
      DateTimeToStr(Now),
      2);

    if ((not bAlreadyLoggedToTNetLog) and (bVerboseLoggingEnabled)) then
    begin
      Logamessage(
        ',' + sMessage + ',' + sUserName + ',' + sUserSid + ',' + sVid + ',' +
        sPid + ',' + sDescr + ',' + '[' + sDriveLetter + ':],' + sSerialNum +
        ',' + sIsMassStorage + ',' + sLocationInfo + ',' + IntToStr(Category) +
        ',' + DateTimeToStr(Now),
        2);
    end;
    
  end;
end;

function TUsbNinjaSvc.DeviceApprovedBeforeLogon(sVid, sPid: string): boolean;
var
  bApprovedBeforeLogon, bFoundApprovedDeviceMatch: boolean;
  sCurDevice: string;
  sVidAndPid, sEjectIfNLO: string;
  sCurVid, sCurPid: string;
begin
  
  sEjectIfNLO := '';
  bApprovedBeforeLogon := True;
  
  bFoundApprovedDeviceMatch := False;

  for sCurDevice in ApprovedUsbVidAndPid_ByGroup do
  begin
    SDUSplitString(
      sCurDevice,
      sVidAndPid,
      sEjectIfNLO,
      '|');
    
    SDUSplitString(
      sVidAndPid,
      sCurVid,
      sCurPid,
      '=');

    sCurVid := Trim(sCurVid);
    sCurPid := Trim(sCurPid);
    sEjectIfNLO := Trim(sEjectIfNLO);

    if AnsiUpperCase(sCurVid) = AnsiUpperCase(sVid) then 
    begin
      if AnsiUpperCase(sVid) = 'ALL' then
      begin
        bFoundApprovedDeviceMatch := True;
        
        break;
      end
      else if AnsiUpperCase(sCurPid) = AnsiUpperCase(sPid) then
      begin
        bFoundApprovedDeviceMatch := True; 
        break;
      end;
    end;
  end;

  if bFoundApprovedDeviceMatch then
  begin
    
    if AnsiUpperCase(Trim(sEjectIfNLO)) = 'YES' then
    begin
      bApprovedBeforeLogon := False;
    end
    else
    begin
      bApprovedBeforeLogon := True;
    end;
  end;

  Result := bApprovedBeforeLogon;
end;

function TUsbNinjaSvc.GetUsbLocationFromRegistry(var sBackupSerialNumber
  : string; var sLocationInformation: string;
  sWindowsHardwareID: string): boolean;
const
  reg_usb: string = 'SYSTEM\CurrentControlSet\Enum\USB';
var
  UsbRegistry: TStringList;
  curKey: string;
  bResult: boolean;
begin
  bResult := False;

  try
    begin
      if JclRegistry.RegKeyExists(HKEY_LOCAL_MACHINE, reg_usb) then
      begin
        bResult := True;
        UsbRegistry := TStringList.Create;
        try
          JclRegistry.RegGetKeyNames(
            HKEY_LOCAL_MACHINE,
            reg_usb + '\' + sWindowsHardwareID,
            UsbRegistry);
          for curKey in UsbRegistry do
          begin
            sBackupSerialNumber := curKey;
            if JclRegistry.RegKeyExists(HKEY_LOCAL_MACHINE,
              reg_usb + '\' + sWindowsHardwareID + '\' + curKey) then
            begin
              
              if (JclRegistry.RegValueExists(HKEY_LOCAL_MACHINE,
                reg_usb + '\' + sWindowsHardwareID + '\' + curKey,
                'LocationInformation')) then
              begin
                sLocationInformation := JclRegistry.RegReadString(
                  HKEY_LOCAL_MACHINE,
                  reg_usb + '\' + sWindowsHardwareID + '\' + curKey,
                  'LocationInformation');

                break;
              end;
              
            end;
          end;
        finally
          if assigned(UsbRegistry) then
            FreeAndNil(UsbRegistry);
        end;
      end;
    end;
  except
    
    LogMessage(
      #13 + #10 + 'Error querying registry at: ' + reg_usb,
      EVENTLOG_ERROR_TYPE,
      101,
      1198);
  end;

  Result := bResult;
end;

procedure TUsbNinjaSvc.USBWorker(Dev: TUSBDevice; sUserName: string;
  sUserSid: string);
var
  iDevDescripterCount: integer;
  sWindowsHardwareID: string;
  sVid, sPid: string;
  bExplicitlyApprovedDevice: boolean;
  bExplicitlyDeniedDevice, bEjectDevice: boolean;
  IsMassStorage: boolean;
  sLocationInformation, sSerialNumber, sDescription: string;
  sDriveLetter: string;
  sBackupSerialNumber: string;
  
  iDriveSize: int64;
  sWinFolder, sWinDriveLetter: string;
  bDeviceAllowedIfNLO: boolean;
  curDevice: string;
  bEjectionInProgress: boolean;
  
begin
  bEjectionInProgress := False;
  bEjectDevice := True; 
  IsMassStorage := False; 
  sDriveLetter := '';

  if Dev.DeviceInstance <> 0 then
  begin
    for iDevDescripterCount := 0 to Dev.DescriptorCount - 1 do
    begin
      if Terminated then
      begin
        break;
      end;

      if Dev.Descriptors[iDevDescripterCount].DescriptorType = USB_INTERFACE_DESCRIPTOR_TYPE
      then
      begin
        IsMassStorage := Dev.Descriptors[iDevDescripterCount]
          .InterfaceDescr.bInterfaceClass = USB_DEVICE_CLASS_STORAGE;
        
      end;

      if IsMassStorage then
      begin
        break;
      end;
    end;
  end;

  sLocationInformation := ''; 
  bExplicitlyApprovedDevice := False; 
  bExplicitlyDeniedDevice := False; 

  if IsMassStorage = True then
  begin
    sWindowsHardwareID := Format(
      'Vid_%.4x',
      [Dev.ConnectionInfo.DeviceDescriptor.idVendor]);
    sWindowsHardwareID := sWindowsHardwareID + '&' +
      Format('Pid_%.4x', [Dev.ConnectionInfo.DeviceDescriptor.idProduct]);

    sVid := AnsiUpperCase(Format('%.4x',
      [Dev.ConnectionInfo.DeviceDescriptor.idVendor]));
    sPid := AnsiUpperCase(Format('%.4x',
      [Dev.ConnectionInfo.DeviceDescriptor.idProduct]));

    GetUsbLocationFromRegistry(
      sBackupSerialNumber,
      sLocationInformation,
      sWindowsHardwareID);
    Logamessage(
      'Mass storage device vid: ' + sVid + ' pid ' + sPid,
      3);
  end;

  if Length(Trim(sVid)) <= 0 then
  begin
    
    Exit;
  end;

  bExplicitlyApprovedDevice := IsExplicitlyApprovedDevice(
    sVid,
    sPid);
  bExplicitlyDeniedDevice := IsExplicitlyDeniedDevice(
    sVid,
    sPid);

  if (IsMassStorage = False) and (bExplicitlyDeniedDevice = False) then
  begin
    Exit;
  end;

  if ((bExplicitlyApprovedDevice) and (not bExplicitlyDeniedDevice)) then
  begin
    
    bEjectDevice := False;
  end;

  if ((not bExplicitlyApprovedDevice) or (bExplicitlyDeniedDevice)) then
  begin
    bEjectDevice := True;
  end;

  if Dev.SerialNumber <> '' then
  begin
    sSerialNumber := Format(
      '%s',
      [Dev.SerialNumber]);
  end
  else
  begin
    sSerialNumber := sBackupSerialNumber;
  end;

  if Dev.DeviceDescription <> '' then
  begin
    sDescription := Dev.DeviceDescription;
  end
  else
  begin
    sDescription := '';
  end;
  
  GetDriveFromSerial(
    sSerialNumber,
    sDriveLetter);

  sWinFolder := JclSysInfo.GetWindowsFolder;
  if Length(sWinFolder) > 0 then
  begin
    sWinDriveLetter := AnsiReplaceStr(
      ExtractFileDrive(sWinFolder),
      ':',
      '');
    
    if AnsiUpperCase(sWinDriveLetter) = AnsiUpperCase(sDriveLetter) then
    begin
      
      bEjectDevice := False;
      LogAuditToEventLog(
        EVENTLOG_WARNING_TYPE,
        102,
        1198,
        '*** WINDOWS VOLUME DETECTED AS USB *** - Device was NOT ejected',
        sUserName,
        sUserSid,
        sDescription,
        sDriveLetter,
        sSerialNumber,
        sVid,
        sPid,
        '',
        sLocationInformation);
    end;
  end;

  bDeviceAllowedIfNLO := DeviceApprovedBeforeLogon(
    sVid,
    sPid);
  if ((UserIsLoggedOn() = False) and (bDeviceAllowedIfNLO = False)) then
  begin
    
    bEjectDevice := True;

  end;

  if bEjectDevice = False then 
  begin

    LogAuditToEventLog(
      EVENTLOG_INFORMATION_TYPE,
      103,
      1198,
      '*** APPROVED DEVICE ***',
      sUserName,
      sUserSid,
      sDescription,
      sDriveLetter,
      sSerialNumber,
      sVid,
      sPid,
      '',
      sLocationInformation);
  end
  else if (bEjectDevice = True) then
  begin
    
    if Dev.SerialNumber <> '' then
    begin
      sSerialNumber := Format(
        '%s',
        [Dev.SerialNumber]);
    end
    else
    begin
      sSerialNumber := sBackupSerialNumber;
    end;

    if Dev.DeviceDescription <> '' then
    begin
      sDescription := Dev.DeviceDescription + ' - ' + Dev.Product;
    end
    else
    begin
      sDescription := '';
    end;
    
    if Length(sDriveLetter) <= 0 then
    begin
      
      sleep(500);
      GetDriveFromSerial(
        sSerialNumber,
        sDriveLetter);
    end;
    
    if (Length(sDriveLetter) > 0) then
    begin
      try
        iDriveSize := GetDriveTotalSize(sDriveLetter[1]);
        if (iDriveSize < 1) then
        begin
          
          bEjectDevice := False;
        end;
      except
        bEjectDevice := False;
      end;
    end;

    Crit_DeviceList.Enter;
    sleep(250); 
    for curDevice in DeviceList do
    begin
      if (AnsiUpperCase(sVid + sPid + sUserSid) = AnsiUpperCase(curDevice)) then
      begin
        bEjectionInProgress := True;
        
      end;
    end;
    Crit_DeviceList.Leave;

    if (((Length(sDriveLetter) > 0) and (bEjectDevice = True)) or
      ((IsMassStorage = True) and (bEjectDevice = True))) then
    begin
      begin

        if (bAuditOnlyModeEnabled = False) then
        begin
          if (not bEjectionInProgress) then
          begin
            
            with TEjectDeviceThread.Create(True) do
            begin
              MyUsb := Dev;
              UserName := sUserName;
              UserSid := sUserSid;
              SerialNumber := sSerialNumber;
              Vid := sVid;
              Pid := sPid;
              LocationInformation := sLocationInformation;
              Description := sDescription;
              DriveLetter := sDriveLetter;

              FreeOnTerminate := True;
              Resume;
            end;
          end;
        end
        else
        begin
          
          LogAuditToEventLog(
            EVENTLOG_INFORMATION_TYPE,
            103,
            1198,
            '*** AUDIT ONLY *** - This would have been ejected:',
            sUserName,
            sUserSid,
            sDescription,
            sDriveLetter,
            sSerialNumber,
            sVid,
            sPid,
            '',
            sLocationInformation);
        end;
      end;
    end;
  end;
end;

procedure TUsbNinjaSvc.USBDeviceTree1DeviceChange(Sender: TObject);
var
  iUSBDeviceTreeCount: integer;
  HC: TUSBHostController;
  sUserName, sUserSid: string;
  
  procedure AddNodes(Dev: TUSBDevice);
  var
    iDevIsHubCount: integer;
  begin
    if Terminated then
    begin
      Exit;
    end;

    if Dev = nil then
    begin
      Exit;
    end;

    if Dev.IsHub then
    begin
      for iDevIsHubCount := 0 to Dev.Count - 1 do
      begin
        AddNodes(Dev.Devices[iDevIsHubCount]);
      end;
    end
    else
    begin
      if Dev.DeviceInstance <> 0 then
      begin
        USBWorker(
          Dev,
          sUserName,
          sUserSid);
        
      end;
    end;
  end;

begin

  if Terminated then
  begin
    Exit;
  end;
  
  sUserName := GetInteractiveUserString;
  sUserSid := GetAccountSID(sUserName);

  if (Length(sUserName) < 1) or (Length(sUserSid) < 1) then
  begin
    sUserName := 'Unknown';
    sUserSid := 'Unknown';
  end;
  
  GetApprovedUsbDevicesForUser(
    sUserName,
    edLevel);
  GetDeniedUsbDevices();

  for iUSBDeviceTreeCount := 0 to USBDeviceTree1.Count - 1 do
  begin
    if USBDeviceTree1.HostControllers[iUSBDeviceTreeCount].DevicesConnected > 0
    then
    begin
      if Terminated then
      begin
        Exit;
      end;

      HC := USBDeviceTree1.HostControllers[iUSBDeviceTreeCount];
      AddNodes(HC.RootHub);
    end;
  end;
end;

procedure TUsbNinjaSvc.GetDeniedUsbDevices(bForceQuery: boolean = False);
var
  Regex: TPerlRegex;
  TempStringList: TStringList;
  sCurTemp: string;
  regPath: string;
  seconds: int64;
begin
  if (dtTimeExcludedUsbVidAndPidFromRegistryLastQueried < 1) then
  begin
    seconds := iRefreshQuerySeconds + 1;
  end
  else
  begin
    seconds := GetSecondsSincelastQuery
      (dtTimeExcludedUsbVidAndPidFromRegistryLastQueried);
  end;

  if ((not JclRegistry.RegKeyExists(HKLM, sGpoRegistryLocation)) and
    (not JclRegistry.RegKeyExists(HKLM, sLocalRegistryLocation))) then
    Exit;

  if (seconds > iRefreshQuerySeconds) or (bForceQuery = True) then
  begin

    if Assigned(DeniedUsbVidAndPidForUserOrComputer)then
      if DeniedUsbVidAndPidForUserOrComputer.Count > 0 then
        DeniedUsbVidAndPidForUserOrComputer.Clear;

    TempStringList := TStringList.Create;
    Regex := TPerlRegex.Create();
    try
      begin
        Regex.Regex := '^(.*)\\.*$';
        Regex.Options := [preCaseless, preMultiLine];
        if GetApplicationRegistryLocation(regPath, 'DeniedDevices') then
        begin
          if JclRegistry.RegValueExists(HKLM, regPath, 'DeniedDevices') then
          begin
            RegReadDelimitedUsbDevices(
              HKLM,
              regPath,
              'DeniedDevices',
              TempStringList);
            if TempStringList.Count > 0 then
            begin
              for sCurTemp in TempStringList do
              begin
                
                DeniedUsbVidAndPidForUserOrComputer.Add
                  (AnsiUpperCase(AnsiReplaceStr(sCurTemp, '|', '=')));
              end;
            end;
          end;
        end;
      end;
    finally
      if assigned(TempStringList) then
        FreeAndNil(TempStringList);
      if assigned(Regex) then
        FreeAndNil(Regex);
    end;

    dtTimeApprovedUsbVidAndPidFromRegistryLastQueried := Now();

  end;
end;

procedure TUsbNinjaSvc.GetApprovedUsbDevicesForUser(sUserName: string;
  EjectDeviceLevel: TEjectDeviceLevel; bForceQuery: boolean = False);
var
  Regex: TPerlRegex;
  sDomainFromAccountName: string;
  TempStringList: TStringList;
  sCurTemp: string;
  regPath: string;
  
  seconds: int64;
  bSameUser: boolean;
  
  slGroupNames: TStringList; 
  slGpoOrLocalReg: TStringList;
  sCurGroupLine: string;
  sRegistryGroupSid_FromDomain: string;
  sCurDomainGroupSid: string;
  bMemberOfThisCurrentGroup: boolean;
  bThisIsASid: boolean;
begin
  
  if ((not JclRegistry.RegKeyExists(HKLM, sGpoRegistryLocation +
    '\ApprovedDevicesByGroup')) and (not JclRegistry.RegKeyExists(HKLM,
    sLocalRegistryLocation + '\ApprovedDevicesByGroup'))) then
  begin
    Logamessage(
      'No approved devices exist in the registry',
      1);
    Exit;
  end;
  Logamessage(
    'getapproved devices for user',
    3);
  
  if (dtTimeApprovedUsbVidAndPidFromRegistryLastQueried < 1) then
  begin
    seconds := iRefreshQuerySeconds + 1;
  end
  else
  begin
    seconds := GetSecondsSincelastQuery
      (dtTimeApprovedUsbVidAndPidFromRegistryLastQueried);
  end;
  
  bSameUser := IsTheSameUserLoggedOnAsBefore();

  if ((seconds > iRefreshQuerySeconds) or (bSameUser = False) or
    (bForceQuery = True) or (ApprovedUsbVidAndPid_ByGroup.Count = 0)) then
  begin

    slGpoOrLocalReg := TStringList.Create;
    try

      if (not JclRegistry.RegKeyExists(HKLM, sGpoRegistryLocation +
        '\ApprovedDevicesByGroup')) then
      begin
        if (not JclRegistry.RegKeyExists(HKLM, sLocalRegistryLocation +
          '\ApprovedDevicesByGroup')) then
        begin
          Exit; 
        end
        else
        begin
          JclRegistry.RegGetValueNames(
            HKEY_LOCAL_MACHINE,
            sLocalRegistryLocation + '\ApprovedDevicesByGroup',
            slGpoOrLocalReg);
          if slGpoOrLocalReg.Count > 0 then
          begin
            regPath := sLocalRegistryLocation;
            
          end;
        end;
      end
      else
      begin
        JclRegistry.RegGetValueNames(
          HKEY_LOCAL_MACHINE,
          sGpoRegistryLocation + '\ApprovedDevicesByGroup',
          slGpoOrLocalReg);
        if slGpoOrLocalReg.Count > 0 then
        begin
          regPath := sGpoRegistryLocation;
          
        end;
      end;

    finally
      if assigned(slGpoOrLocalReg) then
        FreeAndNil(slGpoOrLocalReg);
    end;

    TempStringList := TStringList.Create;
    Regex := TPerlRegex.Create();
    Regex.Regex := '^(.*)\\.*$';
    Regex.Options := [preCaseless, preMultiLine];
    try
      begin
        
        if (EjectDeviceLevel = edNone) then
        begin
          Exit;
        end;
        
        if (EjectDeviceLevel = edApprovedDevicesByGroup) then
        begin
          Logamessage(
            'Checking for group approval of USB devices',
            3);

          GetUsersDomainGroup_Sids(sUserName);

          if TempStringList.Count > 0 then
            TempStringList.Clear;

          regPath := regPath + '\ApprovedDevicesByGroup';
          slGroupNames := TStringList.Create;
          slGroupNames.Clear;
          try
            begin
              if JclRegistry.RegKeyExists(HKLM, regPath) then
              begin
                JclRegistry.RegGetValueNames(
                  HKEY_LOCAL_MACHINE,
                  regPath,
                  slGroupNames);
              end;

              for sCurGroupLine in slGroupNames do
              begin
                
                Regex.Subject := sUserName;
                if Regex.Match then
                begin
                  if Regex.GroupCount >= 1 then
                  begin
                    sDomainFromAccountName := Regex.Groups[1] + '\';
                  end
                  else
                  begin
                    sDomainFromAccountName := '';
                  end;
                end;

                bThisIsASid := IsThisASID(sCurGroupLine);
                if bThisIsASid then
                begin
                  sRegistryGroupSid_FromDomain := sCurGroupLine;
                end
                else
                begin
                  if IsCorrectlyFormattedGroup(sCurGroupLine) then
                  begin
                    sRegistryGroupSid_FromDomain :=
                      GetAccountSID(sCurGroupLine);
                  end
                  else
                  begin
                    LogAuditToEventLog(
                      EVENTLOG_WARNING_TYPE,
                      102,
                      1198,
                      '*** "' + sCurGroupLine +
                      '" is not a correctly formatted local or domain group.',
                      '',
                      '',
                      '',
                      '',
                      '',
                      '',
                      '',
                      '',
                      '');

                    continue; 
                  end;
                end;
                
                bMemberOfThisCurrentGroup := False;

                for sCurDomainGroupSid in UsersDomainGroup_Sids do
                begin
                  if AnsiUpperCase(sRegistryGroupSid_FromDomain)
                    = AnsiUpperCase(sCurDomainGroupSid) then
                  begin
                    
                    bMemberOfThisCurrentGroup := True;
                    break;
                  end;
                  
                end;

                if bMemberOfThisCurrentGroup = True then
                begin
                  if JclRegistry.RegValueExists(HKLM, regPath, sCurGroupLine)
                  then
                  begin
                    RegReadDelimitedUsbDevices(
                      HKLM,
                      regPath,
                      sCurGroupLine,
                      TempStringList);
                    if TempStringList.Count > 0 then
                    begin
                      for sCurTemp in TempStringList do
                      begin
                        
                        Logamessage(
                          'adding approved device ' + sCurTemp,
                          3);
                        ApprovedUsbVidAndPid_ByGroup.Add
                          (ParseVidPidFromRegistry(sCurTemp));
                        
                      end;
                    end;
                  end;
                end;
              end;
            end;
          finally
            if assigned(slGroupNames) then
              FreeAndNil(slGroupNames);
          end;
        end;
      end;
    finally
      if assigned(TempStringList) then
        FreeAndNil(TempStringList);
      if assigned(Regex) then
        FreeAndNil(Regex);
    end;
    dtTimeApprovedUsbVidAndPidFromRegistryLastQueried := Now();

  end;
end;

procedure TUsbNinjaSvc.RegReadDelimitedUsbDevices(const RootKey: HKEY;
  const Key, Name: string; const List: TStrings);
var
  sRegValue, sDevice: string;
begin
  
  Logamessage(
    'Adding reg devices to delimited list',
    3);
  sRegValue := JclRegistry.RegReadString(
    RootKey,
    Key,
    Name);

  if Length(sRegValue) < 1 then
    Exit;

  List.BeginUpdate;
  try
    begin
      List.Clear;

      if not AnsiContainsStr(sRegValue, ';') then
      begin
        
        if Length(sRegValue) > 0 then
        begin
          List.Add(sRegValue);

          Exit; 
        end;
      end;

      while AnsiContainsStr(sRegValue, '|') do
      begin
        if not AnsiContainsStr(sRegValue, ';') then
        begin
          
          if Length(sRegValue) > 0 then
          begin
            List.Add(sRegValue);
            
            Exit; 
          end;
        end
        else
        begin
          if Length(sRegValue) > 0 then
          begin
            SDUSplitString(
              sRegValue,
              sDevice,
              sRegValue,
              ';');
            
            List.Add(sDevice); 
          end;
        end;
      end;
    end;
  finally
    List.EndUpdate;
  end;
end;

procedure TUsbNinjaSvc.ServiceExecute(Sender: TService);
begin

  OnLaunchCheckDevices();

  while not Terminated do
  begin
    sleep(50);
    try
      ServiceThread.ProcessRequests(False);
    except on e:Exception do
    begin
      OutputDebugString(PChar(e.Message));
    end;
    end;

  end;
end;

function TUsbNinjaSvc.EnableTCPDebug: boolean;
var
  regPath: string;
  dEnableTCPDebug: DWord;
  bResult: boolean;
begin
  bResult := False;
  
  if not JclRegistry.RegValueExists(HKLM, sLocalRegistryLocation,
    'EnableTCPDebug') then
  begin
    JclRegistry.RegWriteDWORD(
      HKLM,
      sLocalRegistryLocation,
      'EnableTCPDebug',
      0);
  end;

  if not GetApplicationRegistryLocation(regPath, 'EnableTCPDebug') then
  begin
    Result := bResult;
    Exit;
  end;

  dEnableTCPDebug := JclRegistry.RegReadDWORD(
    HKLM,
    regPath,
    'EnableTCPDebug');
  if dEnableTCPDebug = 0 then
  begin
    bResult := False; 
  end
  else
  begin
    bResult := True;
  end;

  Result := bResult;
end;

procedure TUsbNinjaSvc.ServiceStart(Sender: TService; var Started: boolean);
var
  winver: JclSysInfo.TWindowsVersion;
  
  regPath: string;
  dDebug: DWord;
begin

  if not JclRegistry.RegKeyExists(HKLM, sLocalRegistryLocation) then
  begin
    JclRegistry.RegCreateKey(
      HKLM,
      sLocalRegistryLocation);
    
  end;

  if not JclRegistry.RegValueExists(HKLM, sLocalRegistryLocation, 'DebugDelay')
  then
  begin
    JclRegistry.RegWriteDWORD(
      HKLM,
      sLocalRegistryLocation,
      'DebugDelay',
      0);
  end;

  GetApplicationRegistryLocation(
    regPath,
    'DebugDelay');
  
  if JclRegistry.RegValueExists(HKLM, regPath, 'DebugDelay') then
  begin
    dDebug := JclRegistry.RegReadDWORD(
      HKLM,
      regPath,
      'DebugDelay');
    if dDebug = 1 then
    begin
      sleep(10000);
    end;
  end;

  USBDeviceTree1 := TUSBDeviceTree.Create(Self);
  with USBDeviceTree1 do
  begin
    Name := 'USBDeviceTree1';
    OnDeviceChange := USBDeviceTree1DeviceChange;
  end;
  
  Check_AndOr_Set_DisableSelectiveSuspendKey();

  edLevel := GetEjectDeviceLevel();
  winver := JclSysInfo.GetWindowsVersion;

  if not((winver >= wvWin2000) and (winver <= wvWin2003R2)) then
  begin
    
    Exit;
  end;

  UsersDomainGroup_Sids := TStringList.Create; 
  UsersDomainGroup_Sids.Duplicates := dupIgnore;
  UsersDomainGroup_Sids.Sorted := True;

  ApprovedUsbVidAndPid_ByGroup := TStringList.Create;
  ApprovedUsbVidAndPid_ByGroup.Duplicates := dupIgnore;
  ApprovedUsbVidAndPid_ByGroup.Sorted := True;

  DeniedUsbVidAndPidForUserOrComputer := TStringList.Create;
  DeniedUsbVidAndPidForUserOrComputer.Duplicates := dupIgnore;
  DeniedUsbVidAndPidForUserOrComputer.Sorted := True;

  RecentlyLoggedDevice := TStringList.Create;

  AlreadyLogged_WindowsEventLog_ApprovedDevices := TStringList.Create;
  AlreadyLogged_WindowsEventLog_ApprovedDevices.Duplicates := dupIgnore;
  AlreadyLogged_WindowsEventLog_ApprovedDevices.Sorted := True;

  JvLogFile1 := TJvLogFile.Create(Self);
  JvLogFile1.Name := 'JvLogFile1';
  LogFilePath := GetWindowsSystemFolder + '\LogFiles\UsbNinja';
  if FileExists(LogFilePath) then
    JvLogFile1.LoadFromFile(LogFilePath)
  else
    JvLogFile1.FileName := LogFilePath;
  
  if (VerboseLoggingEnabled()) then
  begin
    TimerClearTNetLogLists.Enabled := True;
  end
  else
  begin
    TimerClearTNetLogLists.Enabled := False;
  end;

  CheckDevices.Enabled := True;
  TimerClearEventLogLists.Enabled := True;

  Started := True;
end;

procedure TUsbNinjaSvc.ServiceStop(Sender: TService; var Stopped: boolean);
begin

  if assigned(USBDeviceTree1) then
  begin
    FreeAndNil(USBDeviceTree1);
  end;
  
  SaveLogFile();

  Stopped := True;
end;

initialization

sPreviousUserQueried := '';

finalization

end.
 