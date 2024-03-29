unit uEjectionThread;

interface

uses
  SvcMgr,
  Classes,
  USBDeviceTree,
  Windows,
  JwaWinIoctl,
  SyncObjs,
  uDriveHelpers;

type
  TEjectDeviceThread = class(TThread)
  private
    
    fMyUsb: TUSBDevice;
    fUserName: string;
    fUserSid: string;
    fSerialNumber: string;
    fVid: string;
    fPid: string;
    fLocationInformation: string;
    fDescription: string;
    fDriveLetter: string;
  protected
    procedure Execute; override;
  public
    property MyUsb: TUSBDevice read fMyUsb write fMyUsb;
    property UserName: string read fUserName write fUserName;
    property UserSid: string read fUserSid write fUserSid;
    property SerialNumber: string read fSerialNumber write fSerialNumber;
    property Vid: string read fVid write fVid;
    property Pid: string read fPid write fPid;
    property LocationInformation: string read fLocationInformation
      write fLocationInformation;
    property Description: string read fDescription write fDescription;
    property DriveLetter: string read fDriveLetter write fDriveLetter;
  private
    procedure LogToEventLog(EventType: DWord; Category, ID: integer;
      sMessage: string = ''; sUserName: string = ''; sUserSid: string = '';
      sDescr: string = ''; sDriveLetter: string = ''; sSerialNum: string = '';
      sVid: string = ''; sPid: string = ''; sIsMassStorage: string = '';
      sLocationInfo: string = '');

    function FlushDataToDisk(sDriveLetter: string): boolean;
  end;

const
  Name: string = 'UsbNinjaSvc';

var
  DeviceList: TStringList;
  Crit_DeviceList, Crit_EjectDevice: TCriticalSection;

implementation

uses
  SysUtils,
  CfgMgr32,
  Cfg,
  main;

procedure TEjectDeviceThread.Execute;
var
  iEjectCounter: integer;
  bSuccess, bDataFlushed: boolean;
  res: CfgMgr32.CONFIGRET;
  VetoType: PNP_VETO_TYPE;
  VetoName: string;
  sDataFlushed: string;
  vetochar: Pchar;
begin

  Crit_DeviceList.Enter;
  DeviceList.Add(fVid + fPid + fUserSid);
  Crit_DeviceList.Leave;
  try
    Crit_EjectDevice.Enter;

    bDataFlushed := False;

    if not GetVolumeMountPointFromVidAndPid(fVid, fPid) then
    begin
      
      OutputDebugString(Pchar('Vid and Pid no longer found: ' + fVid +
        ' - ' + fPid));
      
      Exit;
    end;

    SetLength(
      VetoName,
      MAX_PATH);

    for iEjectCounter := 0 to 3 do
    begin
      
      vetochar := AllocMem(MAX_PATH);
      try
        res := CfgMgr32.CM_Request_Device_Eject(
          fMyUsb.DeviceInstance,
          @VetoType,
          vetochar,
          MAX_PATH,
          0);
        
        SetString(
          VetoName,
          vetochar,
          MAX_PATH);
      finally
        if Assigned(vetochar) then
          FreeMem(
            vetochar,
            MAX_PATH);
      end;

      bSuccess := ((res = CR_SUCCESS) and (VetoType = PNP_VetoTypeUnknown));

      SetLength(
        TrueBoolStrs,
        1);
      SetLength(
        FalseBoolStrs,
        1);

      TrueBoolStrs[0] := 'YES';
      FalseBoolStrs[0] := 'NO';

      sDataFlushed := BoolToStr(
        bDataFlushed,
        True);

      if bSuccess then
      begin
        LogToEventLog(
          EVENTLOG_INFORMATION_TYPE,
          103,
          1198,
          '*** DEVICE EJECTED ***',
          fUserName,
          fUserSid,
          fDescription + ' instance: ' + inttostr(MyUsb.DeviceInstance),
          fDriveLetter,
          fSerialNumber,
          fVid,
          fPid,
          '[' + inttostr(iEjectCounter + 1) + '] attempts',
          fLocationInformation);

        Break;
      end;

      if (not bSuccess) and (iEjectCounter = 3) then
      begin

        LogToEventLog(
          EVENTLOG_INFORMATION_TYPE,
          103,
          1198,
          '*** UNABLE TO EJECT ***',
          fUserName,
          fUserSid,
          fDescription,
          fDriveLetter,
          fSerialNumber,
          fVid,
          fPid,
          '[' + inttostr(iEjectCounter + 1) + '] attempts vetotype ' +
          inttostr(VetoType) + ' veto ' + VetoName,
          fLocationInformation);
        
      end;

      Sleep(100);
    end;
  finally
    Crit_EjectDevice.Leave;

    Crit_DeviceList.Enter;
    DeviceList.Delete(DeviceList.IndexOf(fVid + fPid + fUserSid));
    Crit_DeviceList.Leave;
  end;
end;

function TEjectDeviceThread.FlushDataToDisk(sDriveLetter: string): boolean;
var
  hDrive: THandle;
  S: string;
  OSFlushed: boolean;
  bResult: boolean;
begin
  bResult := False;
  S := '\\.\' + sDriveLetter + ':';
  
  hDrive := CreateFile(
    Pchar(S),
    GENERIC_READ or GENERIC_WRITE,
    FILE_SHARE_READ or FILE_SHARE_WRITE,
    nil,
    OPEN_EXISTING,
    0,
    0);
  OSFlushed := FlushFileBuffers(hDrive);

  CloseHandle(hDrive);

  if OSFlushed then
  begin
    bResult := True;
  end
  else
  begin
    bResult := False;
  end;

  Result := bResult;
end;

procedure TEjectDeviceThread.LogToEventLog(EventType: DWord;
  Category, ID: integer; sMessage: string = ''; sUserName: string = '';
  sUserSid: string = ''; sDescr: string = ''; sDriveLetter: string = '';
  sSerialNum: string = ''; sVid: string = ''; sPid: string = '';
  sIsMassStorage: string = ''; sLocationInfo: string = '');
var
  MyEventLogger: SvcMgr.TEventLogger;
  LogLevel: word;
begin
  LogLevel := 3; 

  MyEventLogger := SvcMgr.TEventLogger.Create(Name);
  try
    if EventType = EVENTLOG_ERROR_TYPE then
    begin
      LogLevel := 3; 
    end
    else
    begin
      LogLevel := 1;
    end;

    OutputDebugString(Pchar(sMessage + ',' + sUserName + ',' + sUserSid + ',' +
      sVid + ',' + sPid + ',' + sDescr + ',' + '[' + sDriveLetter + ':],' +
      sSerialNum + ',' + sIsMassStorage + ',' + sLocationInfo + ',' +
      inttostr(Category) + ',' + DateTimeToStr(Now)));

    MyEventLogger.LogMessage(
      #13 + #10 + #13 + #10 + sMessage + #13 + #10 + 'User Name: ' + sUserName +
      #13 + #10 + 'User SID: ' + sUserSid + #13 + #10 + 'Vendor ID: ' + sVid +
      #13 + #10 + 'Product ID: ' + sPid + #13 + #10 + 'Description: ' + sDescr +
      #13 + #10 + 'Drive: [' + sDriveLetter + ':]' + #13 + #10 +
      'Serial Number: ' + sSerialNum + #13 + #10 + 'Is Mass Storage?: ' +
      sIsMassStorage + #13 + #10 + 'Location: ' + sLocationInfo + #13 + #10 +
      'Event Category: ' + inttostr(Category) + #13 + #10 + 'Time of event: ' +
      DateTimeToStr(Now),
      EventType,
      Category,
      ID);

  finally
    FreeAndNil(MyEventLogger);
  end;
end;

initialization

DeviceList := TStringList.Create;
Crit_EjectDevice := TCriticalSection.Create;
Crit_DeviceList := TCriticalSection.Create;

finalization

FreeAndNil(DeviceList);
Crit_EjectDevice.Free;
Crit_DeviceList.Free;

end.
 