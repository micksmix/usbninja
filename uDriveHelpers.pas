
unit uDriveHelpers;

interface

uses
  Windows,
  Classes,
  SysUtils,
  SvcMgr,
  uHelpers,
  CfgMgr32,
  WinSvc;

function ModifyFloppy(bDisableFloppy: boolean): boolean;
function GetDeviceID(Inst: DEVINST): string;
procedure ModifyWindowsCdWriting(bEnableCdBurning: boolean);
function GetVolumeMountPointFromVidAndPid(sVid, sPid: string): boolean;
function GetVolumeNameForVolumeMountPointString(Name: string): string;
procedure FillDriveList(var DriveList: TStringList);
procedure FillInRemovableDriveMountPoints(var MountPoints: TStrings);
function ExtractBus(DeviceID: string): string;
function GetSymbolicName(Inst: DEVINST): string;
function ExtractNum(const SymbolicName, Prefix: string): integer;
function ExtractSerialNumber(SymbolicName: string): string;
function GetDriveInstanceID(MountPointName: string; var DeviceInst: DEVINST): boolean;
function GetDriveTotalSize(drive: char): int64;
function GetDriveFromSerial(sKnownUsbSerial: string; var sDriveLetter: string): boolean;

const
  Name: string = 'UsbNinjaSvc';
  sCdServiceName: string = 'ImapiService';

implementation

uses
  JwaWinBase,
  JwaWinType,
  JclRegistry,
  JclSvcCtrl,
  Cfg,
  SetupApi,
  JwaWinIoctl,
  StrUtils,
  System.RegularExpressionsCore;

function ServiceSetStartType(
  sMachine: string;
  sServiceName: string;
  dwStartType: DWord): Boolean;
var
  hSCManager: SC_Handle;
  hSCService: SC_Handle;
begin
  Result := False;
  hSCManager := OpenSCManager(PChar(sMachine), nil, SC_MANAGER_CONNECT);
  if (hSCManager > 0) then
  begin
    hSCService := OpenService(hSCManager, PChar(sServiceName), SERVICE_CHANGE_CONFIG);
    if (hSCService > 0) then
    begin
      Result := ChangeServiceConfig(
        hSCService,
        SERVICE_NO_CHANGE,
        dwstartType,
        SERVICE_NO_CHANGE, nil, nil, nil, nil, nil, nil, nil);
      CloseServiceHandle(hSCService);
    end;
    CloseServiceHandle(hSCManager);
  end;
end;

procedure uDriveHelpers_LogToEventLog(sMessage: string; EventType: DWord;
  Category, ID: integer);
var
  MyEventLogger: TEventLogger;
begin
  MyEventLogger := TEventLogger.Create(Name);
  try
    if EventType = EVENTLOG_ERROR_TYPE then
      OutputDebugString(PChar(',' + sMessage))
        
    else
      OutputDebugString(PChar(',' + sMessage));

      OutputDebugString(PChar(sMessage + #13 + #10
      + 'User Name: ' + #13 + #10
      + 'User SID: ' + #13 + #10
      + 'Vendor ID: ' + #13 + #10
      + 'Product ID: ' + #13 + #10
      + 'Description: ' + #13 + #10
      + 'Drive: []' + #13 + #10
      + 'Serial Number: ' + #13 + #10
      + 'Is Mass Storage?: ' + #13 + #10
      + 'Location: ' + #13 + #10
      + 'Event Category: ' + IntToStr(Category) + #13 + #10
      + 'Time of event: ' + DateTimeToStr(Now)));
    
  finally
    FreeAndNil(MyEventLogger);
  end;
end;

procedure ModifyWindowsCdWriting(bEnableCdBurning: boolean);
var
  SvcState: JclSvcCtrl.TJclServiceState;
begin

  try
    begin
      SvcState := JclSvcCtrl.GetServiceStatusByName('', sCdServiceName);

      if bEnableCdBurning = True then
      begin
        if SvcState <> JclSvcCtrl.ssRunning then
        begin
          ServiceSetStartType('', sCdServiceName, SERVICE_AUTO_START);
          JclSvcCtrl.StartServiceByName('', sCdServiceName);

          uDriveHelpers_LogToEventLog('"' + sCdServiceName + '" has just been started.',
            Windows.EVENTLOG_INFORMATION_TYPE, 103, 1198);
        end;
      end
      else
      begin 
        if SvcState <> JclSvcCtrl.ssStopped then
        begin
          ServiceSetStartType('', sCdServiceName, SERVICE_DISABLED);
          JclSvcCtrl.StopServiceByName('', sCdServiceName);

          uDriveHelpers_LogToEventLog('"' + sCdServiceName +
            '" has just been stopped and disabled.',
            Windows.EVENTLOG_INFORMATION_TYPE, 103, 1198);
          
        end;
      end;
    end;
  except
    
  end;

end;

function VolumeNameToDeviceName(const VolName: string): string;
var
  s: string;
  TargetPath: array[0..MAX_PATH] of WideChar;
  bSucceeded: Boolean;
begin
  Result := '';
  
  s := Copy(VolName, 5, Length(VolName) - 5);

  bSucceeded := QueryDosDeviceW(PWideChar(WideString(s)), TargetPath, MAX_PATH) <> 0;
  if bSucceeded then
  begin
    Result := TargetPath;
  end
  else
  begin
    
  end;

end;

function GetDeviceID(Inst: DEVINST): string;
var
  Buffer: PChar;
  Size: ULONG;
begin

  Buffer := nil;

  try
    CM_Get_Device_ID_Size(Size, Inst, 0); 
    Inc(Size);
    Buffer := AllocMem(Size * SizeOf(char));
    CM_Get_Device_ID(Inst, Buffer, Size, 0);
    Result := Buffer;
  finally
    if Assigned(Buffer) then
      FreeMem(Buffer);
  end;

end;

function ModifyFloppy(bDisableFloppy: boolean): boolean;
var
  bSuccess, bFlpyExists, bSfloppyExists: boolean;
  Value1, Value2: DWORD;
begin
  bSuccess := False; 
  bFlpyExists := False; 
  bSfloppyExists := False; 

  if JclRegistry.RegKeyExists(HKLM, 'System\CurrentControlSet\Services\Flpydisk') then
  begin
    if JclRegistry.RegValueExists(HKLM, 'System\CurrentControlSet\Services\Flpydisk',
      'Start') then
    begin
      Value1 := JclRegistry.RegReadDWORD(
        HKLM, 'System\CurrentControlSet\Services\Flpydisk', 'Start');
      bFlpyExists := True;
    end;
  end;

  if JclRegistry.RegKeyExists(HKLM, 'System\CurrentControlSet\Services\Sfloppy') then
  begin
    if JclRegistry.RegValueExists(HKLM, 'System\CurrentControlSet\Services\Sfloppy',
      'Start') then
    begin
      Value2 := JclRegistry.RegReadDWORD(
        HKLM, 'System\CurrentControlSet\Services\Sfloppy', 'Start');
      bSfloppyExists := True;
    end;
  end;

  if bDisableFloppy = False then
  begin
    
    if bFlpyExists then
    begin
      if Value1 <> 3 then
      begin
        JclRegistry.RegWriteDWORD(
          HKLM, 'System\CurrentControlSet\Services\Flpydisk', 'Start', 3);
        
      end;

      Sleep(200); 

      if JclRegistry.RegValueExists(
        HKLM, 'System\CurrentControlSet\Services\Flpydisk', 'Start') then
      begin
        Value1 := JclRegistry.RegReadDWORD(
          HKLM, 'System\CurrentControlSet\Services\Flpydisk', 'Start');

        if Value1 <> 3 then
        begin
          bSuccess := False;
        end
        else
        begin
          
          bSuccess := True;
        end;
      end;
    end;

    if bSfloppyExists then
    begin
      if Value2 <> 3 then
      begin
        JclRegistry.RegWriteDWORD(
          HKLM, 'System\CurrentControlSet\Services\Sfloppy', 'Start', 3);
        
      end;

      Sleep(200); 

      if JclRegistry.RegValueExists(
        HKLM, 'System\CurrentControlSet\Services\Sfloppy', 'Start') then
      begin
        Value2 := JclRegistry.RegReadDWORD(
          HKLM, 'System\CurrentControlSet\Services\Sfloppy', 'Start');

        if Value2 <> 3 then
        begin
          bSuccess := False;
        end
        else
        begin
          
          bSuccess := True;
        end;
      end;
    end;
  end
  else if bDisableFloppy = True then
  begin
    
    if bFlpyExists then
    begin
      if Value1 <> 4 then
      begin
        JclRegistry.RegWriteDWORD(
          HKLM, 'System\CurrentControlSet\Services\Flpydisk', 'Start', 4);
        
      end;

      Sleep(200); 

      if JclRegistry.RegValueExists(
        HKLM, 'System\CurrentControlSet\Services\Flpydisk', 'Start') then
      begin
        Value1 := JclRegistry.RegReadDWORD(
          HKLM, 'System\CurrentControlSet\Services\Flpydisk', 'Start');

        if Value1 <> 4 then
        begin
          bSuccess := False;
        end
        else
        begin
          bSuccess := True;
        end;
      end;
    end;

    if bSfloppyExists then
    begin
      if Value2 <> 4 then
      begin
        JclRegistry.RegWriteDWORD(
          HKLM, 'System\CurrentControlSet\Services\Sfloppy', 'Start', 4);
        
      end;

      Sleep(200); 

      if JclRegistry.RegValueExists(
        HKLM, 'System\CurrentControlSet\Services\Sfloppy', 'Start') then
      begin
        Value2 := JclRegistry.RegReadDWORD(
          HKLM, 'System\CurrentControlSet\Services\Sfloppy', 'Start');

        if Value2 <> 4 then
        begin
          bSuccess := False;
        end
        else
        begin
          bSuccess := True;
        end;
      end;
    end;
  end;

  Result := bSuccess;
end;

function GetVolumeNameForVolumeMountPointString(Name: string): string;
var
  Volume: array[0..MAX_PATH] of char;
begin
  FillChar(Volume[0], SizeOf(Volume), 0);
  JwaWinBase.GetVolumeNameForVolumeMountPoint(PChar(Name), @Volume[0], SizeOf(Volume));
  Result := Volume;
end;

function GetVolumeMountPointFromVidAndPid(sVid, sPid: string): boolean;
var
  StorageGUID: TGUID;
  PnPHandle: HDEVINFO;
  DevData: TSPDevInfoData;
  DeviceInterfaceData: TSPDeviceInterfaceData;
  FunctionClassDeviceData: PSPDeviceInterfaceDetailData;
  Success: longbool;
  Devn: integer;
  BytesReturned: DWORD;
  Inst: DEVINST;
  S, FileName, DevID: string; 
  GuidListArray: array of TGUID;
  iArrayCount: integer;
  Regex: TPerlRegex;
  reVid, rePid, reSerial: string;
  bFoundVidPidDevID: boolean;
  sTemp_regex, sIgnore: string;
begin
  Result := False;
  bFoundVidPidDevID := False;

  SetLength(GuidListArray, 10);
  GuidListArray[0] := JwaWinIoctl.GUID_DEVINTERFACE_CDROM;
  GuidListArray[1] := JwaWinIoctl.GUID_DEVINTERFACE_COMPORT;
  GuidListArray[2] := JwaWinIoctl.GUID_DEVINTERFACE_DISK;
  GuidListArray[3] := JwaWinIoctl.GUID_DEVINTERFACE_FLOPPY;
  GuidListArray[4] := JwaWinIoctl.GUID_DEVINTERFACE_MEDIUMCHANGER;
  GuidListArray[5] := JwaWinIoctl.GUID_DEVINTERFACE_PARTITION;
  GuidListArray[6] := JwaWinIoctl.GUID_DEVINTERFACE_STORAGEPORT;
  GuidListArray[7] := JwaWinIoctl.GUID_DEVINTERFACE_TAPE;
  GuidListArray[8] := JwaWinIoctl.GUID_DEVINTERFACE_VOLUME;
  GuidListArray[9] := JwaWinIoctl.GUID_DEVINTERFACE_WRITEONCEDISK;

  try
    begin
      for iArrayCount := 0 to 9 do
      begin
        StorageGUID := GuidListArray[iArrayCount];

        PnPHandle := SetupDiGetClassDevs(@StorageGUID, nil, 0, DIGCF_PRESENT or
          DIGCF_DEVICEINTERFACE);

        if PnPHandle = Pointer(INVALID_HANDLE_VALUE) then
        begin
          SetupDiDestroyDeviceInfoList(PnPHandle);
          Exit;
        end;

        Devn := 0;

        Regex := TPerlRegEx.Create();
        try
          begin
            Regex.RegEx := '^.*VID_(.*)&PID_(.*)\\(.*)$';
            Regex.Options := [preCaseless, preMultiLine];
            repeat
              DeviceInterfaceData.cbSize := SizeOf(TSPDeviceInterfaceData);
              Success := SetupDiEnumDeviceInterfaces(PnPHandle, nil, StorageGUID,
                Devn, DeviceInterfaceData);
              if Success then
              begin
                DevData.cbSize := SizeOf(DevData);
                BytesReturned := 0;
                SetupDiGetDeviceInterfaceDetail(PnPHandle, @DeviceInterfaceData,
                  nil, 0, BytesReturned, @DevData);
                if (BytesReturned <> 0) and (GetLastError =
                  Windows.ERROR_INSUFFICIENT_BUFFER) then
                begin
                  FunctionClassDeviceData := AllocMem(BytesReturned);
                  try
                    FunctionClassDeviceData.cbSize :=
                      SizeOf(TSPDeviceInterfaceDetailData);
                    if SetupDiGetDeviceInterfaceDetail(PnPHandle, @DeviceInterfaceData,
                      FunctionClassDeviceData, BytesReturned, BytesReturned, @DevData)
                        then
                    begin
                      FileName := PTSTR(@FunctionClassDeviceData.DevicePath[0]);
                      
                      Inst := DevData.DevInst;
                      CM_Get_Parent(Inst, Inst, 0);
                      DevID := GetDeviceID(Inst);

                      if not AnsiContainsStr(DevID, 'VID_') then
                      begin
                        CM_Get_Parent(Inst, Inst, 0); 
                        DevID := GetDeviceID(Inst);

                        if not AnsiContainsStr(DevID, 'VID_') then
                        begin
                          bFoundVidPidDevID := False; 
                        end
                        else
                        begin
                          bFoundVidPidDevID := True;
                        end;
                      end
                      else
                      begin
                        bFoundVidPidDevID := True;
                      end;

                      if bFoundVidPidDevID = True then
                      begin
                        Regex.Subject := DevID;
                        if Regex.Match then
                        begin
                          if Regex.GroupCount >= 1 then
                          begin
                            reVid := Regex.Groups[1];
                            rePid := Regex.Groups[2];
                            reSerial := Regex.Groups[3];

                            if AnsiContainsStr(reVid, '&') then
                            begin
                              sTemp_regex := reVid;
                              uHelpers.SDUSplitString(sTemp_regex, reVid, sIgnore, '&');
                            end;

                            if AnsiContainsStr(rePid, '&') then
                            begin
                              sTemp_regex := rePid;
                              uHelpers.SDUSplitString(sTemp_regex, rePid, sIgnore, '&');
                            end;

                          end;
                        end;
                        
                        begin
                          S := '\';
                          S := PTSTR(@FunctionClassDeviceData.DevicePath) + S;

                          if Length(S) < 2 then
                          begin
                            CM_Get_Parent(Inst, Inst, 0); 
                            DevID := GetDeviceID(Inst);
                            S := '\';
                            S := PTSTR(@FunctionClassDeviceData.DevicePath) + S;
                          end;
                          
                          if (sVid = reVid) and (sPid = rePid) then
                          begin
                                                      
                            Result := True;

                            SetupDiDestroyDeviceInfoList(PnPHandle);
                            break;
                            
                          end;
                        end;
                      end;

                    end;
                  finally
                    FreeMem(FunctionClassDeviceData);
                  end;
                end;
              end;

              Inc(Devn);
            until not Success;
          end;
        finally
          FreeAndNil(Regex);
          SetupDiDestroyDeviceInfoList(PnPHandle);
        end;
      end;
    end;
  finally
    
  end;
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
        break;
      end;
    end;

    Result := bResult;
  finally
    UnloadConfigManagerApi;
    UnloadSetupApi;
    FreeAndNil(DriveMountPoints);
  end;
end;

function GetDriveTotalSize(drive: char): int64;
var
  RootPath: array[0..4] of char;
  RootPtr: PChar;
  current_dir: string;
  Free_size, Total_size: int64;
begin
  RootPath[0] := Drive;
  RootPath[1] := ':';
  RootPath[2] := '\';
  RootPath[3] := #0;
  RootPtr := RootPath;
  current_dir := GetCurrentDir;
  if SetCurrentDir(drive + ':\') then
  begin
    Windows.GetDiskFreeSpaceEx(RootPtr, Free_size, Total_size, nil);
    
    SetCurrentDir(current_dir);
  end
  else
  begin
    Free_size := -1;
    Total_size := -1;
  end;

  Result := Total_size;
end;

procedure FillInRemovableDriveMountPoints(var MountPoints: TStrings);
const
  MAX_DRIVES = 26;
var
  I: integer;
  dwDriveMask: DWORD;
  DriveName: string;
  iDriveType: cardinal;
begin
  if MountPoints.Count > 0 then
    MountPoints.Clear;
  
  dwDriveMask := GetLogicalDrives;
  DriveName := 'A:\';
  
  for I := 0 to MAX_DRIVES - 1 do
    
    if (dwDriveMask and (1 shl I)) <> 0 then
    begin
      DriveName[1] := 'A';
      Inc(DriveName[1], I);
      
      iDriveType := GetDriveType(PChar(DriveName));
      if ((iDriveType = DRIVE_REMOVABLE) or (iDriveType = DRIVE_FIXED)) then
        
      begin
        
        MountPoints.AddObject(GetVolumeNameForVolumeMountPointString(DriveName),
          TObject(DriveName[1]));
      end;
    end;
end;

function ExtractBus(DeviceID: string): string;
begin
  Result := Copy(DeviceID, 1, Pos('\', DeviceID) - 1);
end;

function GetSymbolicName(Inst: DEVINST): string;
var
  Len: DWORD;
  Key: HKEY;
  
  Buffer: array[0..4095] of char;
begin
  CM_Open_DevNode_Key(Inst, KEY_READ, 0,
    REGDISPOSITION(RegDisposition_OpenExisting), Key, 0);
  Buffer[0] := #0;
  if Key <> INVALID_HANDLE_VALUE then
  begin
    Len := SizeOf(Buffer);
    RegQueryValueEx(Key, 'SymbolicName', nil, nil, @Buffer[0], @Len);
    RegCloseKey(Key);
  end;
  Result := Buffer;
end;

function ExtractNum(const SymbolicName, Prefix: string): integer;
var
  S: string;
  N: integer;
begin
  S := LowerCase(SymbolicName);
  N := Pos(Prefix, S);

  if N > 0 then
  begin
    S := '$' + Copy(SymbolicName, N + Length(Prefix), 4);
    Result := StrToInt(S);
  end
  else
    Result := 0;

end;

function ExtractSerialNumber(SymbolicName: string): string;
var
  N: integer;
begin
  N := Pos('#', SymbolicName);
  if N >= 0 then
  begin
    SymbolicName := Copy(SymbolicName, N + 1, Length(SymbolicName));
    N := Pos('#', SymbolicName);
    if N >= 0 then
    begin
      SymbolicName := Copy(SymbolicName, N + 1, Length(SymbolicName));
      N := Pos('#', SymbolicName);
      if N >= 0 then
        Result := Copy(SymbolicName, 1, N - 1)
      else
        Result := '';
    end;
  end
  else
    Result := '';
end;

function GetDriveInstanceID(MountPointName: string; var DeviceInst: DEVINST): boolean;
const
  GUID_DEVINTERFACE_VOLUME: TGUID = '{53f5630d-b6bf-11d0-94f2-00a0c91efb8b}';
var
  StorageGUID: TGUID;
  PnPHandle: HDEVINFO;
  DevData: TSPDevInfoData;
  DeviceInterfaceData: TSPDeviceInterfaceData;
  FunctionClassDeviceData: PSPDeviceInterfaceDetailData;
  Success: longbool;
  Devn: integer;
  BytesReturned: DWORD;
  Inst: DEVINST;
  S, FileName, MountName, DevID: string;
begin
  Result := False;
  DeviceInst := 0;
  
  StorageGUID := GUID_DEVINTERFACE_VOLUME;
  PnPHandle := SetupDiGetClassDevs(@StorageGUID, nil, 0, DIGCF_PRESENT or
    DIGCF_DEVICEINTERFACE);

  if PnPHandle = Pointer(INVALID_HANDLE_VALUE) then
  begin
    SetupDiDestroyDeviceInfoList(PnPHandle);
    Exit;
  end;

  Devn := 0;

  try
    begin
      repeat
        DeviceInterfaceData.cbSize := SizeOf(TSPDeviceInterfaceData);
        Success := SetupDiEnumDeviceInterfaces(PnPHandle, nil, StorageGUID,
          Devn, DeviceInterfaceData);
        if Success then
        begin
          DevData.cbSize := SizeOf(DevData);
          BytesReturned := 0;
          SetupDiGetDeviceInterfaceDetail(PnPHandle, @DeviceInterfaceData,
            nil, 0, BytesReturned, @DevData);
          if (BytesReturned <> 0) and (GetLastError =
            Windows.ERROR_INSUFFICIENT_BUFFER) then
          begin
            FunctionClassDeviceData := AllocMem(BytesReturned);
            try
              FunctionClassDeviceData.cbSize := SizeOf(TSPDeviceInterfaceDetailData);
              if SetupDiGetDeviceInterfaceDetail(PnPHandle, @DeviceInterfaceData,
                FunctionClassDeviceData, BytesReturned, BytesReturned, @DevData) then
              begin
                FileName := PTSTR(@FunctionClassDeviceData.DevicePath[0]);
                
                Inst := DevData.DevInst;
                CM_Get_Parent(Inst, Inst, 0);
                CM_Get_Parent(Inst, Inst, 0);
                DevID := GetDeviceID(Inst);
                
                begin
                  S := '\';
                  S := PTSTR(@FunctionClassDeviceData.DevicePath) + S;
                  MountName := GetVolumeNameForVolumeMountPointString(S);
                  if MountName = MountPointName then
                  begin

                    DeviceInst := Inst;
                    Result := True;
                    SetupDiDestroyDeviceInfoList(PnPHandle);
                    Exit;
                  end;
                end;
              end;
            finally
              FreeMem(FunctionClassDeviceData);
            end;
          end;
        end;

        Inc(Devn);
      until not Success;
    end;
  finally
    SetupDiDestroyDeviceInfoList(PnPHandle);
  end;
end;

procedure FillDriveList(var DriveList: TStringList);
var
  S: string;
  I: integer;
  Inst: DEVINST;
  SymbolicName: string;
  DriveMountPoints: TStrings;
  curCommaText: string;
begin
  DriveMountPoints := TStringList.Create;
  SymbolicName := '';
  try
    
    FillInRemovableDriveMountPoints(DriveMountPoints);

    S := 'A:';
    for I := 0 to DriveMountPoints.Count - 1 do
    begin
      S[1] := char(DriveMountPoints.Objects[I]);
      GetDriveInstanceID(DriveMountPoints[I], Inst);
      SymbolicName := GetSymbolicName(Inst);
        
      DriveList.AddObject(S + '' + ExtractSerialNumber(SymbolicName), TObject(Inst));
      curCommaText := DriveList.CommaText;
      DriveList.CommaText := curCommaText + ',' + ExtractSerialNumber(
        SymbolicName) + '=' + S;
    end;

  finally
    FreeAndNil(DriveMountPoints);
  end;
end;
end.

 