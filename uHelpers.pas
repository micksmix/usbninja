unit uHelpers;

interface

uses
  Windows,
  SysUtils,
  Classes,
  System.RegularExpressionsCore;

procedure GetGroupMembershipSids(var GroupsAndSids: TStringList); 
function GetAccountSID(sAccountName: string): string;
function UserIsLoggedOn(): boolean;
function IsCorrectlyFormattedGroup(sDomainAndGroup: string): boolean;
function ParseVidPidFromRegistry(sVidPidToParse: string): string;

function SDUSplitString(wholeString: string; var firstItem: string; var theRest: string;
  splitOn: char = ' '): boolean;
function GetGroupNameFromStringSid(Sddl_Sid: PChar; var domain: string; var
  user_or_group: string): boolean;
implementation

uses
  JwaSddl,
  JclSysInfo,
 JwaWinBase,
  JwaWinType,
  JwaWinNT,
  JwaAclApi,
  JwaAccCtrl,
  JwaWinError,
  
  StrUtils;

function SDUSplitString(wholeString: string; var firstItem: string; var theRest: string;
  splitOn: char = ' '): boolean;
begin
  Result := False;
  firstItem := wholeString;

  if pos(splitOn, wholeString) > 0 then
  begin
    firstItem := copy(wholeString, 1, (pos(splitOn, wholeString) - 1));
    theRest := copy(wholeString, length(firstItem) + length(splitOn) + 1,
      (length(wholeString) - (length(firstItem) + length(splitOn))));
    Result := True;
  end
  else
  begin
    theRest := '';
  end;
end;

function ParseVidPidFromRegistry(sVidPidToParse: string): string;
var
  pos: integer;
  sResult: string;
begin
  
  pos := AnsiPos('|', sVidPidToParse);
  sResult := StuffString(sVidPidToParse, pos, 1, '=');

  if (not AnsiContainsStr(sResult, '|')) then
  begin
    sResult := sResult + '|NO';
  end;

  Result := AnsiUpperCase(sResult);
end;

function MyGetShellProcessHandle: THandle;
var
  Pid: Longword;
  thResult: THandle;
begin
  Pid := JclSysInfo.GetPidFromProcessName(GetShellProcessName);

  thResult := 0; 

  if Pid = INVALID_HANDLE_VALUE then
  begin
    Result := thResult;
    Exit;
  end;
  
  thResult := OpenProcess(PROCESS_ALL_ACCESS, False, Pid);

  Result := thResult;
end;

function UserIsLoggedOn: boolean;
var
  Handle: THandle;
  bResult: boolean;
begin
  Handle := 0;
  try
    Handle := uHelpers.MyGetShellProcessHandle;
    bResult := False; 

    if Handle = 0 then
    begin
      bResult := False;
      
      OutputDebugString('No user logged on');
    end
    else
    begin
      bResult := True;
    end;
  finally
    CloseHandle(Handle);
  end;

  Result := bResult;
end;

function IsCorrectlyFormattedGroup(sDomainAndGroup: string): boolean;
var
  Regex: TPerlRegEx;
  bResult: boolean;
  sExtractedDomain, sExtractedGroup: string;
begin
  bResult := False;

  Regex := TPerlRegEx.Create();
  try
    begin
      Regex.RegEx := '^(.*)\\(.*)$';
      Regex.Options := [preCaseless, preMultiLine];
      Regex.Subject := sDomainAndGroup;
      if Regex.Match then
      begin
        if Regex.GroupCount >= 1 then
        begin
          sExtractedDomain := Regex.Groups[1];
          sExtractedGroup := Regex.Groups[2];

          if ((Length(Trim(sExtractedDomain)) > 0) and
            (Length(Trim(sExtractedGroup)) > 0) and
            (sExtractedGroup <> '?')) then
          begin
            bResult := True;
          end;
        end;
      end;
    end;
  finally
    FreeAndNil(Regex);
  end;

  if ((AnsiUpperCase(sDomainAndGroup) = '\EVERYONE') or
    (AnsiUpperCase(sDomainAndGroup) = 'EVERYONE')) then
  begin
    bResult := True;
  end
  else if ((AnsiUpperCase(sDomainAndGroup) = '\LOCAL') or
    (AnsiUpperCase(sDomainAndGroup) = 'LOCAL')) then
  begin
    bResult := True;
  end;

  Result := bResult;
end;

function GetAccountSID(sAccountName: string): string;
var
  domain: string;
  domainSize: DWORD;
  sid: PSID;
  sidSize: DWORD;
  use: DWORD;
  sidstr: PChar;
  sResult: string;
  bLookup: boolean;
begin
  sResult := '';
  bLookup := False;

  if Win32Platform <> VER_PLATFORM_WIN32_NT then
  begin
    Exit;
  end;

  if ((AnsiUpperCase(sAccountName) = '\EVERYONE') or
    (AnsiUpperCase(sAccountName) = 'EVERYONE')) then
  begin
    sResult := 'S-1-1-0';
    Result := sResult;
    Exit;
  end
  else if ((AnsiUpperCase(sAccountName) = '\LOCAL') or
    (AnsiUpperCase(sAccountName) = 'LOCAL')) then
  begin
    sResult := 'S-1-2-0';
    Result := sResult;
    Exit;
  end;

  try
    begin
      domainSize := 0;
      sidSize := 0;

      LookupAccountName(nil, PChar(sAccountName), nil, sidSize, nil, domainSize, use);
      sid := AllocMem(sidSize);
      try
        SetLength(domain, domainSize);
        try
          bLookup := (LookupAccountName(nil, PChar(sAccountName), sid,
            sidSize, PChar(domain), domainSize, use));
        except
          OutputDebugString(PChar('Error with "LookupAccountName" for: ' + sAccountName));
          
        end;

        if bLookup = True then
        begin
          JwaSddl.ConvertSIDtoStringSID(sid, sidstr);

          sResult := sidstr;
        end;
      finally
        FreeMem(sid);
      end;
    end;
  except
    OutputDebugString(PChar('Error retrieving Account SID for: ' + sAccountName));

  end;

  Result := sResult;
end;

procedure GetGroupMembershipSids(var GroupsAndSids: TStringList); 
var
  accessToken: THandle;
  MySid: PSID;
  groups: PTokenGroups;
  iGroup: integer;
  infoBufferSize: DWORD;
  success: BOOL;
  sidstr: PChar;
  sStringSid: string;
  Handle: THandle;
  
begin
  OutputDebugString('Entering get groupmembershipsids');
  Handle := uHelpers.MyGetShellProcessHandle;
  if Handle = 0 then 
  begin
    OutputDebugString('No user is logged on, exiting refresh');
    Exit;
  end;

  try
    begin
      try
        
        Win32Check(OpenProcessToken(Handle, TOKEN_QUERY, accessToken));
      except
        OutputDebugString('Error with "OpenProcessToken"');
          
        CloseHandle(accessToken);
        Exit;
      end;

      try
        begin
          
          if GetTokenInformation(accessToken, TokenGroups, nil, 0, infoBufferSize) or
            (GetLastError <> ERROR_INSUFFICIENT_BUFFER) then
          begin
            OutputDebugString('Error with "GetTokenInformation"');
            
            Exit;
          end;

          GetMem(groups, infoBufferSize);
          try
            begin
              success := GetTokenInformation(accessToken, TokenGroups,
                groups, infoBufferSize, infoBufferSize);
              
              if success then
              begin
                AllocateAndInitializeSid(@SECURITY_NT_AUTHORITY, 1,
                  SECURITY_AUTHENTICATED_USER_RID, 0, 0, 0, 0, 0, 0, 0, MySid);
{$R-}
                for iGroup := 0 to groups.GroupCount - 1 do
                begin
                  JwaSddl.ConvertSIDtoStringSID(groups.Groups[iGroup].Sid, sidstr);
                  sStringSid := sidstr;
                  
                  GroupsAndSids.Add(sStringSid); 

                end; 
{$R+}
                FreeSid(MySid);
              end;
            end;
          finally
            FreeMem(groups);
          end;
        end;
      finally
        CloseHandle(accessToken);
      end;
    end;
  finally
    CloseHandle(Handle);
  end;

end;

function GetGroupNameFromStringSid(Sddl_Sid: PChar; var domain: string;
  var user_or_group: string): boolean;
var
  sid: PSID;
  domainSize: DWORD;
  sidUse: SID_NAME_USE;
  userSize: DWORD;
begin
  JwaSddl.ConvertStringSidToSid(Sddl_Sid, sid);
  try
    begin
      userSize := 0;
      domainSize := 0;

      LookupAccountSid(nil, Sid, nil, userSize, nil, domainSize, sidUse);
      SetLength(user_or_group, userSize);
      SetLength(domain, domainSize);

      if not LookupAccountSID(nil, sid, PChar(user_or_group), userSize, PChar(domain),
        domainSize, sidUse) then
      begin
        if GetLastError = ERROR_NONE_MAPPED then
        begin
          user_or_group := '?';
        end;
      end;

      user_or_group := PChar(user_or_group);
      domain := PChar(domain);
    end;
  finally
    if Assigned(sid) then
      LocalFree(Cardinal(sid));
  end;

  Result := True;
end;

end.

 