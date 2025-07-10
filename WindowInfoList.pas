unit WindowInfoList;

{
  Unit Name   : WindowInfoList
  Description: ���݂̃f�X�N�g�b�v�ɑ��݂���E�B���h�E�̈ꗗ���擾����N���X��񋟂��܂��B
               �e�E�B���h�E�̃n���h���A�^�C�g���A�N���X���A�v���Z�XID�Ȃǂ��擾�\�ł��B

  Author     : vramwiz
  Created    : 2025-07-10
  Updated    : 2025-07-10

  Usage      :
    - TWindowInfoList.Create �ɂ��E�B���h�E�ꗗ�����W
    - Items �v���p�e�B����e�E�B���h�E�̏��i�n���h���A�^�C�g���Ȃǁj�ɃA�N�Z�X�\
    - �t�B���^�������\���E�B���h�E�̏��O�ȂǁA�����ɉ������g�����\

  Dependencies:
    - Windows, Messages, TlHelp32, PsAPI �ȂǁiWin32 API �g�p�j

  Notes:
    - HWND���Ƃ̏����\���̂܂��̓��R�[�h�P�ʂŕێ�
    - �}���`���j�^�^���z�f�X�N�g�b�v�Ȃǂ̓�����͖��Ή��i�K�v�ɉ����Ċg���j

}

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls,ComCtrls,System.Generics.Collections,
  PsAPI,Winapi.ShellAPI, Winapi.ActiveX, Winapi.ShlObj, Winapi.CommCtrl, Vcl.ImgList;

type
  TLoadWindowsOption = (
    lwUWP,          // UWP�A�v�����܂߂�
    lwDuplicates,   // ����v���Z�X��1�ɍi��
    lwSelfProcess   // �������g���\������
  );
  TLoadWindowsOptions = set of TLoadWindowsOption;


type
  TWindowInfo = class(TPersistent)
  private
    FHandle      : HWND;              // �E�C���h�E�n���h��
    FProcessID   : DWORD;             // �v���Z�XID
    FProcessName : string;            // �v���Z�X��
    FTitle       : string;            // �E�C���h�E�^�C�g��
    FExeName     : string;            // ���s�t�@�C���t���p�X
    // True:UWP�A�v��
    function GetIsUWP: Boolean;
    // �A�C�R�����擾�iUWP��p�j
    procedure GetProcessIconUWP(Icon : TIcon);
    // �A�C�R�����擾�iUWP�ȊO�j
    procedure GetProcessIconNonUWP(Icon : TIcon);
    // �����ƃE�C���h�E�n���h���̃`�F�b�N�𓯎��ɍs��
    class function CreateFromWindow(hWnd: HWND): TWindowInfo;
  public
    // �A�C�R�����擾 ������Icon�͐����ς݂̂��̂��Q�Ƃ�����
    procedure GetProcessIcon(Icon : TIcon);
    // �E�C���h�E�n���h��
    property Handle      : HWND    read FHandle;
    // �v���Z�XID
    property ProcessID   : DWORD   read FProcessID;
    // �v���Z�X��
    property ProcessName : string  read FProcessName;
    // �E�C���h�E�^�C�g��
    property Title       : string  read FTitle;
    // ���s�t�@�C���t���p�X
    property ExeName     : string  read FExeName;
    // True:UWP�A�v��
    property IsUWP       : Boolean read GetIsUWP;
  end;

type
	TWindowInfoList = class(TObjectList<TWindowInfo>)
	private
		{ Private �錾 }
    FOptions       : TLoadWindowsOptions;         // �\���I�v�V����

    procedure AddProcessInfo(hWnd: HWND;PID: DWORD);
    // ���X�g�ɒǉ����邩�̑�������
    function IsSkipSystemProcess(hWnd: HWND;PID: DWORD) : Boolean;
    // �������g�̃E�C���h�E�n���h��������
    function IsSelfProcess(PID: DWORD): Boolean;
    // �ʏ�̃g�b�v���x���E�C���h�E���ǂ����𔻒肷��
    function IsDisplayableWindow(hWnd: HWND): Boolean;
    // UWP���ǂ����̔���
    function IsSystemUWP(hWnd: HWND) : Boolean;
    // �V�X�e���v���Z�X���ǂ����̔��f
    function IsSystemShellWindow(hWnd: HWND) : Boolean;
	public
		{ Public �錾 }
    // �L���ȃA�v���̃��X�g���쐬
    procedure LoadActiveWindows(const Options: TLoadWindowsOptions = []);
    // �v���Z�XID����v����C���f�b�N�X�l���擾
    function IndexOfProcessID(PID: DWORD): Integer;
    // �n���h������v����C���f�b�N�X�l���擾
    function IndexOfProcessHandle(Handle      : HWND ): Integer;
    // �^�C�g������v����C���f�b�N�X�l���擾
    function IndexOfTitle(const Title: string): Integer;
	end;

implementation


uses
  Winapi.Dwmapi,Winapi.propSys,Winapi.PropKey;


  function QueryFullProcessImageNameW(Process: THandle; Flags: DWORD; Buffer: PChar;
    Size: PDWORD): DWORD; stdcall; external 'kernel32.dll';


{ TWindowInfo }

//-----------------------------------------------------------------------------
//  �E�B���h�E�̃N���X�����擾����֐�
//-----------------------------------------------------------------------------
function GetWindowClassName(hWindow: HWND): string;
var
  LBuff : array [0..MAX_PATH - 1] of Char;
begin
  Result := '';
  FillChar(LBuff, SizeOf(LBuff), #0);
  if (GetClassName(hWindow, @LBuff, MAX_PATH) > 0) then begin
    Result := LBuff;
  end;
end;

//-----------------------------------------------------------------------------
//  �E�B���h�E�̃^�C�g�����擾����֐�
//-----------------------------------------------------------------------------
function GetWindowTitle(hWindow: HWND): string;
var
  LBuff : array [0..MAX_PATH - 1] of Char;
begin
  Result := '';
  FillChar(LBuff, SizeOf(LBuff), #0);
  if (GetWindowText(hWindow, @LBuff, MAX_PATH) > 0) then begin
    Result := LBuff;
  end;
end;


//-----------------------------------------------------------------------------
// Windows�̃V�X�e���C���[�W���X�g�i�A�C�R���ꗗ�j���擾����֐�
//-----------------------------------------------------------------------------
function GetSystemImageList(ASHILValue: Cardinal): HIMAGELIST;
type
  TGetImageList = function(iImageList: Integer; const riid: TGUID; var ppv: Pointer): HRESULT; stdcall;
const
  IID_IImageList: TGUID = '{46EB5926-582E-4017-9FDF-E8998DAA0950}';
var
  aHandle      : THandle;
  GetImageList : TGetImageList;
  P               : Pointer;
begin
  Result := 0;
  P := nil;

  aHandle := LoadLibrary('Shell32.dll');
  if aHandle <> 0 then
  try
    @GetImageList := GetProcAddress(aHandle, 'SHGetImageList');
    if Assigned(GetImageList) then
    begin
      if Succeeded(GetImageList(ASHILValue, IID_IImageList, P)) then
        Result := HIMAGELIST(P);
    end;
  finally
    FreeLibrary(aHandle);
  end;
end;


//-----------------------------------------------------------------------------
// ���s�t�@�C�������擾
//-----------------------------------------------------------------------------
function GetAppExeFileName(hWindow: HWND): string;
const
  UWP_FRAMEWND  = 'ApplicationFrameWindow';
  PROCESS_QUERY_LIMITED_INFORMATION = $1000;
var
  LIsUWPApp   : Boolean;
  LPropStore  : IPropertyStore;
  LPropVar    : TPropVariant;
  LAppPath    : string;
  LProcessID  : DWORD;
  LhProcess   : THandle;
  LBuffer     : array [0..MAX_PATH - 1] of Char;
  LSTR_SIZE   : DWORD;
begin
  // �N���X�� ApplicationFrameWindow �̃E�C���h�E�� UWP �A�v��
  LIsUWPApp := GetWindowClassName(hWindow) = UWP_FRAMEWND;

  if LIsUWPApp then begin
    // AppUserModelId �̒l���擾
    if SHGetPropertyStoreForWindow(hWindow,
                                IID_IPropertyStore,
                                Pointer(LPropStore)) <> S_OK then Exit;
    if LPropStore.GetValue(PKEY_AppUserModel_ID, LPropVar) <> S_OK then Exit;
    LAppPath := LPropVar.bstrVal;
    LAppPath := 'shell:AppsFolder\' + LAppPath;
  end else begin
    // �v���Z�XID���擾
    // ���̒l����v���Z�X�̃I�[�v���n���h���擾
    GetWindowThreadProcessId(hWindow, Addr(LProcessID));
    LhProcess := OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, LProcessID);
    try
      // EXE �̃t���p�X���擾
      FillChar(LBuffer, SizeOf(LBuffer), #0);
      LSTR_SIZE := Length(LBuffer);
      LAppPath:='';
      if QueryFullProcessImageNameW(LhProcess,
                                    0,
                                    @LBuffer[0],
                                    @LSTR_SIZE) = 0 then Exit;
      LAppPath := LBuffer;
    finally
      CloseHandle(LhProcess);
    end;
  end;
  result := LAppPath;
end;


//-----------------------------------------------------------------------------
// �A�v���̃��X�g���擾����R�[���o�b�N�֐�
//-----------------------------------------------------------------------------
function EnumWindowProc(hWnd: HWND; lParam: LPARAM): BOOL; stdcall;
var
  PID: DWORD;
  WinList: TWindowInfoList;
begin
  Result := True;
  WinList := TObject(Pointer(lParam)) as TWindowInfoList;

  if not WinList.IsDisplayableWindow(hWnd) then Exit;

  GetWindowThreadProcessId(hWnd, @PID);

  WinList.AddProcessInfo(hWnd,PID);

end;


//-----------------------------------------------------------------------------
// �����ƃE�C���h�E�n���h���̃`�F�b�N�𓯎��ɍs��
//-----------------------------------------------------------------------------
class function TWindowInfo.CreateFromWindow(hWnd: HWND): TWindowInfo;
var
  hProcess: THandle;
  FileName: array[0..MAX_PATH - 1] of Char;
  aTitle: array[0..255] of Char;
   PID: DWORD;
  d : TWindowInfo;
begin
  result := nil;
  PID := 0;
  GetWindowThreadProcessId(hWnd, @PID);
  // �v���Z�X���擾
  hProcess := OpenProcess(PROCESS_QUERY_INFORMATION or PROCESS_VM_READ, False, PID);
  if hProcess <> 0 then begin
    if GetModuleFileNameEx(hProcess, 0, FileName, MAX_PATH) > 0 then begin
      GetWindowText(hWnd, aTitle, Length(aTitle));
      d := TWindowInfo.Create;
      d.FProcessID := PID;
      d.FProcessName := ExtractFileName(FileName);
      d.FExeName := GetAppExeFileName(hWnd);
      d.FHandle := hWnd;
      d.FTitle := aTitle;
      result := d;
    end;
    CloseHandle(hProcess);
  end;

end;

//-----------------------------------------------------------------------------
// True:UWP�A�v��  �{���̂����ł͔���o���Ȃ��̂ŋ�������
//-----------------------------------------------------------------------------
function TWindowInfo.GetIsUWP: Boolean;
begin
  Result := SameText(ProcessName, 'ApplicationFrameHost.exe') or
            (Pos('�ݒ�', Title) > 0) or
            (Pos('Windows ���̓G�N�X�y���G���X', Title) > 0);
end;


//-----------------------------------------------------------------------------
// �A�C�R�����擾 ������Icon�͐����ς݂̂��̂��Q�Ƃ�����
//-----------------------------------------------------------------------------
procedure TWindowInfo.GetProcessIcon(Icon: TIcon);
begin
  if IsUWP then begin
    GetProcessIconUWP(Icon);
  end
  else begin
    GetProcessIconNonUWP(Icon);
  end;
end;

//-----------------------------------------------------------------------------
// �A�C�R�����擾�iUWP�ȊO�j
//-----------------------------------------------------------------------------
procedure TWindowInfo.GetProcessIconNonUWP(Icon: TIcon);
var
  i : Integer;
  LSHFileInfo : TSHFileInfo;
  list    : HIMAGELIST;
begin
  // �t�@�C���� (�܂��̓t�@�C���̊֘A�t��) ����擾
  SHGetFileInfo(PChar(ExeName),
                        0,
                        LSHFileInfo,
                        SizeOf(LSHFileInfo),
                        SHGFI_ICON or SHGFI_SMALLICON or SHGFI_SHELLICONSIZE);

  //LIcon.Handle :=  LSHFileInfo.hIcon;

  i := LSHFileInfo.iIcon;
  if i = -1 then i := 0;                                  // �A�C�R���f�[�^���Ȃ��ꍇ�͖�����
  list := GetSystemImageList(SHIL_LARGE);                 // �w��T�C�Y�̃A�C�R�����i�[�����V�X�e���̃C���[�W���X�g���擾
  Icon.Handle := ImageList_GetIcon(list, i, ILD_NORMAL); // �C���[�W���X�g����w��C���f�b�N�X�̃A�C�R���̃n���h�����擾
end;

//-----------------------------------------------------------------------------
// �A�C�R�����擾�iUWP��p�j
//-----------------------------------------------------------------------------
procedure TWindowInfo.GetProcessIconUWP(Icon: TIcon);
var
  LITemIDPath : PItemIDList;
  LSHFileInfo : TSHFileInfo;
  i : Integer;
  list    : HIMAGELIST;
begin
  LITemIDPath := ILCreateFromPath(PWideChar(FExeName));
  SHGetFileInfo(Pointer(LITemIDPath),
                0,
                LSHFileInfo,
                SizeOf(TSHFileInfo),
                SHGFI_PIDL or SHGFI_ICON or SHGFI_SMALLICON);
  //LIcon.Handle := LSHFILeInfo.hIcon;

  i := LSHFileInfo.iIcon;
  if i = -1 then i := 0;                                  // �A�C�R���f�[�^���Ȃ��ꍇ�͖�����
  list := GetSystemImageList(SHIL_LARGE);                 // �w��T�C�Y�̃A�C�R�����i�[�����V�X�e���̃C���[�W���X�g���擾
  Icon.Handle := ImageList_GetIcon(list, i, ILD_NORMAL); // �C���[�W���X�g����w��C���f�b�N�X�̃A�C�R���̃n���h�����擾

  CoTaskMemFree(LITemIDPath);

end;

{ TWindowInfoList }


procedure TWindowInfoList.AddProcessInfo(hWnd: HWND; PID: DWORD);
var
  Info: TWindowInfo;
begin
  // �����ɏ]���A�\����\���𔻒�
  if not IsSkipSystemProcess(hWnd,PID) then exit;

  Info := TWindowInfo.CreateFromWindow(hWnd);
  if Assigned(Info) then
  begin
    if not (lwUWP in FOptions) and Info.IsUWP then Exit;
    inherited Add(Info);
  end;
end;

//-----------------------------------------------------------------------------
// ���X�g����Y������v���Z�XID�̃C���f�b�N�X�l���擾
//-----------------------------------------------------------------------------
function TWindowInfoList.IndexOfProcessID(PID: DWORD): Integer;
var
  i: Integer;
begin
  Result := -1;
  for i := 0 to Count - 1 do begin
    if Items[i].ProcessID = PID then Exit(i);
  end;
end;

function TWindowInfoList.IndexOfTitle(const Title: string): Integer;
var
  I: Integer;
  d : TWindowInfo;
begin
  for I := 0 to Count - 1 do begin
    d := Items[I];
    if Pos(Title, d.Title) > 0 then
      Exit(I);
  end;
  Result := -1;
end;

function TWindowInfoList.IndexOfProcessHandle(Handle: HWND): Integer;
var
  i: Integer;
begin
  Result := -1;
  for i := 0 to Count - 1 do begin
    if Items[i].Handle = Handle then Exit(i);
  end;
end;


//-----------------------------------------------------------------------------
// �ʏ�̃g�b�v���x���E�C���h�E���ǂ����𔻒肷�� True:�g�b�v���x���E�C���h�E
//-----------------------------------------------------------------------------
function TWindowInfoList.IsDisplayableWindow(hWnd: HWND): Boolean;
var
  exStyle: DWORD;
begin
  // �E�C���h�E���\����� ���� �I�[�i�[�������Ȃ�
  Result := IsWindowVisible(hWnd) and (GetWindow(hWnd, GW_OWNER) = 0);
  if Result then begin
    exStyle := GetWindowLong(hWnd, GWL_EXSTYLE);
    Result := (exStyle and WS_EX_TOOLWINDOW) = 0;  // �c�[���E�C���h�E�łȂ�
  end;
end;

//-----------------------------------------------------------------------------
// ���X�g�ɕ\�����ׂ��A�v��������
//-----------------------------------------------------------------------------
function TWindowInfoList.IsSelfProcess(PID: DWORD): Boolean;
begin
  Result := PID = GetCurrentProcessId;
end;

function TWindowInfoList.IsSkipSystemProcess(hWnd: HWND;  PID: DWORD): Boolean;
begin
  result := False;
  // �V�X�e���֌W�̃A�v���͕\�����Ȃ�
  if not IsSystemShellWindow(hWnd) then exit;

  // �������g�͕\�����Ȃ�
  if not (lwSelfProcess in FOptions) then begin
    if IsSelfProcess(PID) then exit;
  end;

  // �N���X�� ApplicationFrameWindow �̃E�C���h�E�� UWP �A�v��
  if IsSystemUWP(hWnd) then begin
    if not (lwUWP in FOptions) then exit;
  end;

  if lwDuplicates in FOptions then begin
    if IndexOfProcessID(PID) <> -1 then Exit;    // PID �̏d���`�F�b�N
  end;

  result := True;
end;

//-----------------------------------------------------------------------------
// �V�X�e���v���Z�X���ǂ����̔��f
//-----------------------------------------------------------------------------
function TWindowInfoList.IsSystemShellWindow(hWnd: HWND): Boolean;
const
  CLASS_BRIDGE  = 'Windows.UI.Composition.DesktopWindowContentBridge';
var
  LAppTitle   : string;
begin
  result := False;
  LAppTitle := GetWindowTitle(hWnd);
  if LAppTitle = '' then Exit;
  if SameText(LAppTitle, 'Program Manager') then Exit;
  if FindWindowEx(hWnd, 0, PChar(CLASS_BRIDGE), nil) <> 0 then Exit;
  result := True;
end;

//-----------------------------------------------------------------------------
// UWP���ǂ����̔���
//-----------------------------------------------------------------------------
function TWindowInfoList.IsSystemUWP(hWnd: HWND): Boolean;
const
  UWP_FRAMEWND  = 'ApplicationFrameWindow';
begin
  result := GetWindowClassName(hWnd) = UWP_FRAMEWND;
end;

//-----------------------------------------------------------------------------
// �L���ȃA�v���̃��X�g���쐬
//-----------------------------------------------------------------------------
procedure TWindowInfoList.LoadActiveWindows(const Options: TLoadWindowsOptions);
begin
  FOptions := Options;
  Clear();
  EnumWindows(@EnumWindowProc, LPARAM(Pointer(Self)));
end;

end.
