unit WindowInfoList;

{
  Unit Name   : WindowInfoList
  Description: 現在のデスクトップに存在するウィンドウの一覧を取得するクラスを提供します。
               各ウィンドウのハンドル、タイトル、クラス名、プロセスIDなどを取得可能です。

  Author     : vramwiz
  Created    : 2025-07-10
  Updated    : 2025-07-10

  Usage      :
    - TWindowInfoList.Create によりウィンドウ一覧を収集
    - Items プロパティから各ウィンドウの情報（ハンドル、タイトルなど）にアクセス可能
    - フィルタ処理や非表示ウィンドウの除外など、条件に応じた拡張も可能

  Dependencies:
    - Windows, Messages, TlHelp32, PsAPI など（Win32 API 使用）

  Notes:
    - HWNDごとの情報を構造体またはレコード単位で保持
    - マルチモニタ／仮想デスクトップなどの特殊環境は未対応（必要に応じて拡張可）

}

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls,ComCtrls,System.Generics.Collections,
  PsAPI,Winapi.ShellAPI, Winapi.ActiveX, Winapi.ShlObj, Winapi.CommCtrl, Vcl.ImgList;

type
  TLoadWindowsOption = (
    lwUWP,          // UWPアプリも含める
    lwDuplicates,   // 同一プロセスは1つに絞る
    lwSelfProcess   // 自分自身も表示する
  );
  TLoadWindowsOptions = set of TLoadWindowsOption;


type
  TWindowInfo = class(TPersistent)
  private
    FHandle      : HWND;              // ウインドウハンドル
    FProcessID   : DWORD;             // プロセスID
    FProcessName : string;            // プロセス名
    FTitle       : string;            // ウインドウタイトル
    FExeName     : string;            // 実行ファイルフルパス
    // True:UWPアプリ
    function GetIsUWP: Boolean;
    // アイコンを取得（UWP専用）
    procedure GetProcessIconUWP(Icon : TIcon);
    // アイコンを取得（UWP以外）
    procedure GetProcessIconNonUWP(Icon : TIcon);
    // 生成とウインドウハンドルのチェックを同時に行う
    class function CreateFromWindow(hWnd: HWND): TWindowInfo;
  public
    // アイコンを取得 引数のIconは生成済みのものを参照させる
    procedure GetProcessIcon(Icon : TIcon);
    // ウインドウハンドル
    property Handle      : HWND    read FHandle;
    // プロセスID
    property ProcessID   : DWORD   read FProcessID;
    // プロセス名
    property ProcessName : string  read FProcessName;
    // ウインドウタイトル
    property Title       : string  read FTitle;
    // 実行ファイルフルパス
    property ExeName     : string  read FExeName;
    // True:UWPアプリ
    property IsUWP       : Boolean read GetIsUWP;
  end;

type
	TWindowInfoList = class(TObjectList<TWindowInfo>)
	private
		{ Private 宣言 }
    FOptions       : TLoadWindowsOptions;         // 表示オプション

    procedure AddProcessInfo(hWnd: HWND;PID: DWORD);
    // リストに追加するかの総合判定
    function IsSkipSystemProcess(hWnd: HWND;PID: DWORD) : Boolean;
    // 自分自身のウインドウハンドルか判定
    function IsSelfProcess(PID: DWORD): Boolean;
    // 通常のトップレベルウインドウかどうかを判定する
    function IsDisplayableWindow(hWnd: HWND): Boolean;
    // UWPかどうかの判定
    function IsSystemUWP(hWnd: HWND) : Boolean;
    // システムプロセスかどうかの判断
    function IsSystemShellWindow(hWnd: HWND) : Boolean;
	public
		{ Public 宣言 }
    // 有効なアプリのリストを作成
    procedure LoadActiveWindows(const Options: TLoadWindowsOptions = []);
    // プロセスIDが一致するインデックス値を取得
    function IndexOfProcessID(PID: DWORD): Integer;
    // ハンドルが一致するインデックス値を取得
    function IndexOfProcessHandle(Handle      : HWND ): Integer;
    // タイトルが一致するインデックス値を取得
    function IndexOfTitle(const Title: string): Integer;
	end;

implementation


uses
  Winapi.Dwmapi,Winapi.propSys,Winapi.PropKey;


  function QueryFullProcessImageNameW(Process: THandle; Flags: DWORD; Buffer: PChar;
    Size: PDWORD): DWORD; stdcall; external 'kernel32.dll';


{ TWindowInfo }

//-----------------------------------------------------------------------------
//  ウィンドウのクラス名を取得する関数
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
//  ウィンドウのタイトルを取得する関数
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
// Windowsのシステムイメージリスト（アイコン一覧）を取得する関数
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
// 実行ファイル名を取得
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
  // クラス名 ApplicationFrameWindow のウインドウは UWP アプリ
  LIsUWPApp := GetWindowClassName(hWindow) = UWP_FRAMEWND;

  if LIsUWPApp then begin
    // AppUserModelId の値を取得
    if SHGetPropertyStoreForWindow(hWindow,
                                IID_IPropertyStore,
                                Pointer(LPropStore)) <> S_OK then Exit;
    if LPropStore.GetValue(PKEY_AppUserModel_ID, LPropVar) <> S_OK then Exit;
    LAppPath := LPropVar.bstrVal;
    LAppPath := 'shell:AppsFolder\' + LAppPath;
  end else begin
    // プロセスIDを取得
    // その値からプロセスのオープンハンドル取得
    GetWindowThreadProcessId(hWindow, Addr(LProcessID));
    LhProcess := OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, LProcessID);
    try
      // EXE のフルパスを取得
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
// アプリのリストを取得するコールバック関数
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
// 生成とウインドウハンドルのチェックを同時に行う
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
  // プロセス名取得
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
// True:UWPアプリ  本来のやり方では判定出来ないので強制判定
//-----------------------------------------------------------------------------
function TWindowInfo.GetIsUWP: Boolean;
begin
  Result := SameText(ProcessName, 'ApplicationFrameHost.exe') or
            (Pos('設定', Title) > 0) or
            (Pos('Windows 入力エクスペリエンス', Title) > 0);
end;


//-----------------------------------------------------------------------------
// アイコンを取得 引数のIconは生成済みのものを参照させる
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
// アイコンを取得（UWP以外）
//-----------------------------------------------------------------------------
procedure TWindowInfo.GetProcessIconNonUWP(Icon: TIcon);
var
  i : Integer;
  LSHFileInfo : TSHFileInfo;
  list    : HIMAGELIST;
begin
  // ファイル名 (またはファイルの関連付け) から取得
  SHGetFileInfo(PChar(ExeName),
                        0,
                        LSHFileInfo,
                        SizeOf(LSHFileInfo),
                        SHGFI_ICON or SHGFI_SMALLICON or SHGFI_SHELLICONSIZE);

  //LIcon.Handle :=  LSHFileInfo.hIcon;

  i := LSHFileInfo.iIcon;
  if i = -1 then i := 0;                                  // アイコンデータがない場合は未処理
  list := GetSystemImageList(SHIL_LARGE);                 // 指定サイズのアイコンを格納したシステムのイメージリストを取得
  Icon.Handle := ImageList_GetIcon(list, i, ILD_NORMAL); // イメージリストから指定インデックスのアイコンのハンドルを取得
end;

//-----------------------------------------------------------------------------
// アイコンを取得（UWP専用）
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
  if i = -1 then i := 0;                                  // アイコンデータがない場合は未処理
  list := GetSystemImageList(SHIL_LARGE);                 // 指定サイズのアイコンを格納したシステムのイメージリストを取得
  Icon.Handle := ImageList_GetIcon(list, i, ILD_NORMAL); // イメージリストから指定インデックスのアイコンのハンドルを取得

  CoTaskMemFree(LITemIDPath);

end;

{ TWindowInfoList }


procedure TWindowInfoList.AddProcessInfo(hWnd: HWND; PID: DWORD);
var
  Info: TWindowInfo;
begin
  // 条件に従い、表示非表示を判定
  if not IsSkipSystemProcess(hWnd,PID) then exit;

  Info := TWindowInfo.CreateFromWindow(hWnd);
  if Assigned(Info) then
  begin
    if not (lwUWP in FOptions) and Info.IsUWP then Exit;
    inherited Add(Info);
  end;
end;

//-----------------------------------------------------------------------------
// リストから該当するプロセスIDのインデックス値を取得
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
// 通常のトップレベルウインドウかどうかを判定する True:トップレベルウインドウ
//-----------------------------------------------------------------------------
function TWindowInfoList.IsDisplayableWindow(hWnd: HWND): Boolean;
var
  exStyle: DWORD;
begin
  // ウインドウが表示状態 かつ オーナーを持たない
  Result := IsWindowVisible(hWnd) and (GetWindow(hWnd, GW_OWNER) = 0);
  if Result then begin
    exStyle := GetWindowLong(hWnd, GWL_EXSTYLE);
    Result := (exStyle and WS_EX_TOOLWINDOW) = 0;  // ツールウインドウでない
  end;
end;

//-----------------------------------------------------------------------------
// リストに表示すべきアプリか判定
//-----------------------------------------------------------------------------
function TWindowInfoList.IsSelfProcess(PID: DWORD): Boolean;
begin
  Result := PID = GetCurrentProcessId;
end;

function TWindowInfoList.IsSkipSystemProcess(hWnd: HWND;  PID: DWORD): Boolean;
begin
  result := False;
  // システム関係のアプリは表示しない
  if not IsSystemShellWindow(hWnd) then exit;

  // 自分自身は表示しない
  if not (lwSelfProcess in FOptions) then begin
    if IsSelfProcess(PID) then exit;
  end;

  // クラス名 ApplicationFrameWindow のウインドウは UWP アプリ
  if IsSystemUWP(hWnd) then begin
    if not (lwUWP in FOptions) then exit;
  end;

  if lwDuplicates in FOptions then begin
    if IndexOfProcessID(PID) <> -1 then Exit;    // PID の重複チェック
  end;

  result := True;
end;

//-----------------------------------------------------------------------------
// システムプロセスかどうかの判断
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
// UWPかどうかの判定
//-----------------------------------------------------------------------------
function TWindowInfoList.IsSystemUWP(hWnd: HWND): Boolean;
const
  UWP_FRAMEWND  = 'ApplicationFrameWindow';
begin
  result := GetWindowClassName(hWnd) = UWP_FRAMEWND;
end;

//-----------------------------------------------------------------------------
// 有効なアプリのリストを作成
//-----------------------------------------------------------------------------
procedure TWindowInfoList.LoadActiveWindows(const Options: TLoadWindowsOptions);
begin
  FOptions := Options;
  Clear();
  EnumWindows(@EnumWindowProc, LPARAM(Pointer(Self)));
end;

end.
