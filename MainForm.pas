unit MainForm;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs,WindowInfoList, Vcl.ComCtrls,
  System.ImageList, Vcl.ImgList;

type
  TAppInfo = record
    ProcessID    : DWORD;
    ProcessName  : string;
    ExeFilename  : string;
    WindowHandle : HWND;
    WindowTitle  : string;
  end;

type
  TIconLoadThread = class(TThread)
  private
    FListView   : TListView;                    // 描画先リストビュー
    FImages     : TImageList;                   // 画像登録先イメージリスト
    FOnFinish   : TNotifyEvent;
  protected
    procedure Execute; override;
  public
    constructor Create(ListView : TListView;ImageList : TImageList;OnFinish : TNotifyEvent);
    destructor Destroy;override;
  end;


type
  TFormMain = class(TForm)
    ListView1: TListView;
    ImageList1: TImageList;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure ListView1Click(Sender: TObject);
  private
    { Private 宣言 }
    FInfos           : TWindowInfoList;
    FThread           : TIconLoadThread;
    FIsThreaded      : Boolean;               // スレッド実行中
    procedure ThreadExecute();
    procedure ThreadFinish();
    procedure OnTreadFinish(Sender: TObject);
    // 指定したインデックス値のウインドウをアクティブに
    procedure SetActiveFocus(Index : Integer);
  public
    { Public 宣言 }
  end;

var
  FormMain: TFormMain;

implementation

{$R *.dfm}

{ TFormMain }

procedure TFormMain.FormCreate(Sender: TObject);
begin
  FInfos := TWindowInfoList.Create;

  ListView1.DoubleBuffered := True;
  ListView1.ViewStyle := vsReport;
  ImageList1.Width := 32;
  ImageList1.Height := 32;
  ImageList1.BkColor := ListView1.Color;
  ListView1.SmallImages := ImageList1;

  ListView1.Columns.Add;
  ListView1.Columns[0].Caption :=' アプリケーション';
  ListView1.Columns[0].Width   := 100;
  ListView1.RowSelect := True;
end;

procedure TFormMain.FormDestroy(Sender: TObject);
begin
  FInfos.Free;
end;

procedure TFormMain.FormShow(Sender: TObject);
var
  i,j: Integer;
  s : string;
  dl : TListItem;
  d : TWindowInfo;
begin
  ListView1.Columns[0].Width   := ClientWidth;

  FInfos.LoadActiveWindows([lwDuplicates,lwUWP]);
  ListView1.Items.BeginUpdate;
  ListView1.Clear;

  for j := 0 to FInfos.Count-1 do begin
    dl := ListView1.Items.Add();
    d := FInfos[j];
    s := d.Title;          // アプリタイトル取得
    dl.Caption := s;
    dl.SubItems.Clear;
    dl.Data := d;
    dl.SubItems.Add(s);
    dl.ImageIndex := -1;
  end;

  ListView1.Items.EndUpdate;
  ThreadExecute();
end;



procedure TFormMain.ListView1Click(Sender: TObject);
begin
  SetActiveFocus(ListView1.ItemIndex);
end;

procedure TFormMain.OnTreadFinish(Sender: TObject);
begin
  FIsThreaded  := False;
end;

procedure TFormMain.SetActiveFocus(Index: Integer);
var
  LAppWindow  : HWND;
  i : Integer;
  s : string;
begin
  i := Index;
  if i = -1 then exit;
  if i >= ListView1.Items.Count then exit;

  LAppWindow := FInfos[i].Handle;
  s := ListView1.Items[i].Caption;
  if IsIconic(LAppWindow) then                  // ウィンドウが最小化されているか確認
    ShowWindow(LAppWindow, SW_RESTORE);         // 最小化されている場合は元に戻す  LAppWindow := HWND(FListView.Items[i].Data);
  SetForegroundWindow(LAppWindow);
  PostMessage(LAppWindow, WM_SETFOCUS, 0, 0);
end;

procedure TFormMain.ThreadExecute;
begin
  if FThread <> nil then
    ThreadFinish;                               // 中で Free + nil 済み
  FIsThreaded  := True;
  FThread := TIconLoadThread.Create(ListView1,ImageList1,OnTreadFinish);
end;

procedure TFormMain.ThreadFinish;
begin
  if FThread<>nil then begin                // スレッドが生成済みの場合
    if not FThread.Finished then begin      // スレッドが終わっていない場合
      FThread.Terminate;                    // スレッドを終了させる
      FThread.WaitFor;                      // 終了を待つ
    end;
    FThread.Free;                           // スレッドを破棄
    FThread := nil;                         // スレッドを無に
  end;
end;

{ TIconLoadThread }

constructor TIconLoadThread.Create(ListView: TListView; ImageList: TImageList;
  OnFinish: TNotifyEvent);
begin
  inherited Create(False);
  FOnFinish  := OnFinish;

  FListView := ListView;
  FImages := ImageList;
  FreeOnTerminate := False;

end;

destructor TIconLoadThread.Destroy;
begin
  Terminate;

  inherited;
end;

procedure TIconLoadThread.Execute;
var
  i : Integer;
  d : TWindowInfo;
  aicon : TIcon;
begin
  for i := 0 to FListView.Items.Count - 1 do begin
    if Terminated or (FListView = nil) then Break;

    d := TWindowInfo(FListView.Items[i].Data);
    aicon := TIcon.Create;
    try
      d.GetProcessIcon(aIcon);

      TThread.Synchronize(nil, procedure
      var
        idx: Integer;
      begin
        idx := FImages.AddIcon(aIcon);
        FListView.Items[i].ImageIndex := idx;
      end);

    finally
      aicon.Free;
    end;
  end;

  // スレッド終了通知（UI操作ではないが、念のためQueueでもOK）
  if Assigned(FOnFinish) then
    TThread.Queue(nil, procedure
    begin
      FOnFinish(Self);
    end);
end;

end.
