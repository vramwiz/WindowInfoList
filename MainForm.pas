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
    FListView   : TListView;                    // �`��惊�X�g�r���[
    FImages     : TImageList;                   // �摜�o�^��C���[�W���X�g
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
    { Private �錾 }
    FInfos           : TWindowInfoList;
    FThread           : TIconLoadThread;
    FIsThreaded      : Boolean;               // �X���b�h���s��
    procedure ThreadExecute();
    procedure ThreadFinish();
    procedure OnTreadFinish(Sender: TObject);
    // �w�肵���C���f�b�N�X�l�̃E�C���h�E���A�N�e�B�u��
    procedure SetActiveFocus(Index : Integer);
  public
    { Public �錾 }
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
  ListView1.Columns[0].Caption :=' �A�v���P�[�V����';
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
    s := d.Title;          // �A�v���^�C�g���擾
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
  if IsIconic(LAppWindow) then                  // �E�B���h�E���ŏ�������Ă��邩�m�F
    ShowWindow(LAppWindow, SW_RESTORE);         // �ŏ�������Ă���ꍇ�͌��ɖ߂�  LAppWindow := HWND(FListView.Items[i].Data);
  SetForegroundWindow(LAppWindow);
  PostMessage(LAppWindow, WM_SETFOCUS, 0, 0);
end;

procedure TFormMain.ThreadExecute;
begin
  if FThread <> nil then
    ThreadFinish;                               // ���� Free + nil �ς�
  FIsThreaded  := True;
  FThread := TIconLoadThread.Create(ListView1,ImageList1,OnTreadFinish);
end;

procedure TFormMain.ThreadFinish;
begin
  if FThread<>nil then begin                // �X���b�h�������ς݂̏ꍇ
    if not FThread.Finished then begin      // �X���b�h���I����Ă��Ȃ��ꍇ
      FThread.Terminate;                    // �X���b�h���I��������
      FThread.WaitFor;                      // �I����҂�
    end;
    FThread.Free;                           // �X���b�h��j��
    FThread := nil;                         // �X���b�h�𖳂�
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

  // �X���b�h�I���ʒm�iUI����ł͂Ȃ����A�O�̂���Queue�ł�OK�j
  if Assigned(FOnFinish) then
    TThread.Queue(nil, procedure
    begin
      FOnFinish(Self);
    end);
end;

end.
