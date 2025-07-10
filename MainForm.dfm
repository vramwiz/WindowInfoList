object FormMain: TFormMain
  Left = 0
  Top = 0
  Caption = #36215#21205#12375#12390#12356#12427'Window'#21462#24471
  ClientHeight = 231
  ClientWidth = 505
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object ListView1: TListView
    Left = 0
    Top = 0
    Width = 505
    Height = 231
    Align = alClient
    Columns = <>
    TabOrder = 0
    OnClick = ListView1Click
    ExplicitLeft = 120
    ExplicitTop = 80
    ExplicitWidth = 250
    ExplicitHeight = 150
  end
  object ImageList1: TImageList
    Left = 184
    Top = 128
  end
end
