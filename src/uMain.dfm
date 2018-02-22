object frmMain: TfrmMain
  Left = 0
  Top = 0
  Caption = 'frmMain'
  ClientHeight = 299
  ClientWidth = 635
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object pcMain: TcxPageControl
    Left = 0
    Top = 0
    Width = 635
    Height = 299
    Align = alClient
    TabOrder = 0
    Properties.ActivePage = tsProgrammClose
    Properties.CustomButtons.Buttons = <>
    ClientRectBottom = 295
    ClientRectLeft = 4
    ClientRectRight = 631
    ClientRectTop = 24
    object tsNewMail: TcxTabSheet
      Caption = 'tsNewMail'
      ImageIndex = 0
      object cxLabel1: TcxLabel
        Left = 152
        Top = 96
        Caption = #1055#1088#1080#1096#1083#1086' '#1087#1080#1089#1100#1084#1086' :):):)'
        ParentFont = False
        Style.Font.Charset = DEFAULT_CHARSET
        Style.Font.Color = clBlue
        Style.Font.Height = -37
        Style.Font.Name = 'Tahoma'
        Style.Font.Style = []
        Style.IsFontAssigned = True
        Transparent = True
      end
      object cxButton1: TcxButton
        Left = 152
        Top = 168
        Width = 367
        Height = 81
        Caption = #1057#1074#1077#1088#1085#1091#1090#1100
        TabOrder = 1
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -32
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        OnClick = cxButton1Click
      end
    end
    object tsProgrammClose: TcxTabSheet
      Caption = 'tsProgrammClose'
      ImageIndex = 1
      object cxLabel2: TcxLabel
        Left = 152
        Top = 96
        Caption = 'Outlook '#1079#1072#1082#1088#1099#1090' :(:(:('
        ParentFont = False
        Style.Font.Charset = DEFAULT_CHARSET
        Style.Font.Color = clBlue
        Style.Font.Height = -37
        Style.Font.Name = 'Tahoma'
        Style.Font.Style = []
        Style.IsFontAssigned = True
        Transparent = True
      end
    end
    object tsEmpty: TcxTabSheet
      Caption = 'tsEmpty'
      ImageIndex = 2
    end
  end
  object Timer1: TTimer
    OnTimer = Timer1Timer
    Left = 128
    Top = 56
  end
end
