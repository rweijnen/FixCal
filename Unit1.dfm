object Form1: TForm1
  Left = 353
  Top = 253
  Width = 475
  Height = 283
  BorderIcons = [biSystemMenu, biMinimize, biMaximize, biHelp]
  Caption = 'Fix Calendar'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  Icon.Data = {
    0000010001001010100001000400280100001600000028000000100000002000
    00000100040000000000C0000000000000000000000000000000000000000000
    0000000080000080000000808000800000008000800080800000C0C0C0008080
    80000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF000000
    000000000000000000000000000000007700000000000000F7F7000000000007
    F7F7F7000000000F7FF7F7F70000007F7F7FF7F7F70000F7FF7F7117F7F008F7
    F7FF1F71F7000C88F7F71F717F0000CC88F7F11F70000000CC88F7F7F0000000
    008888F70000000000008887000000000000008000000000000000000000FFFF
    0000F3FF0007E0FFE8E8E03F8F70C00F0000C00300008000FE8E8000FF700001
    00000001000000037FF8C003F700F0070000FC070000FF0F077FFFCF7000}
  OldCreateOrder = False
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 13
  object Gauge1: TGauge
    Left = 8
    Top = 80
    Width = 337
    Height = 25
    ForeColor = clBlue
    ParentShowHint = False
    Progress = 0
    ShowHint = True
  end
  object Label1: TLabel
    Left = 256
    Top = 152
    Width = 67
    Height = 13
    Caption = 'Shift (in hours)'
  end
  object Gauge2: TGauge
    Left = 8
    Top = 120
    Width = 337
    Height = 25
    ForeColor = clBlue
    ParentShowHint = False
    Progress = 0
    ShowHint = True
  end
  object Button3: TButton
    Left = 360
    Top = 48
    Width = 81
    Height = 25
    Caption = 'Start'
    TabOrder = 0
    OnClick = Button3Click
  end
  object Mbxfrom: TLabeledEdit
    Left = 8
    Top = 168
    Width = 225
    Height = 21
    EditLabel.Width = 36
    EditLabel.Height = 13
    EditLabel.Caption = 'Mailbox'
    LabelPosition = lpAbove
    LabelSpacing = 3
    TabOrder = 1
    Text = 'tstagnda'
  end
  object Button1: TButton
    Left = 360
    Top = 88
    Width = 81
    Height = 25
    Caption = 'Exit'
    TabOrder = 2
    OnClick = Button1Click
  end
  object StaticText1: TStaticText
    Left = 8
    Top = 16
    Width = 337
    Height = 49
    AutoSize = False
    BevelKind = bkSoft
    TabOrder = 3
  end
  object StatusBar1: TStatusBar
    Left = 0
    Top = 230
    Width = 467
    Height = 26
    Panels = <>
    SimplePanel = False
    SizeGrip = False
  end
  object ComboBox1: TComboBox
    Left = 256
    Top = 168
    Width = 89
    Height = 21
    ItemHeight = 13
    ItemIndex = 22
    TabOrder = 5
    Text = '-1'
    Items.Strings = (
      '-23'
      '-22'
      '-21'
      '-20'
      '-19'
      '-18'
      '-17'
      '-16'
      '-15'
      '-14'
      '-13'
      '-12'
      '-11'
      '-10'
      '-9'
      '-8'
      '-7'
      '-6'
      '-5'
      '-4'
      '-3'
      '-2'
      '-1'
      '0'
      '+1'
      '+2'
      '+3'
      '+4'
      '+5'
      '+6'
      '+7'
      '+8'
      '+9'
      '+10'
      '+11'
      '+12'
      '+13'
      '+14'
      '+15'
      '+16'
      '+17'
      '+18'
      '+19'
      '+20'
      '+21'
      '+22'
      '+23 ')
  end
  object ComboBox2: TComboBox
    Left = 8
    Top = 200
    Width = 337
    Height = 21
    ItemHeight = 13
    TabOrder = 6
    Text = 'Click to select from addresslist'
    OnEnter = ComboBox2Click
  end
  object OutlookApplication1: TOutlookApplication
    AutoConnect = False
    ConnectKind = ckRunningOrNew
    AutoQuit = False
    Left = 288
    Top = 65520
  end
  object Timer1: TTimer
    Interval = 50
    OnTimer = Timer1Timer
    Left = 256
    Top = 65520
  end
end
