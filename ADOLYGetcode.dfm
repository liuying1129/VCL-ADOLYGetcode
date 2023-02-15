object frmADOGetcode: TfrmADOGetcode
  Left = 131
  Top = 132
  BorderStyle = bsNone
  Caption = #20449#24687#36873#25321
  ClientHeight = 253
  ClientWidth = 427
  Color = clWindow
  Ctl3D = False
  Font.Charset = GB2312_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = #23435#20307
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = False
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel4: TPanel
    Left = 0
    Top = 0
    Width = 223
    Height = 253
    Align = alLeft
    TabOrder = 0
    object Panel1: TPanel
      Left = 1
      Top = 1
      Width = 221
      Height = 30
      Align = alTop
      BevelOuter = bvNone
      BorderStyle = bsSingle
      TabOrder = 0
      object LabeledEdit1: TLabeledEdit
        Left = 103
        Top = 5
        Width = 50
        Height = 19
        Color = clInfoBk
        EditLabel.Width = 7
        EditLabel.Height = 13
        EditLabel.Caption = 'L'
        LabelPosition = lpLeft
        TabOrder = 0
        OnChange = LabeledEdit1Change
        OnKeyDown = LabeledEdit1KeyDown
      end
      object BitBtn2: TBitBtn
        Left = 162
        Top = 3
        Width = 55
        Height = 24
        Caption = #21462#28040
        TabOrder = 1
        OnClick = BitBtn2Click
        Kind = bkCancel
      end
    end
    object DBGrid1: TDBGrid
      Left = 1
      Top = 31
      Width = 221
      Height = 221
      Align = alClient
      Color = clInfoBk
      Ctl3D = False
      DataSource = Ds_codestr
      ParentCtl3D = False
      ReadOnly = True
      TabOrder = 1
      TitleFont.Charset = GB2312_CHARSET
      TitleFont.Color = clWindowText
      TitleFont.Height = -13
      TitleFont.Name = #23435#20307
      TitleFont.Style = []
      OnDblClick = DBGrid1DblClick
      OnKeyDown = DBGrid1KeyDown
    end
  end
  object StringGrid1: TStringGrid
    Left = 223
    Top = 0
    Width = 204
    Height = 253
    Align = alClient
    Color = clInfoBk
    ColCount = 2
    DefaultColWidth = 100
    RowCount = 1
    FixedRows = 0
    TabOrder = 1
    OnDrawCell = StringGrid1DrawCell
    OnSelectCell = StringGrid1SelectCell
    RowHeights = (
      24)
  end
  object ADO_codestr: TADOQuery
    AfterScroll = ADO_codestrAfterScroll
    Parameters = <>
    Left = 112
    Top = 72
  end
  object Ds_codestr: TDataSource
    Left = 152
    Top = 72
  end
  object MyQry_codestr: TUniQuery
    AfterScroll = ADO_codestrAfterScroll
    Left = 112
    Top = 104
  end
end
