{**********************************************************}
{                                                          }
{  ͨ��ȡ�����:TLYGetcode Component Version 04.10.27      }
{                                                          }
{  ���ߣ���ӥ                                              }
{                                                          }
{                                                          }
{  �¹��ܣ�1.����ѯ���ֻ��һ����¼ʱ������ʾȡ���        }
{          2.����openstr�����������Щ�ֶ�ֵ               }
{          3.���InValue���к����򲻵���                   }
{          4.�������ΪTStrings����                        }
{          5.20050609����GetCodePos����,                   }
{                ��������һ���ֿ�ʼƥ��                    }
{          6.20050707����IfNullGetCode����,                }
{                ��������ֵ¼��ʱ�Ƿ�ȡ��                  }
{          7.20050827����IfShowDialogOneRecord����,        }
{                ����ֻ��ѯ��һ����¼ʱ�Ƿ���ʾ�Ի���      }
{          8.20050906֧���Ӳ�ѯ                            }
{          9.20050912����InFieldLabel����,��ʾ��¼���ǰ   }
{          10.����ifAbetChineseChar����,�Ƿ���ú���ȡ��   }
{          11.20210501���Ӷ�MySQL��֧��.������MyDAC���,   }
{              ��·������***\VCL\MyDAC7.6.11\Source\Delphi7}
{              ������������,�Ҳ���mydac70.bpl��dac70.bpl }
{                                                          }
{  ����:                                                   }
{  �޸���ʷ��1.20050513�޸��˺���CombinSQLʹ�������ֶ�ʱ��bug
{  ���÷�����һ���ڱ༭���KeyDown�¼��е��á�             }
{begin                                                     }
{  if key<>13 then exit;                                   }
{  LYGetCode1.Connection:=ADOConnection1;                  }
{  LYGetCode1.OpenStr:='select name from sj_sccj where isuseful=1';
{  LYGetCode1.InField:='ID,WBM,PYM';                       }
{  LYGetCode1.InValue:=tLabeledEdit(sender).Text;          }
{                                                          }
{  if LYGetCode1.Execute then                              }
{  begin                                                   }
{    tLabeledEdit(SENDER).Text:=LYGetCode1.OutValue[0];    }
{    LabeledEdit6.Text:=LYGetCode1.OutValue[1];            }
{    LabeledEdit8.Text:=LYGetCode1.OutValue[2];            }
{    LabeledEdit3.Text:=LYGetCode1.OutValue[3];            }
{    LabeledEdit7.Text:=LYGetCode1.OutValue[4];            }
{    LabeledEdit5.Text:=LYGetCode1.OutValue[5];            }
{    LabeledEdit4.Text:=LYGetCode1.OutValue[6];            }
{    LabeledEdit9.Text:=LYGetCode1.OutValue[7];            }
{    LabeledEdit1.Text:=LYGetCode1.OutValue[8];            }
{  end;                                                    }
{end;                                                      }
{                                                          }
{                                                          }
{  ����һ��������,������޸�����,ϣ���������ҿ�����Ľ���}
{                                                          }
{  �ҵ� Email: Liuying1129@163.com                         }
{                                                          }
{  ��Ȩ����,��������ҵ��;,��������ϵ!!!                   }
{                                                          }
{  Bug:                                                    }
{  1.����group by�Ӿ䣬Ҫ��group��by֮��ֻ����һ���ո�     }
{  2.����order by�Ӿ䣬Ҫ��order��by֮��ֻ����һ���ո�     }
{  3.��֧��where�Ӿ��е�select from�Ӳ�ѯ                  }
{**********************************************************}

unit ADOLYGetcode;

interface

uses
  Windows, SysUtils, Forms,Classes,
  Grids, DBGrids, DB, ADODB,StdCtrls, Controls, ExtCtrls,StrUtils, Buttons,
  DBAccess, MemDS, MyAccess;

type
  TGetCodePos = (gcLeft,gcNone,gcRight,gcAll);

type
  TfrmADOGetcode = class(TForm)
    ADO_codestr: TADOQuery;
    MyQry_codestr: TMyQuery;
    Ds_codestr: TDataSource;
    Panel4: TPanel;
    Panel1: TPanel;
    LabeledEdit1: TLabeledEdit;
    DBGrid1: TDBGrid;
    StringGrid1: TStringGrid;
    BitBtn2: TBitBtn;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure LabeledEdit1Change(Sender: TObject);
    procedure LabeledEdit1KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DBGrid1KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure ADO_codestrAfterScroll(DataSet: TDataSet);
    procedure BitBtn2Click(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure StringGrid1SelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure DBGrid1DblClick(Sender: TObject);
    procedure StringGrid1DrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
  private
    { Private declarations }
  public
    pInValue,pInField,pInFieldLabel:string;
    pOpenstr:string;
    pResult:boolean;
    pOutValue:TStrings;
    pGetCodePos:tGetCodePos;
  end;

type
  TADOLYGetcode = class(TComponent)
  private
    { Private declarations }
    FConnection:TADOConnection;
    FMyConnection:TMyConnection;
    Fopenstr:STRING;
    FInValue:STRING;
    FInField:STRING;
    FInFieldLabel:string;
    FOutValue:TStrings;
    ffrmAdoGetcode: TfrmAdoGetcode;
    fGetCodePos: TGetCodePos;
    FIfNullGetCode:boolean;
    FIfShowDialogOneRecord:boolean;
    FIfShowDialogZeroRecord:boolean;
    FifAbetChineseChar:boolean;
    FShowX,FShowY:integer;
    procedure FSetConnection(value:TADOConnection);
    procedure FSetMyConnection(value:TMyConnection);
    procedure FSetOpenStr(value:string);
    procedure FSetInValue(value:string);
    procedure FSetInField(value:string);
    procedure FSetInFieldLabel(value:string);
    procedure fsetGetCodePos(CONST value:tGetCodePos);
    procedure FsetIfNullGetCode(CONST value:boolean);
    procedure FsetIfShowDialogOneRecord(const value:boolean);
    procedure FsetIfShowDialogZeroRecord(const value:boolean);
    procedure FsetifAbetChineseChar(const value:boolean);
    procedure FSetShowX(Const Value:integer);
    procedure FSetShowY(Const Value:integer);
  protected
    { Protected declarations }
  public
    { Public declarations }
    constructor create(aowner:tcomponent);override;
    destructor destroy;override;
    function Execute:boolean;
    property OutValue:TStrings read FOutValue;
    property ShowX:integer read FShowX write FSetShowX;
    property ShowY:integer read FShowY write FSetShowY;
  published
    { Published declarations }
    property Connection:TADOConnection read FConnection write FSetConnection;
    property MyConnection:TMyConnection read FMyConnection write FSetMyConnection;
    property OpenStr:string read FOpenStr write FSetOpenStr;
    property IfNullGetCode:boolean read FIfNullGetCode write FsetIfNullGetCode;
    Property IfShowDialogOneRecord:boolean read FIfShowDialogOneRecord write FsetIfShowDialogOneRecord;//ֻ��һ����¼ʱ�Ƿ���ʾ�Ի���
    Property IfShowDialogZeroRecord:boolean read FIfShowDialogZeroRecord write FsetIfShowDialogZeroRecord;//�޼�¼ʱ�Ƿ���ʾ�Ի���
    property InValue:string read FInValue write FSetInValue;
    property InField:string read FInField write FSetInField;
    property InFieldLabel:string read FInFieldLabel write FSetInFieldLabel;
    property GetCodePos:TGetCodePos read fGetCodePos write fsetGetCodePos;
    property ifAbetChineseChar:boolean read FifAbetChineseChar write FsetifAbetChineseChar;//�Ƿ�֧�ֺ���ȡ��
  end;

procedure Register;

implementation
  
{$R *.DFM}

procedure Register;
begin
  RegisterComponents('Eagle_Ly', [TADOLYGetcode]);
end;

function LastPos(const subStr,sourStr:string):integer;
//ȡ��subStr��sourStr�����һ�γ��ֵ�λ��
var
  sub,sour:string;
begin
  if Pos(subStr,sourStr)=0 then
  begin
    Result:=0;
    exit;
  end;
  sub:=ReverseString(subStr);
  sour:=ReverseString(sourStr);
  Result:=length(sourStr)-Pos(sub,sour)+1-length(subStr)+1;
end; 

function CombinSQL(const pOpenStr,pInField,tj:string;gcPos:tGetCodePos):STRING;
var
  TempInFieldFull,TempInField:string;
  InCommaPos,FinallyFromPos,FinallyWherePos,FinallyGroupByPos,FinallyOrderByPos:integer;
  tmpOpenStr,GroupBy,OrderBy:string;
  logicExp:string;
BEGIN
  tmpOpenStr:=pOpenStr;
  
  FinallyFromPos:=LastPos(' FROM ',UPPERCASE(tmpOpenStr));//�����˲�֧��where�Ӿ��е�select from�Ӳ�ѯ

  FinallyOrderByPos:=LastPos(' ORDER BY ',UPPERCASE(tmpOpenStr));//���Ҫ��order��by֮��ֻ����һ���ո�
  if (FinallyOrderByPos<>0) and (FinallyOrderByPos>FinallyFromPos)then//ȡ��order by�Ӿ�
  begin
    OrderBy:=tmpOpenStr;
    delete(OrderBy,1,FinallyOrderByPos-1);
    tmpOpenStr:=copy(tmpOpenStr,1,FinallyOrderByPos-1);//ȥ�����Ӿ�order by������
  end;

  FinallyGroupByPos:=LastPos(' GROUP BY ',UPPERCASE(tmpOpenStr));//���Ҫ��group��by֮��ֻ����һ���ո�
  if (FinallyGroupByPos<>0) and (FinallyGroupByPos>FinallyFromPos) then//ȡ��group by�Ӿ�
  begin
    GroupBy:=tmpOpenStr;
    delete(GroupBy,1,FinallyGroupByPos-1);
    tmpOpenStr:=copy(tmpOpenStr,1,FinallyGroupByPos-1);//ȥ�����Ӿ�group by������
  end;

  FinallyWherePos:=LastPos(' WHERE ',UPPERCASE(tmpOpenStr));

  IF(FinallyWherePos<>0)and(FinallyWherePos>FinallyFromPos)THEN
    tmpOpenStr:=tmpOpenStr+' AND (' ELSE tmpOpenStr:=tmpOpenStr+' where (';

  TempInFieldFull:=pInField;

    while TempInFieldFull<>'' do
    begin
      InCommaPos:=pos(',',TempInFieldFull);
      //TempInField:=leftstr(TempInFieldFull,InCommaPos-1);
      //����leftstr��bug��leftstr('����,����',4)ȡ��������'����,��',�ʸ�Ϊcopy
      TempInField:=copy(TempInFieldFull,1,InCommaPos-1);
      delete(TempInFieldFull,1,InCommaPos);
      if(InCommaPos=0)and(TempInFieldFull<>'')then
      begin
        TempInField:=TempInFieldFull;
        TempInFieldFull:='';
      end;
      case gcPos of
        gcAll  :tmpOpenStr:=tmpOpenStr+logicExp+TempInField+' =    '''+trim(tj)+''' ';
        gcLeft :tmpOpenStr:=tmpOpenStr+logicExp+TempInField+' like '''+trim(tj)+'%'' ';
        gcRight:tmpOpenStr:=tmpOpenStr+logicExp+TempInField+' like ''%'+trim(tj)+''' ';
      else      tmpOpenStr:=tmpOpenStr+logicExp+TempInField+' like ''%'+trim(tj)+'%'' ';
      end;
      logicExp:=' or ';
    end;
    tmpOpenStr:=tmpOpenStr+')';
    tmpOpenStr:=tmpOpenStr+' '+GroupBy+' '+OrderBy;

    result:=tmpOpenStr;
END;

function CombinOutValue(ADataSet:TDataSet):TStrings;
var
  sFieldName:string;
  i,iField:integer;
begin
  Result := TStringList.Create;
  
  if not ADataSet.Active then exit;
  if ADataSet.RecordCount=0 then exit;
  
  iField:=ADataSet.FieldCount;
  for i :=0  to iField-1 do
  begin
    sFieldName:=ADataSet.Fields[i].FieldName;
    Result.Add(trim(ADataSet.fieldbyname(sFieldName).AsString));
  end;
end;

procedure clearstringgrid(aa:tstringgrid);
var
  colnum:integer;
  i:integer;
begin
  colnum:=aa.ColCount;
  for i :=0  to colnum-1 do
  begin
    aa.Cols[i].Clear;
  end;
end;

procedure TfrmADOGetcode.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  action:=cafree;
  //if self<>nil then self:=nil;
end;

procedure TfrmADOGetcode.FormShow(Sender: TObject);
begin
  LabeledEdit1.EditLabel.Caption:=pInFieldLabel+'(&I)';
  LabeledEdit1.Text:=pInValue;
end;

procedure TfrmADOGetcode.LabeledEdit1Change(Sender: TObject);
var
  sqltemp:string;
begin
  Ds_codestr.DataSet.Close;

  ADO_codestr.SQL.Clear;
  MyQry_codestr.SQL.Clear;
  
  sqltemp:=CombinSQL(POPENSTR,pInField,LabeledEdit1.Text,pGetCodePos);

  ADO_codestr.SQL.Add(sqltemp);
  MyQry_codestr.SQL.Add(sqltemp);
  
  Ds_codestr.DataSet.Open;
  
  if Ds_codestr.DataSet.RecordCount=0 then clearstringgrid(StringGrid1);//����StringGrid�е�����
end;

procedure TfrmADOGetcode.LabeledEdit1KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  //���¼�
  if key=40 then dbgrid1.SetFocus;
end;

procedure TfrmADOGetcode.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if key=13 then DBGrid1DblClick(DBGrid1);//���ô����KeyPreviewΪ��Żᴥ�����¼�
end;

procedure TfrmADOGetcode.DBGrid1DblClick(Sender: TObject);
begin
  if(not (Sender as TDBGrid).DataSource.DataSet.Active)or((Sender as TDBGrid).DataSource.DataSet.RecordCount=0) then
  begin
    pResult:=false;
    exit;
  end;

  pOutValue:=CombinOutValue((Sender as TDBGrid).DataSource.DataSet);

  pResult:=true;
  close;
end;

procedure TfrmADOGetcode.DBGrid1KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  //���ϼ�
  if(key=38)and((Sender as TDBGrid).DataSource.DataSet.Bof) then LabeledEdit1.SetFocus;
end;

procedure TfrmADOGetcode.ADO_codestrAfterScroll(DataSet: TDataSet);
//ADO_codestr��MyQry_codestr���ø��¼�
var
  i:integer;
  sFieldName:string;
  iField:integer;
begin
  //======����StringGrid�е�����==============================================//
  if not DataSet.Active then exit;
  iField:=DataSet.FieldCount;
  StringGrid1.RowCount:=iField;
  for i :=0  to iField-1 do
  begin
    sFieldName:=DataSet.Fields[i].FieldName;
    StringGrid1.Cells[0,i]:=sFieldName;
    StringGrid1.Cells[1,i]:=DataSet.fieldbyname(sFieldName).AsString;
  end;
  //==========================================================================//
end;

procedure TfrmADOGetcode.BitBtn2Click(Sender: TObject);
begin
    pResult:=false;
    close;
end;

{ TLYAboutBox }

constructor TADOLYGetcode.create(aowner: tcomponent);
begin
  inherited Create(AOwner);
  fOutValue := TStringList.Create;
  fGetCodePos:=gcNone;
  FIfNullGetCode:=false;
  FIfShowDialogOneRecord:=false;
  FIfShowDialogZeroRecord:=true;
  FInFieldLabel:='����';
  FifAbetChineseChar:=false;
end;

destructor TADOLYGetcode.destroy;
begin
  fOutValue.Free;
  inherited Destroy;
end;

function TADOLYGetcode.Execute: boolean;
var
  i:integer;
  sqltemp:string;
begin
  if (trim(fInValue)='')and(not FIfNullGetCode) then
  begin
    result:=false;exit;
  end;
  for i :=1  to length(fInValue) do  //�༭�����к��֣��򲻵���ȡ���
  begin
    if (ord(InValue[i])>127)and(not FifAbetChineseChar) then begin result:=false;exit;end;
  end;
  //if ffrmGetcode=nil then
  ffrmADOGetcode:=TfrmADOGetcode.Create(nil);
  if Assigned(fconnection) then
  begin
    ffrmADOGetcode.ADO_codestr.Connection:=fconnection;
    ffrmADOGetcode.Ds_codestr.DataSet:=ffrmADOGetcode.ADO_codestr;
  end;
  if Assigned(fMyconnection) then
  begin
    ffrmADOGetcode.MyQry_codestr.Connection:=fMyconnection;
    ffrmADOGetcode.Ds_codestr.DataSet:=ffrmADOGetcode.MyQry_codestr;
  end;
  ffrmADOGetcode.Ds_codestr.DataSet.Close;
  ffrmADOGetcode.ADO_codestr.SQL.Clear;
  ffrmADOGetcode.MyQry_codestr.SQL.Clear;
  sqltemp:=CombinSQL(OPENSTR,InField,InValue,GetCodePos);
  ffrmADOGetcode.ADO_codestr.SQL.Add(sqltemp);
  ffrmADOGetcode.MyQry_codestr.SQL.Add(sqltemp);
  ffrmADOGetcode.Ds_codestr.DataSet.Open;
  if(ffrmADOGetcode.Ds_codestr.DataSet.RecordCount=0)and(not FIfShowDialogZeroRecord)then
  begin
    result:=false;
    ffrmADOGetcode.Free;
    ffrmADOGetcode:=nil;
    exit;
  end;
  if(ffrmADOGetcode.Ds_codestr.DataSet.RecordCount=1)and(not FIfShowDialogOneRecord)then
  begin
    result:=true;
    ffrmADOGetcode.pOutValue:=CombinOutValue(ffrmADOGetcode.Ds_codestr.DataSet);
    fOutValue:=ffrmADOGetcode.pOutValue;
    ffrmADOGetcode.Free;
    ffrmADOGetcode:=nil;
    exit;
  end;

  ffrmADOGetcode.pInField:=fInField;
  ffrmADOGetcode.pInFieldLabel:=FInFieldLabel;
  ffrmADOGetcode.pInValue:=fInValue;
  ffrmADOGetcode.pOpenstr:=fOpenstr;
  ffrmADOGetcode.pGetCodePos:=fGetCodePos;
  //==20060713ȷ��������ʾ��λ��
  ffrmADOGetcode.Left:=fShowX;
  ffrmADOGetcode.Top:=fShowY;
  //============================
  ffrmADOGetcode.ShowModal;
  fOutValue:=ffrmADOGetcode.pOutValue;
  result:=ffrmADOGetcode.pResult;
  ffrmADOGetcode.Free;
end;

procedure TADOLYGetcode.FSetConnection(value: TADOConnection);
begin
  if value=FConnection then exit;
  FConnection:=value;
end;

procedure TADOLYGetcode.FSetMyConnection(value: TMyConnection);
begin
  if value=FMyConnection then exit;
  FMyConnection:=value;
end;

procedure TADOLYGetcode.fsetGetCodePos(CONST value: tGetCodePos);
begin
  if value=fGetCodePos then exit;
  fGetCodePos:=value;
end;

procedure TADOLYGetcode.FsetifAbetChineseChar(const value: boolean);
begin
  if value=FifAbetChineseChar then exit;
  FifAbetChineseChar:=value;
end;

procedure TADOLYGetcode.FsetIfNullGetCode(const value: boolean);
begin
  if value=FIfNullGetCode then exit;
  FIfNullGetCode:=value;
end;

procedure TADOLYGetcode.FsetIfShowDialogOneRecord(const value: boolean);
begin
  if value=FIfShowDialogOneRecord then exit;
  FIfShowDialogOneRecord:=value;
end;

procedure TADOLYGetcode.FsetIfShowDialogZeroRecord(const value: boolean);
begin
  if value=FIfShowDialogZeroRecord then exit;
  FIfShowDialogZeroRecord:=value;
end;

procedure TADOLYGetcode.FSetInField(value: string);
begin
  if value=FInField then exit;
  FInField:=value;
end;

procedure TADOLYGetcode.FSetInFieldLabel(value: string);
begin
  if value=FInFieldLabel then exit;
  FInFieldLabel:=value;
end;

procedure TADOLYGetcode.FSetInValue(value: string);
begin
  if value=FInValue then exit;
  FInValue:=value;
end;

procedure TADOLYGetcode.FSetOpenStr(value: string);
begin
  if value=FOpenStr then exit;
  FOpenStr:=value;
end;

procedure TADOLYGetcode.FSetShowX(const Value: integer);
begin
  if value=FShowX then exit;
  FShowX:=value;
end;

procedure TADOLYGetcode.FSetShowY(const Value: integer);
begin
  if value=FShowY then exit;
  FShowY:=value;
end;

//initialization
//  ffrmGetcode:=nil;

procedure TfrmADOGetcode.StringGrid1SelectCell(Sender: TObject; ACol,
  ARow: Integer; var CanSelect: Boolean);
begin
  TStringGrid(Sender).Options:=TStringGrid(Sender).Options-[goEditing];
end;

procedure TfrmADOGetcode.StringGrid1DrawCell(Sender: TObject; ACol,
  ARow: Integer; Rect: TRect; State: TGridDrawState);
begin
  //ȥ��stringgrid����ɫѡ����
  if gdSelected in State then
  begin
    TStringGrid(Sender).Canvas.Brush.Color:=TStringGrid(Sender).color;
    TStringGrid(Sender).Canvas.FillRect(rect);
    TStringGrid(Sender).Canvas.Font.Color:=0;//��ɫ
    TStringGrid(Sender).Canvas.TextRect(rect,Rect.Left+2,rect.top+2,TStringGrid(Sender).Cells[ACol,ARow]);
  end;
end;

end.
