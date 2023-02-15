{**********************************************************}
{                                                          }
{  通用取码组件:TLYGetcode Component Version 04.10.27      }
{                                                          }
{  作者：刘鹰                                              }
{                                                          }
{                                                          }
{  新功能：1.当查询结果只有一条记录时，不显示取码框        }
{          2.按照openstr来决定输出哪些字段值               }
{          3.如果InValue中有汉字则不调用                   }
{          4.输出内容为TStrings类型                        }
{          5.20050609增加GetCodePos属性,                   }
{                决定从哪一部分开始匹配                    }
{          6.20050707增加IfNullGetCode属性,                }
{                决定当空值录入时是否取码                  }
{          7.20050827增加IfShowDialogOneRecord属性,        }
{                决定只查询到一条记录时是否显示对话框      }
{          8.20050906支持子查询                            }
{          9.20050912增加InFieldLabel属性,显示在录入框前   }
{          10.增加ifAbetChineseChar属性,是否可用汉字取码   }
{          11.20210501增加对MySQL的支持.因依赖MyDAC组件,   }
{              包路径增加***\VCL\MyDAC7.6.11\Source\Delphi7}
{              否则包编译出错,找不到mydac70.bpl与dac70.bpl }
{                                                          }
{  功能:                                                   }
{  修改历史：1.20050513修改了函数CombinSQL使用中文字段时的bug
{  调用方法：一般在编辑框的KeyDown事件中调用。             }
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
{  他是一个免费软件,如果你修改了他,希望我能有幸看到你的杰作}
{                                                          }
{  我的 Email: Liuying1129@163.com                         }
{                                                          }
{  版权所有,欲用于商业用途,请与我联系!!!                   }
{                                                          }
{  Bug:                                                    }
{  1.若有group by子句，要求group与by之间只能有一个空格     }
{  2.若有order by子句，要求order与by之间只能有一个空格     }
{  3.不支持where子句中的select from子查询                  }
{**********************************************************}

unit ADOLYGetcode;

interface

uses
  Windows, SysUtils, Forms,Classes,
  Grids, DBGrids, DB, ADODB,StdCtrls, Controls, ExtCtrls,StrUtils, Buttons,
  DBAccess, Uni, MemDS;

type
  TGetCodePos = (gcLeft,gcNone,gcRight,gcAll);

type
  TfrmADOGetcode = class(TForm)
    ADO_codestr: TADOQuery;
    MyQry_codestr: TUniQuery;
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
    FMyConnection:TUniConnection;
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
    procedure FSetMyConnection(value:TUniConnection);
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
    property MyConnection:TUniConnection read FMyConnection write FSetMyConnection;
    property OpenStr:string read FOpenStr write FSetOpenStr;
    property IfNullGetCode:boolean read FIfNullGetCode write FsetIfNullGetCode;
    Property IfShowDialogOneRecord:boolean read FIfShowDialogOneRecord write FsetIfShowDialogOneRecord;//只有一条记录时是否显示对话框
    Property IfShowDialogZeroRecord:boolean read FIfShowDialogZeroRecord write FsetIfShowDialogZeroRecord;//无记录时是否显示对话框
    property InValue:string read FInValue write FSetInValue;
    property InField:string read FInField write FSetInField;
    property InFieldLabel:string read FInFieldLabel write FSetInFieldLabel;
    property GetCodePos:TGetCodePos read fGetCodePos write fsetGetCodePos;
    property ifAbetChineseChar:boolean read FifAbetChineseChar write FsetifAbetChineseChar;//是否支持汉字取码
  end;

procedure Register;

implementation
  
{$R *.DFM}

procedure Register;
begin
  RegisterComponents('Eagle_Ly', [TADOLYGetcode]);
end;

function LastPos(const subStr,sourStr:string):integer;
//取得subStr在sourStr中最后一次出现的位置
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
  
  FinallyFromPos:=LastPos(' FROM ',UPPERCASE(tmpOpenStr));//决定了不支持where子句中的select from子查询

  FinallyOrderByPos:=LastPos(' ORDER BY ',UPPERCASE(tmpOpenStr));//这就要求order与by之间只能有一个空格
  if (FinallyOrderByPos<>0) and (FinallyOrderByPos>FinallyFromPos)then//取得order by子句
  begin
    OrderBy:=tmpOpenStr;
    delete(OrderBy,1,FinallyOrderByPos-1);
    tmpOpenStr:=copy(tmpOpenStr,1,FinallyOrderByPos-1);//去除了子句order by后的语句
  end;

  FinallyGroupByPos:=LastPos(' GROUP BY ',UPPERCASE(tmpOpenStr));//这就要求group与by之间只能有一个空格
  if (FinallyGroupByPos<>0) and (FinallyGroupByPos>FinallyFromPos) then//取得group by子句
  begin
    GroupBy:=tmpOpenStr;
    delete(GroupBy,1,FinallyGroupByPos-1);
    tmpOpenStr:=copy(tmpOpenStr,1,FinallyGroupByPos-1);//去除了子句group by后的语句
  end;

  FinallyWherePos:=LastPos(' WHERE ',UPPERCASE(tmpOpenStr));

  IF(FinallyWherePos<>0)and(FinallyWherePos>FinallyFromPos)THEN
    tmpOpenStr:=tmpOpenStr+' AND (' ELSE tmpOpenStr:=tmpOpenStr+' where (';

  TempInFieldFull:=pInField;

    while TempInFieldFull<>'' do
    begin
      InCommaPos:=pos(',',TempInFieldFull);
      //TempInField:=leftstr(TempInFieldFull,InCommaPos-1);
      //上面leftstr的bug：leftstr('姓名,工号',4)取出来的是'姓名,工',故改为copy
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
  
  if Ds_codestr.DataSet.RecordCount=0 then clearstringgrid(StringGrid1);//更新StringGrid中的内容
end;

procedure TfrmADOGetcode.LabeledEdit1KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  //向下键
  if key=40 then dbgrid1.SetFocus;
end;

procedure TfrmADOGetcode.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if key=13 then DBGrid1DblClick(DBGrid1);//设置窗体的KeyPreview为真才会触发该事件
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
  //向上键
  if(key=38)and((Sender as TDBGrid).DataSource.DataSet.Bof) then LabeledEdit1.SetFocus;
end;

procedure TfrmADOGetcode.ADO_codestrAfterScroll(DataSet: TDataSet);
//ADO_codestr、MyQry_codestr共用该事件
var
  i:integer;
  sFieldName:string;
  iField:integer;
begin
  //======更新StringGrid中的内容==============================================//
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
  FInFieldLabel:='代码';
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
  for i :=1  to length(fInValue) do  //编辑框中有汉字，则不调用取码框
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
  //==20060713确定窗体显示的位置
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

procedure TADOLYGetcode.FSetMyConnection(value: TUniConnection);
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
  //去掉stringgrid的蓝色选择条
  if gdSelected in State then
  begin
    TStringGrid(Sender).Canvas.Brush.Color:=TStringGrid(Sender).color;
    TStringGrid(Sender).Canvas.FillRect(rect);
    TStringGrid(Sender).Canvas.Font.Color:=0;//黑色
    TStringGrid(Sender).Canvas.TextRect(rect,Rect.Left+2,rect.top+2,TStringGrid(Sender).Cells[ACol,ARow]);
  end;
end;

end.
