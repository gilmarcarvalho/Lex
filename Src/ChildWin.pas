unit CHILDWIN;

interface

uses Windows, Classes, Graphics, Forms, Controls, StdCtrls, Grids, DB, DBClient,
  DBGrids,SysUtils,Dialogs,Variants, ExtCtrls, ActnList, ComCtrls, ToolWin,
  Menus, ImgList,ComObj,StrUtils, System.Actions;

type
  TMDIChild = class(TForm)
    DBGrid1: TDBGrid;
    ClientDataSet1: TClientDataSet;
    ClientDataSet1MovType: TStringField;
    ClientDataSet1MovDate: TDateField;
    ClientDataSet1Value: TFloatField;
    ClientDataSet1Document: TStringField;
    ClientDataSet1Desc: TStringField;
    ClientDataSet1BankID: TStringField;
    ClientDataSet1AccountID: TStringField;
    ClientDataSet1AccountType: TStringField;
    ClientDataSet1InitialBalance: TFloatField;
    ClientDataSet1FinalBalance: TFloatField;
    DataSource1: TDataSource;
    Panel1: TPanel;
    Label1: TLabel;
    Label2: TLabel;
    Label4: TLabel;
    Label9: TLabel;
    ActionList1: TActionList;
    MainMenu1: TMainMenu;
    actPrint: TAction;
    actClose: TAction;
    ImageList1: TImageList;
    StaticText1: TStaticText;
    StaticText2: TStaticText;
    StaticText4: TStaticText;
    StaticText5: TStaticText;
    ClientDataSet1BankName: TStringField;
    ClientDataSet1DateStart: TDateField;
    ClientDataSet1DateEnd: TDateField;
    Label3: TLabel;
    StaticText3: TStaticText;
    Label5: TLabel;
    StaticText6: TStaticText;
    Label6: TLabel;
    StaticText7: TStaticText;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure DBGrid1TitleClick(Column: TColumn);
    procedure FormActivate(Sender: TObject);
  private
    { Private declarations }
  public
  OFXfile:string;
  BankID:integer;
  AccountID:String;
  AccountType:string;
  FinalBalance:Double;
  InitialBallance:Double;
  BankName:string;
  DateStart:TDate;
  DateEnd:TDate;
  Procedure readFile;
    { Public declarations }
  end;

implementation

{$R *.dfm}

uses uYMOFXReader, Main;

procedure TMDIChild.DBGrid1TitleClick(Column: TColumn);
var i:integer;
begin
  ClientDataset1.IndexFieldNames:=Column.FieldName;
  for I := 0 to dbgrid1.Columns.Count-1 do begin
    dbgrid1.Columns[i].Title.Font.Style:=[];
  end;
  Column.Title.Font.Style:=[fsBold];
end;

procedure TMDIChild.FormActivate(Sender: TObject);
begin
  Mainform.EnabledChildControls(True);
end;

procedure TMDIChild.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action := caFree;
  if mainform.MDIChildCount <=1 then
      Mainform.EnabledChildControls(False);
end;

Procedure TMDIChild.readFile;
var
  YMOFXReader1: TYMOFXReader;
  i:integer;
begin
  YMOFXReader1:= TYMOFXReader.create(self);
  YMOFXReader1.OFXFile := OFXFile;
  try
  YMOFXReader1.Process;
  clientDataset1.CreateDataSet;
  clientDataset1.DisableControls;
  for i := 0 to YMOFXReader1.Count-1 do  begin
    clientDataset1.Append;
    clientDataset1.Fields[0].Value:=YMOFXReader1.Get(i).MovType;
    clientDataset1.Fields[1].Value:=YMOFXReader1.Get(i).MovDate;
    clientDataset1.Fields[2].Value:=YMOFXReader1.Get(i).Value;
    clientDataset1.Fields[3].Value:=YMOFXReader1.Get(i).Document;
    clientDataset1.Fields[4].Value:=YMOFXReader1.Get(i).Desc;
    clientDataset1.Fields[5].Value:=YMOFXReader1.BankID;
    clientDataset1.Fields[6].Value:=YMOFXReader1.AccountID;
    clientDataset1.Fields[7].Value:=YMOFXReader1.AccountType;
    clientDataset1.Fields[8].Value:=YMOFXReader1.InitialBalance;
    clientDataset1.Fields[9].Value:=YMOFXReader1.FinalBalance;
    clientDataset1.Fields[10].Value:=YMOFXReader1.BankName;
    clientDataset1.Fields[11].Value:=YMOFXReader1.DateStart;
    clientDataset1.Fields[12].Value:=YMOFXReader1.DateEnd;
    clientDataset1.Post;
    clientDataset1.EnableControls;
  end;
  clientDataset1.EnableControls;
  BankID:=YMOFXReader1.BankID;
  AccountID:=leftStr(YMOFXReader1.AccountID,50);
  AccountType:=leftStr(YMOFXReader1.AccountType,50);
  FinalBalance:=YMOFXReader1.FinalBalance;
  InitialBallance:=YMOFXReader1.InitialBalance;
  bankName:=leftStr(YMOFXReader1.BankName,50);
  DateStart:=YMOFXReader1.DateStart;
  DateEnd:=YMOFXReader1.DateEnd;
  StaticText1.Caption:=inttostr(bankID);
  StaticText2.Caption:=accountID;
  StaticText4.Caption:=formatfloat('#,##0.00',InitialBallance);
  StaticText5.Caption:=formatfloat('#,##0.00',FinalBalance);
  StaticText3.Caption:=BankName;
  StaticText6.Caption:=formatdatetime('dd/MM/yyyy',DateStart);
  StaticText7.Caption:=formatdatetime('dd/MM/yyyy',DateEnd);
  caption:=ofxfile;
  except
    application.MessageBox('Não foi possível ler arquivo OFX.','Erro ao ler',mb_iconerror);
    close;
  end;
end;
end.
