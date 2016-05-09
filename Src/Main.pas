unit MAIN;

interface

uses Windows, SysUtils, Classes, Graphics, Forms, Controls, Menus,
  StdCtrls, Dialogs, Buttons, Messages, ExtCtrls, ComCtrls, StdActns,
  ActnList, ToolWin, ImgList,DBGrids,ComObj,DB,DBClient, AppEvnts,Variants,
  frxClass, frxDBSet, frxExportRTF, frxExportPDF, System.Actions;

type
  TMainForm = class(TForm)
    MainMenu1: TMainMenu;
    File1: TMenuItem;
    FileOpenItem: TMenuItem;
    FileCloseItem: TMenuItem;
    Window1: TMenuItem;
    Help1: TMenuItem;
    N1: TMenuItem;
    FileExitItem: TMenuItem;
    WindowCascadeItem: TMenuItem;
    WindowTileItem: TMenuItem;
    HelpAboutItem: TMenuItem;
    OpenDialog: TOpenDialog;
    FileSaveItem: TMenuItem;
    StatusBar: TStatusBar;
    ActionList1: TActionList;
    Exit: TAction;
    FileOpen1: TAction;
    WindowCascade1: TWindowCascade;
    WindowTileHorizontal1: TWindowTileHorizontal;
    HelpAbout1: TAction;
    FileClose1: TWindowClose;
    WindowTileVertical1: TWindowTileVertical;
    WindowTileItem2: TMenuItem;
    ToolBar2: TToolBar;
    ToolButton1: TToolButton;
    ToolButton3: TToolButton;
    ToolButton8: TToolButton;
    ImageList1: TImageList;
    actExportXLS: TAction;
    actPrint: TAction;
    actCloseOFX: TAction;
    ToolButton4: TToolButton;
    PopupMenu1: TPopupMenu;
    actPrintPreview: TAction;
    Visualizar1: TMenuItem;
    Imprimir1: TMenuItem;
    frxReport1: TfrxReport;
    frxDBDataset1: TfrxDBDataset;
    procedure FileNew1Execute(Sender: TObject);
    procedure FileOpen1Execute(Sender: TObject);
    procedure HelpAbout1Execute(Sender: TObject);
    procedure ExitExecute(Sender: TObject);
    procedure FileClose1Execute(Sender: TObject);
    procedure actExportXLSExecute(Sender: TObject);
    procedure actPrintExecute(Sender: TObject);
    procedure actCloseOFXExecute(Sender: TObject);
    procedure actPrintPreviewExecute(Sender: TObject);
    procedure frxReport1GetValue(const VarName: string; var Value: Variant);
  private
    { Private declarations }
    procedure CreateMDIChild(const Name: string);

  public
    { Public declarations }
    Procedure EnabledChildControls(EnabledControl:boolean);
    Procedure GridtoXLS;
 end;
var
  MainForm: TMainForm;

implementation

{$R *.dfm}

uses CHILDWIN, about;



Procedure TMainForm.EnabledChildControls(EnabledControl:boolean);
begin
  actExportXLS.Enabled:=EnabledControl;
  actPrint.Enabled:=EnabledControl;
  actPrintpreview.Enabled:=EnabledControl;

end;



Procedure TMainForm.GridtoXLS;
var linha:integer;
XApp : variant;
Sheet:variant;
g:tclientDataset;
begin
  XApp:=CreateOleObject('Excel.Application');
  XApp.Visible:=true;
  XApp.DisplayAlerts := False;
  XApp.WorkBooks.Add(-4167);
  XApp.WorkBooks[1].WorkSheets[1].Name:='Extrato';
  sheet:=XApp.WorkBooks[1].WorkSheets['Extrato'];
  Xapp.ScreenUpdating := True;
  Sheet.Range ['A1','J1'].font.Bold:=True;
  Sheet.Range ['A1','J1'].WrapText:=True;
  Sheet.Range['A1','J1'].VerticalAlignment := 2;
  Sheet.Range['A1','J1'].HorizontalAlignment := 3;
  Sheet.Range ['A1','J32000'].font.Size:=10;
  sheet.Cells[1,1]:='Data';
  sheet.Cells[1,2]:='Valor';
  sheet.Cells[1,3]:='Documento';
  sheet.Cells[1,4]:='Descrição';
 g:= MainForm.ActiveMDIChild.FindComponent('ClientDataSet1')as tclientDataset;
 g.First;
 while not g.Eof do begin
    linha:=g.RecNo+1;
    sheet.cells[linha ,1] := g.Fields[1].AsDateTime;
    sheet.cells[linha ,2] := g.Fields[2].AsCurrency;
    sheet.cells[linha ,3] := g.Fields[3].AsString;
    sheet.cells[linha ,4] := g.Fields[4].AsString;
    g.Next;
 end;
 Sheet.columns.Autofit;
end;


procedure TMainForm.actCloseOFXExecute(Sender: TObject);
begin
  MainFOrm.ActiveMDIChild.Close;
end;

procedure TMainForm.actExportXLSExecute(Sender: TObject);
begin
  try
    GridToXLS;
  finally
  end;
end;

procedure TMainForm.actPrintExecute(Sender: TObject);
var ds: tdatasource;
begin
  ds:=mainform.ActiveMDIChild.FindComponent('Datasource1') as Tdatasource;
  frxdbdataset1.DataSource:=ds;
  frxreport1.Print;
end;

procedure TMainForm.actPrintPreviewExecute(Sender: TObject);
var ds: tdatasource;
begin
  ds:=mainform.ActiveMDIChild.FindComponent('Datasource1') as Tdatasource;
  frxdbdataset1.DataSource:=ds;
  frxreport1.ShowReport;
end;

procedure TMainForm.CreateMDIChild(const Name: string);
var
  Child: TMDIChild;
begin
  { create a new MDI child window }
  Child := TMDIChild.Create(Application);
  Child.Caption := Name;
  if FileExists(Name) then begin
     Child.ofxfile:=name;
     Child.readFile;
  end;
end;

procedure TMainForm.FileNew1Execute(Sender: TObject);
begin
  CreateMDIChild('Arquivo' + IntToStr(MDIChildCount + 1));
end;

procedure TMainForm.FileOpen1Execute(Sender: TObject);
begin
  if OpenDialog.Execute then
    CreateMDIChild(OpenDialog.FileName);
end;

procedure TMainForm.frxReport1GetValue(const VarName: string;
  var Value: Variant);
begin
  if varname='appname'  then
    value:=mainform.Caption;
end;

procedure TMainForm.HelpAbout1Execute(Sender: TObject);
begin
  AboutBox.ShowModal;
end;

procedure TMainForm.FileClose1Execute(Sender: TObject);
begin
  MainForm.ActiveMDIChild.Close;
end;

procedure TMainForm.ExitExecute(Sender: TObject);
begin
  Application.Terminate;
end;



end.
