unit uSplash;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, jpeg, ExtCtrls, StdCtrls;

type
  TfSplash = class(TForm)
    Image1: TImage;
    Timer1: TTimer;
    Label1: TLabel;
    procedure Timer1Timer(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure FormClick(Sender: TObject);
  private
    { Private declarations }
    function Sto_GetFmtFileVersion(const FileName: String = ''; const Fmt: String = '%d.%d.%d'): String;
  public
    { Public declarations }
  end;

var
  fSplash: TfSplash;

implementation

uses Main;

{$R *.dfm}

procedure TfSplash.FormClick(Sender: TObject);
begin
  close;
end;

procedure TfSplash.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  freeandnil(fSplash);
end;

procedure TfSplash.Timer1Timer(Sender: TObject);
begin
  fSplash.Close;
end;

procedure TfSplash.FormShow(Sender: TObject);
var appversion:string;
begin
  appversion:=Sto_GetFmtFileVersion(application.ExeName);
  MainForm.Caption:='Lex - '+appversion;
  Label1.Caption:='Versão '+appversion
end;

function TfSplash.Sto_GetFmtFileVersion(const FileName: String = '';
  const Fmt: String = '%d.%d.%d'): String;
var
  sFileName: String;
  iBufferSize: DWORD;
  iDummy: DWORD;
  pBuffer: Pointer;
  pFileInfo: Pointer;
  iVer: array[1..4] of Word;
begin
  // set default value
  Result := '';
  // get filename of exe/dll if no filename is specified
  sFileName := Trim(FileName);
  if (sFileName = '') then
    sFileName := GetModuleName(HInstance);
  // get size of version info (0 if no version info exists)
  iBufferSize := GetFileVersionInfoSize(PChar(sFileName), iDummy);
  if (iBufferSize > 0) then
  begin
    GetMem(pBuffer, iBufferSize);
    try
    // get fixed file info (language independent)
    GetFileVersionInfo(PChar(sFileName), 0, iBufferSize, pBuffer);
    VerQueryValue(pBuffer, '\', pFileInfo, iDummy);
    // read version blocks
    iVer[1] := HiWord(PVSFixedFileInfo(pFileInfo)^.dwFileVersionMS);
    iVer[2] := LoWord(PVSFixedFileInfo(pFileInfo)^.dwFileVersionMS);
    iVer[3] := HiWord(PVSFixedFileInfo(pFileInfo)^.dwFileVersionLS);
    iVer[4] := LoWord(PVSFixedFileInfo(pFileInfo)^.dwFileVersionLS);
    finally
      FreeMem(pBuffer);
    end;
    // format result string
    Result := Format(Fmt, [iVer[1], iVer[2], iVer[3], iVer[4]]);
  end;
end;

end.
