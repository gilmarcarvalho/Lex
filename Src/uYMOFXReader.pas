unit uYMOFXReader;

interface

uses classes, SysUtils, Controls,Dialogs;

type
  TOFXItem = class
    MovType : string;
    MovDate : TDate;
    Value : double;
    ID : string;
    Document : string;
    Desc : string;
  end;

  TYMOFXReader = class(TComponent)
  public
    BankID : integer;
    BankName : string;
    AccountID : string;
    AccountType : string;
    InitialBalance : double;
    FinalBalance : double;
    DateStart:TDate;
    DateEnd:TDate;
    constructor Create( AOwner: TComponent ); override;
    destructor Destroy; override;
    function Process: boolean;
    function Get(iIndex: integer): TOFXItem;
    function Count: integer;
  private
    FOFXFile : string;
    FListItems : TList;
    procedure Clear;
    procedure Delete( iIndex: integer );
    function Add: TOFXItem;
    function PrepareFloat( sString : string ) : string;
    function InfLine ( sLine : string;sStart:string; sEnd:string ): string;
    function FindString ( sSubString, sString : string ): boolean;
    function ReplaceString(sString: string; sOld: string; sNew: string; bInsensitive : boolean = true): string;
    function ExtractDate(sDate: string):TDate;
  protected
  published
    property OFXFile: string read FOFXFile write FOFXFile;
  end;



implementation

constructor TYMOFXReader.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FListItems := TList.Create;
end;

destructor TYMOFXReader.Destroy;
begin
  FListItems.Free;
  inherited Destroy;
end;

procedure TYMOFXReader.Delete( iIndex: integer );
begin
  TOFXItem(FListItems.Items[iIndex]).Free;
  FListItems.Delete( iIndex );
end;

procedure TYMOFXReader.Clear;
//var
//  i: integer;
//  oPointer : Pointer;
begin
  while FListItems.Count > 0 do
    Delete(0);
  FListItems.Clear;
end;

function TYMOFXReader.Count: integer;
begin
  Result := FListItems.Count;
end;

function TYMOFXReader.Get(iIndex: integer): TOFXItem;
begin
  Result := TOFXItem(FListItems.Items[iIndex]);
end;

function TYMOFXReader.Process: boolean;
var
  oFile : TStringList;
  i : integer;
  bOFX : boolean;
  oItem : TOFXItem;
  sLine : string;
  dBalance : double;
begin
  Result := false;
  Clear;
  bOFX := false;
  if not FileExists(FOFXFile) then
    exit;
  oFile := TStringList.Create;
  oFile.LoadFromFile(FOFXFile);
  dBalance := 0;
  i := 0;
  while i < oFile.Count do
  begin
    sLine := oFile.Strings[i];
    if FindString('<OFX>', sLine) then
      bOFX := true;
    if bOFX then
    begin
      // -----------------------------------------------------------------------
      // BankID
      if FindString('<BANKID>', sLine) then BankID := StrToIntDef(InfLine(sLine,'<BANKID>','</BANKID>'), 0);
      // -----------------------------------------------------------------------
      // AccountID
      if FindString('<ACCTID>', sLine) then AccountID := InfLine(sLine,'<ACCTID>','</ACCTID>');
      // -----------------------------------------------------------------------
      // AccountType
      if FindString('<ACCTTYPE>', sLine) then AccountType := InfLine(sLine,'<ACCTYPE>','</ACCTYPE>');
      // -----------------------------------------------------------------------
      // FinalBalance
      if FindString('<LEDGER>', sLine)  then
        FinalBalance := StrToFloat(PrepareFloat(InfLine(sLine,'<LEDGER>','</LEDGER>')));
      if FindString('<BALANMT>', sLine)  then
        FinalBalance := StrToFloat(PrepareFloat(InfLine(sLine,'<BALANMT>','</BALANMT>')));
      if FindString('<BALAMT>', sLine)  then
        FinalBalance := StrToFloat(PrepareFloat(InfLine(sLine,'<BALAMT>','</BALAMT>')));
      if FindString('<ORG>', sLine)  then
        BankName := InfLine(sLine,'<ORG>','</ORG>');
      if FindString('<DTSTART>', sLine)  then
        DateStart := ExtractDate(infLine(sLine,'<DTSTART>','</DTSTART>'));
      if FindString('<DTEND>', sLine)  then
        DateEnd := ExtractDate(infLine(sLine,'<DTEND>','</DTEND>'));

      // -----------------------------------------------------------------------
      // Moviment
      if FindString('<STMTTRN>', sLine) then
      begin
        oItem := Add;
        Inc(i);
        sLine := oFile.Strings[i];
        if FindString('<TRNTYPE>',  sLine) then oItem.MovType := InfLine(sLine,'<TRNTYPE>','</TRNTYPE>');
        Inc(i);
        sLine := oFile.Strings[i];
        if FindString('<DTPOSTED>', sLine) then begin
          oItem.MovDate := ExtractDate(InfLine(sLine,'<DTPOSTED>','</DTPOSTED>'));
        end;
        Inc(i);
        sLine := oFile.Strings[i];
        if FindString('<TRNAMT>',   sLine) then
        begin
          oItem.Value := StrToFloat(PrepareFloat(InfLine(sLine,'<TRNAMT>','</TRNAMT>')));
          dBalance := dBalance - oItem.Value;
        end;
        Inc(i);
        sLine := oFile.Strings[i];
        if FindString('<FITID>',    sLine) then oItem.ID := InfLine(sLine, '<FITID>','</FITID>');
        Inc(i);
        sLine := oFile.Strings[i];
        if FindString('<CHECKNUM>',   sLine) then begin
          oItem.Document := InfLine(sLine, '<CHECKNUM>','</CHECKNUM>');
          inc(i);
        end;
        sLine := oFile.Strings[i];
        if FindString('<REFNUM>', sLine) then begin
          oItem.Document := InfLine(sLine,'<REFNUM>','</REFNUM>');
          Inc(i);
        end;
        sLine := oFile.Strings[i];
        if FindString('<PAYEEID>',   sLine) then //pula
        Inc(i);
        sLine := oFile.Strings[i];
        if FindString('<MEMO>',     sLine) then oItem.Desc := InfLine(sLine,'<MEMO>','</MEMO>');
//        end;
      end;
      // -----------------------------------------------------------------------
    end;
    Inc(i);
  end;
  InitialBalance := FinalBalance + dBalance;
  Result := bOFX;
end;

function TYMOFXReader.PrepareFloat( sString : string ) : string;
var d:TFormatSettings;c:char;
begin
  d:=TFormatSettings.Create;
  c:= d.DecimalSeparator;
  Result := sString;
  Result := ReplaceString(Result, '.', c);
  Result := ReplaceString(Result, ',', c);
end;
function TYMOFXReader.ExtractDate(sDate: string):TDate;
var Res:TDate;
begin
 res:= EncodeDate(StrToIntDef(copy(sDate,1,4), 0),
                  StrToIntDef(copy(sDate,5,2), 0),
                  StrToIntDef(copy(sDate,7,2), 0));
 Result:=res;
end;


function TYMOFXReader.ReplaceString(sString: string; sOld: string; sNew: string; bInsensitive : boolean = true): string;
var
   iPosition: integer ;
   sTempNew: string ;
begin
   iPosition := 1;
   sTempNew := '';
   while (iPosition > 0) do
   begin
      if bInsensitive then
        iPosition := AnsiPos(UpperCase(sOld),UpperCase(sString))
      else
        iPosition := AnsiPos(sOld,sString);
      if (iPosition > 0) then
      begin
         sTempNew := sTempNew + copy(sString, 1, iPosition - 1) + sNew;
         sString := copy(sString, iPosition + Length(sOld), Length(sString) );
      end;
   end;
   sTempNew := sTempNew + sString;
   Result := (sTempNew);
end;

function TYMOFXReader.InfLine ( sLine : string;sStart:string;sEnd:string ): string;
var
  iTemp : integer;
  iStart,iEnd:integer;
begin
  sLine := Trim(sLine);
  if findstring(sStart,sLine) then begin
    iStart:=pos(sStart,sLine)+length(sStart);
    iEnd:=Pos(sEnd,sLine);
    if iEnd=0 then   iEnd:=length(sLine)+1;
    sLine:=copy(sline,iStart,abs(iEnd-iStart));
    Result := sLine;
    end else begin;
    Result := '';
  end;
 end;

function TYMOFXReader.Add: TOFXItem;
var
  oItem : TOFXItem;
begin
  oItem := TOFXItem.Create;
  FListItems.Add(oItem);
  Result := oItem;
end;

function TYMOFXReader.FindString ( sSubString, sString : string ): boolean;
begin
  Result := Pos(UpperCase(sSubString), UpperCase(sString)) > 0;
end;


end.
