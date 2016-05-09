unit about;

interface

uses Windows, Classes, Graphics, Forms, Controls, StdCtrls,
  Buttons, ExtCtrls,  jpeg;

type
  TAboutBox = class(TForm)
    Panel1: TPanel;
    OKButton: TButton;
    Comments: TLabel;
    Image1: TImage;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  AboutBox: TAboutBox;

implementation

{$R *.dfm}

end.
 
