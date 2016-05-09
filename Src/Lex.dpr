program Lex;

uses
  Forms,
  Main in 'Main.pas' {MainForm},
  CHILDWIN in 'CHILDWIN.PAS' {MDIChild},
  about in 'about.pas' {AboutBox},
  uYMOFXReader in 'uYMOFXReader.pas',
  uSplash in 'uSplash.pas' {fSplash};

{$R *.RES}

begin
  Application.Initialize;
  Application.Title := 'Lex';
  Application.CreateForm(TMainForm, MainForm);
  Application.CreateForm(TfSplash, fSplash);
  Application.CreateForm(TAboutBox, AboutBox);
  Application.Run;
end.
