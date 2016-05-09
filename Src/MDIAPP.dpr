program Mdiapp;

uses
  Forms,
  Main in 'Main.pas' {MainForm},
  CHILDWIN in 'CHILDWIN.PAS' {MDIChild},
  about in 'about.pas' {AboutBox},
  uYMOFXReader in 'uYMOFXReader.pas';

{$R *.RES}

begin
  Application.Initialize;
  Application.Title := 'Leitor OFX';
  Application.CreateForm(TMainForm, MainForm);
  Application.CreateForm(TAboutBox, AboutBox);
  Application.Run;
end.
