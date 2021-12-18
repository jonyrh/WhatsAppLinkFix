program WhatsAppLinkFix;

{$mode objfpc}{$H+}

uses
  Classes, SysUtils, CustApp, LazUTF8, StrUtils, ShlObj, ComObj, ActiveX;

type

  { TWhatsAppLinkFix }

  TWhatsAppLinkFix = class(TCustomApplication)
  protected
    procedure DoRun; override;
  public
    constructor Create(TheOwner: TComponent); override;
    destructor Destroy; override;
    procedure ScanDir(StartDir: String; List: TStringList);
    procedure CreateLink(const aSource, aLinkPath: String);
  end;

{ TWhatsAppLinkFix }

procedure TWhatsAppLinkFix.CreateLink(const aSource, aLinkPath: String);
var
  IObject: IUnknown;
  ISLink: IShellLink;
  IPFile: IPersistFile;
begin
  IObject := CreateComObject(CLSID_ShellLink);
  ISLink := IObject as IShellLink;
  IPFile := IObject as IPersistFile;
  ISLink.SetPath(PChar(aSource));
  ISLink.SetWorkingDirectory(PChar(ExtractFilePath(aSource)));
  IPFile.Save(PWideChar(WideString(aLinkPath + '\WhatsApp.lnk')), False);
end;

procedure TWhatsAppLinkFix.ScanDir(StartDir: String; List: TStringList);
var
SearchRec: TSearchRec;
begin
if StartDir[Length(StartDir)]<>'\' then StartDir:= StartDir + '\';

if FindFirst(StartDir + '*.*', faAnyFile, SearchRec)=0 then
 begin
  repeat
  if (SearchRec.Attr and faDirectory) <> faDirectory then List.Add(StartDir + SearchRec.Name)
   else if (SearchRec.Name <> '..') and (SearchRec.Name <> '.')then
    begin
    List.Add(StartDir + SearchRec.Name + '\');
    ScanDir(StartDir +  SearchRec.Name + '\', List);
    end;
  until FindNext(SearchRec) <> 0;
  FindClose(SearchRec);
  end;
end;

procedure TWhatsAppLinkFix.DoRun;
var
  aDir,
  aLinkPath,
  aAppPath: String;
  aSL: TStringList;
  i: Integer;
begin
  aAppPath:='';
  aDir:=      GetEnvironmentVariable('SystemDrive') + '\Program Files\WindowsApps';
  aLinkPath:= GetEnvironmentVariable('AppData')     + '\Microsoft\Windows\Start Menu\Programs\WhatsApp';

  aSL:= TStringList.Create;
  aSL.Clear;

  ScanDir(aDir, aSL);

  for i:=0 to aSL.Count-1 do
   if ContainsText(aSL[i], 'WhatsApp.exe') then aAppPath:=aSL[i];

  aSL.Free;

  if not aAppPath.Trim.IsEmpty then
   begin
   if not DirectoryExists(aAppPath) then ForceDirectories(aAppPath);
   CreateLink(aAppPath, aLinkPath);
   end;

  Terminate;
end;

constructor TWhatsAppLinkFix.Create(TheOwner: TComponent);
begin
  inherited Create(TheOwner);
  StopOnException:=True;
end;

destructor TWhatsAppLinkFix.Destroy;
begin
  inherited Destroy;
end;

var
  Application: TWhatsAppLinkFix;

{$R *.res}

begin
  Application:=TWhatsAppLinkFix.Create(nil);
  Application.Title:='WhatsApp Link Fix';
  Application.Run;
  Application.Free;
end.

