{------------------------------------------------------------------------------}
{ ModuleName: SevenZipVCL.pas                                                  }
{ Author: NNV                                                                  }
{ DateTime: 15.04.2022 16:46:25                                                }
{                                                                              }
{ Description:                                                                 }
{   }
{------------------------------------------------------------------------------}
{ History:                                                                     }
{   (!) 15.04.2022 16:46:25 - Created                                          }
{------------------------------------------------------------------------------}

unit SevenZipVCL;

interface

uses
  Windows, Classes;

type
  AddOptsEnum = (AddRecurseDirs, AddSolid, AddStoreOnlyFilename, AddIncludeDriveLetter, AddEncryptFilename);
  AddOpts = set of AddOptsEnum;

  ExtractOptsEnum = (ExtractNoPath, ExtractOverwrite);
  ExtractOpts = set of ExtractOptsEnum;

  TCompressStrength = (SAVE, FAST, NORMAL, MAXIMUM, ULTRA);
  
  TWideStringList = class(TStringList)
  public
    procedure AddString(const S: WideString);
  end;

  T7zExtractfileEvent = procedure(Sender: TObject; Filename: Widestring; Filesize: Int64) of object;
  T7zProgressEvent = procedure(Sender: TObject; Filename: Widestring; FilePosArc, FilePosFile: Int64) of object;
  T7zPreProgressEvent = procedure(Sender: TObject; MaxProgress: Int64) of object;
  T7zMessageEvent = procedure(Sender: TObject; ErrCode: Integer; AMessage: string; Filename: Widestring) of object;
  TOpenVolume = procedure(var arcFileName: WideString; Removable: Boolean; out Cancel: Boolean) of object;

  TSevenZip = class(TComponent)
  private
    FAddOptions: AddOpts;
    FCanceled: Boolean;
    FErrCode: Integer;
    FExtractOptions: ExtractOpts;
    FExtrBaseDir: Widestring;
    FFiles: TWideStringList;
    FLastError: Integer;
    FLib: Integer;
    FPassword: WideString;
    FOnExtractfile: T7zextractfileEvent;
    FOnProgress: T7zProgressEvent;
    FOnPreProgress: T7zPreProgressEvent;
    FOnMessage: T7zMessageEvent;
    FOnOpenVolume: TOpenVolume;
    FCompstrength: TCompressStrength;
    FSevenZipFileName: Widestring;
    FRootDir: Widestring;
    procedure SetLastError(const Value: Integer);
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    function Add: Integer;                                                                         {+}
    function Extract: Integer;                                                                     {+}
    procedure Cancel;                                                                              {+}
    property AddOptions: AddOpts read FAddOptions write FAddOptions;                               {+}
    property AddRootDir: Widestring read FRootDir write FRootDir;                                  {+}
    property ErrCode: Integer read fErrCode write fErrCode;                                        {+}
    property ExtractOptions: ExtractOpts read FExtractOptions write FExtractOptions;               {+}
    property ExtrBaseDir: Widestring read FExtrBaseDir write FExtrBaseDir;                         {+}
    property Files: TWideStringList read Ffiles write ffiles;                                      {+}
    property LastError: Integer read FLastError write SetLastError;                                {+}
    property LZMACompressStrength: TCompressStrength read FCompstrength write FCompstrength;       {+}
    property Password: WideString read FPassword write FPassword;                                  {+}
    property SZFileName: Widestring read FSevenZipFileName write FSevenZipFilename;                {+}
    property OnExtractfile: T7zextractfileEvent read FOnextractfile write FOnextractfile;          {+}
    property OnProgress: T7zProgressEvent read FOnProgress  write FOnProgress;                     {+}
    property OnPreProgress: T7zPreProgressEvent read FOnPreProgress  write FOnPreProgress;         {+}
    property OnMessage: T7zMessageEvent read fOnMessage write fOnMessage;                          {+}
    property OnOpenVolume: TOpenVolume read FOnOpenVolume write FOnOpenVolume;                     {+}
  end;

const
// SevenZIP onMessage Errorcode
  FNoError             = 0;
  FFileNotFound        = 1;
  FDataError           = 2;
  FCRCError            = 3;
  FUnsupportedMethod   = 4;
  FIndexOutOfRange     = 5;                                    //FHO 21.01.2007
  FUsercancel          = 6;
  FNoSFXarchive        = 7;
  FSFXModuleError      = 8;
  FSXFileCreationError = 9;                                    //FHO 21.01.2007
  FNoFilesToAdd        =10;                                    //FHO 21.01.2007
  FNoFileCreated       =11;

  c7zipResMsg:array[FNoError..FNoFileCreated] of string=       //FHO 21.01.2007
  { 0}('Success',                                              //FHO 21.01.2007
  { 1} 'File not found',                                       //FHO 21.01.2007
  { 2} 'Data Error',                                           //FHO 21.01.2007
  { 3} 'CRC Error',                                            //FHO 21.01.2007
  { 4} 'Unsupported Method',                                   //FHO 21.01.2007
  { 5} 'Index out of Range',                                   //FHO 21.01.2007
  { 6} 'User canceled operation',                              //FHO 21.01.2007
  { 7} 'File is not an 7z SFX archive',                        //FHO 21.01.2007
  { 8} 'SFXModule error ( Not found )',                        //FHO 21.01.2007
  { 9} 'Could not create SFX',                                 //FHO 21.01.2007
  {10} 'No files to add',                                      //FHO 21.01.2007
  {11} 'Could not create file'                                 //FHO 21.01.2007

       );                                                      //FHO 21.01.2007


{$R 7z.res}

{ Extract7zLibFile: Проверить наличие в указанной папке 7z.dll. 
  Если ее там нет, извлечь из ресурса и сохранить ее в этой папке }
function Extract7zLibFile(const DstDir: string = ''): Boolean;

{ Load7zLib: Загрузить библиотеку 7z.dll в память процесса.

  Если библиотека 7z.dll уже загружена, то просто увеличит счетчик ее 
  использования, как при повторном вызове LoadLibrary.
  
  В противном случае будет искать библиотеку в каталоге исполняемого файла.
  Если найдена в нем, то загрузит ее оттуда.
  
  В противном случае  извлечет ее из ресурса во временный каталог (который 
  автоматически будет удален при выходе из приложения) и загрузит ее 
  в память процесса вызовом LoadLibrary. 

  Возвращает в результате Handle загруженной библиотеки, в необязательном 
  параметре ALibPath путь к файлу библиотеки.

  Каждому вызову Load7zLib должен соответствовать свой 
  вызов FreeLibrary }
function Load7zLib(var ALibPath: string): Integer; overload;
function Load7zLib: Integer; overload;

implementation

uses
  SysUtils,
  SysUtilsExt,
  sevenzip,
  SevenZipDllLoader;

function TSevenZip_ProgressCallback(Sender: Pointer; Total: Boolean; Value: int64): HRESULT; stdcall;
var
  Sz: TSevenZip;
begin
  Result := S_OK; 
  
  Sz := TSevenZip(Sender);
  if Total then
  begin
    if Assigned(Sz.OnPreProgress) then 
      Sz.OnPreProgress(Sz, Value);
  end
  else
    if Assigned(Sz.OnProgress) then 
    begin
      Sz.OnProgress(Sz, '', Value, 0);
      if Sz.FCanceled then
        Result := S_FALSE;
    end;
end;

{ TSevenZip }

function TSevenZip.Add: Integer;

  function CompressionStrengthInt(cs: TCompressStrength): Cardinal;
  begin
    case cs of
      SAVE: result := 0;
      FAST: result := 3;
      NORMAL: result := 5;
      MAXIMUM: result := 7;
      ULTRA: result := 9;
    else
      result := 5;
    end;
  end;

var
  A: I7zOutArchive;
  FN: string;
  FP: string;
  I: Integer;
begin
  A := CreateOutArchive(CLSID_CFormat7z);
  SevenZipSetCompressionMethod(A, m7LZMA);
  SetCompressionLevel(A, CompressionStrengthInt(LZMACompressStrength));
  A.SetProgressCallback(Self, TSevenZip_ProgressCallback);
  if Password <> '' then
    A.SetPassword(Password);

  if Files.Count > 0 then
  begin
    for I := 0 to Pred(Files.Count) do
    begin
      FN := Files[I];
      if AddStoreOnlyFilename in AddOptions then
        FP := ExtractFileName(FN)
      else
        FP := System.Copy(FN, Succ(Length(AddRootDir)), MAX_PATH);
      A.AddFile(FN, FP);
    end;
  end
  else
    A.AddFiles(AddRootDir, '', SAnyFileMask, AddRecurseDirs in AddOptions, True);

  Result := A.BatchSize;

  A.SaveToFile(SZFileName);
end;

procedure TSevenZip.Cancel;
begin
  FCanceled := True;
end;

constructor TSevenZip.Create(AOwner: TComponent);
begin
  FLib := Load7zLib;
  
  inherited;
  FFiles := TWideStringList.Create;
end;

destructor TSevenZip.Destroy;
begin
  if FLib <> 0 then
  begin
    FreeLibrary(FLib);
    FLib := 0;
  end;
  FreeAndNil(FFiles);
  inherited;
end;

function TSevenZip.Extract: Integer;
var
  A: I7zInArchive;
begin
  A := CreateInArchive(CLSID_CFormat7z);
  A.OpenFile(SZFileName);
  try
    A.SetProgressCallback(Self, TSevenZip_ProgressCallback);
    if Password <> '' then
       A.SetPassword(Password);
    A.ExtractTo(ExtrBaseDir);
    Result := A.GetNumberOfItems;
  finally
    A.Close;
  end;
end;

procedure TSevenZip.SetLastError(const Value: Integer);
begin
  FLastError := Value;
end;

{ TWideStringList }

procedure TWideStringList.AddString(const S: WideString);
begin
  Add(S);
end;

//——————————————————————————————————————————————————————————————————————————————

const
  SResName = '7zip_library';
  
function Extract7zLibFile(const DstDir: string): Boolean;
var
  MemStream: TResourceStream;
  S: string;
begin
  Result := True;
  
  S := DstDir;
  if S = '' then
    S := ExtractFilePath(ParamStr(0));

  S := AddBk(S) + C_7zDllName;

  if FileExists(S) then
    Exit;

  MemStream := TResourceStream.Create(HInstance, SResName, RT_RCDATA);
  try
    MemStream.SaveToFile(S);
  finally
    MemStream.Free;
  end;

  Result := FileExists(S);
end;

var
  __TempLibDir: string = '';

procedure DeleteTempLibDir; far;
var
  H: Integer;
  Path: packed array [0..MAX_PATH] of Char;
begin
  if __TempLibDir <> '' then
  begin
    H := GetModuleHandle(C_7zDllName);
    if H <> 0 then
    begin
      GetModuleFileName(H, @Path, MAX_PATH);
      if SamePath(__TempLibDir, ExtractFileDir(StrPas(@Path))) then
        FreeLibrary(H);
    end;

    if SafeDeleteFile(__TempLibDir + C_7zDllName) then
    begin
      DeleteTempDirectory(__TempLibDir);
      __TempLibDir := '';
    end;
  end;
end;

function Load7zLib(var ALibPath: string): Integer;
var
  Path: packed array [0..MAX_PATH] of Char;
begin
  { 1. Ищем среди уже загруженных }
  Result := GetModuleHandle(C_7zDllName);
  if Result <> 0 then
  begin
    Result := SafeLoadLibrary(C_7zDllName); // увеличиваем счетчик использования
    GetModuleFileName(Result, @Path, MAX_PATH);
    ALibPath := StrPas(@Path);
    Exit;
  end;

  { 2. Ищем рядом с исполняемым файлом }
  ALibPath := AddBk(ExtractFileDir(ParamStr(0))) + C_7zDllName;
  if FileExists(ALibPath) then
  begin
    Result := SafeLoadLibrary(PChar(ALibPath));
    Exit;
  end;
  
  { 3. Извлекаем из ресурса во временный кататло и грузим оттуда }
  if __TempLibDir = '' then
  begin
    __TempLibDir := CreateTempDirectory(True);
    AddExitProc(DeleteTempLibDir);
  end;

  if Extract7zLibFile(__TempLibDir) then
  begin
    ALibPath := __TempLibDir + C_7zDllName;
    Result := SafeLoadLibrary(PChar(ALibPath));
  end;
end;

function Load7zLib: Integer; 
var
  TempLibPath: string;
begin
  Result := Load7zLib(TempLibPath);
end;

end.
