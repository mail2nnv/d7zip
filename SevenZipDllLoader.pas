{------------------------------------------------------------------------------}
{ ModuleName: SevenZipDllLoader.pas                                            }
{ Author: NNV                                                                  }
{ DateTime: 20.11.2015 14:15:42                                                }
{                                                                              }
{ Description:                                                                 }
{   Линковка библиотеки 7z.dll как ресурса в программу. Извлечение его на RT   }
{   в указанную папку                                                          }
{------------------------------------------------------------------------------}
{ History:                                                                     }
{   (!) 20.11.2015 14:15:42 - Created                                          }
{------------------------------------------------------------------------------}

unit SevenZipDllLoader;

interface

const
  SResName = '7zip_library';
  SLibName = '7z.dll';
  
{$R 7z.res}

{ Extract7zLibFile: Проверить наличие в указанной папке 7z.dll. 
  Если ее там нет, извлечь из ресурса и сохранить ее в этой папке }
function Extract7zLibFile(const DstDir: string = ''): Boolean;

{ Load7zLibFromResource: Загрузить библиотеку 7z.dll в память процесса.

  Если библиотека 7z.dll уже загружена, то просто увеличит счетчик ее 
  использования, как при повторном вызове LoadLibrary.
  Если библиотека 7z.dll еще не загружена, то извлечет ее из ресурса во 
  временный каталог и загрузит ее в память процесса вызовом LoadLibrary. 

  Возвращает в результате handle загруженной библиотеки, в необязательном 
  параметре ALibPath путь к файлу библиотеки.

  Каждому вызову Load7zLibFromResource должен соответствовать свой 
  вызов FreeLibrary.

  По выходу из приложения  временный каталог удаляется автоматически. }
function Load7zLibFromResource(var ALibPath: string): Integer; overload;
function Load7zLibFromResource: Integer; overload;

implementation

uses
  Windows,
  Classes,
  SysUtils,
  SysUtilsExt;

function Extract7zLibFile(const DstDir: string): Boolean;
var
  MemStream: TResourceStream;
  S: string;
begin
  Result := True;
  
  S := DstDir;
  if S = '' then
    S := ExtractFilePath(ParamStr(0));

  S := AddBk(S) + SLibName;

  if FileExists(S) then
    Exit;

  MemStream := TResourceStream.Create( HInstance, SResName, RT_RCDATA);
  try
    MemStream.SaveToFile(S);
  finally
    MemStream.Free;
  end;

  Result := FileExists(S);
end;

var
  __TempLibDir: string = '';
  __Lib: HMODULE = 0;

procedure DeleteTempLibDir; far;
begin
  if __TempLibDir <> '' then
  begin
    if __Lib <> 0 then
    begin
      FreeLibrary(__Lib);
      __Lib := 0;
      SafeDeleteFile(__TempLibDir + SLibName);
      DeleteTempDirectory(__TempLibDir);
    end;
    __TempLibDir := '';
  end;
end;

function Load7zLibFromResource(var ALibPath: string): Integer;
var
  Path: packed array [0..MAX_PATH] of Char;
begin
  Result := GetModuleHandle(SLibName);
  if Result = 0 then
  begin
    if __TempLibDir = '' then
      __TempLibDir := CreateTempDirectory(True);

    if Extract7zLibFile(__TempLibDir) then
    begin
      AddExitProc(DeleteTempLibDir);
      ALibPath := __TempLibDir + SLibName;
      __Lib :=  LoadLibrary(PChar(ALibPath));
      Result := __Lib;
    end;
  end
  else
  begin
    Result := LoadLibrary(SLibName); // увеличиваем счетчик использования
    GetModuleFileName(Result, @Path, MAX_PATH);
    ALibPath := StrPas(@Path);
  end;
end;

function Load7zLibFromResource: Integer; 
var
  TempLibPath: string;
begin
  Result := Load7zLibFromResource(TempLibPath);
end;

end.