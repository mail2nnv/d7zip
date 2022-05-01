{------------------------------------------------------------------------------}
{ ModuleName: SevenZipVCL_Test.pas                                             }
{ Author: NNV                                                                  }
{ DateTime: 28.04.2022 16:08:37                                                }
{                                                                              }
{ Description:                                                                 }
{   Тесты для SevenZipVCL                                                      }
{------------------------------------------------------------------------------}
{ History:                                                                     }
{   (!) 28.04.2022 16:08:37 - Created                                          }
{------------------------------------------------------------------------------}

unit SevenZipVCL_Test;

interface

uses
  Windows,
  Classes, 
  SysUtils,

  TestFrameWork, 
  TestRandom;

type
  T7z_Test = class(TIteratedTestCase)
  protected
    procedure Setup; override;
    procedure TearDown; override;
  published
    procedure Test_00_BasicUsage;
    procedure Test_01_PasswordAndProgress;
  end;


implementation

uses
  SysUtilsExt,
  SevenZipVCL,
  ZipThreads,
  FormProgress;

type
  TDirStructParams = packed record
    Files,    // файлов в одном каталоге
    FileSize, // размер файла
    SubDirs,  // подкаталогов в каталоге
    Depth:    // вложенность
      packed record
        Min, Max: Integer;
      end;
    MaxTime: Int64; // ьаксимальное время генерации
  end;

  TDirStructStat = packed record
    Files,    // сколько файлов
    SubDirs,  // сколько подкаталогов
    FileSize: // общий размер
      Integer;
  end;

procedure GenerateDirStruct(
  const ATest: TAbstractTest;       // тест
  const ARootDir: string;           // корневой каталог
  const APars: TDirStructParams;    // параметры
  const AFileList: TStrings;        // список файлов
  var AStat: TDirStructStat);       // статистика

  var 
    Depth, MaxDepth: Integer;
    Start: Int64;
    BinFiles: IStringList;

  function GetBinFiles: IStringList;
  begin
    if BinFiles = nil then
    begin
      BinFiles := MakeStringList();
      GetFileList(ExtractFilePath(ParamStr(0)), '*.exe;*.dll', 
        BinFiles.AsStrings, False);
    end;
    Result := BinFiles;
  end;

  function TimeExpired: Boolean;
  var
    NowTime, Freq, TimeMs: Int64;
  begin
    Result := False;
    QueryPerformanceCounter(NowTime);

    if QueryPerformanceFrequency(Freq) then
    begin
      TimeMs := 1000 * (NowTime - Start) div Freq;
      Result := TimeMs > APars.MaxTime;
    end;
  end;

  procedure DoGenerate(const ADir: string);

    function RandomFileName: string;
    const
      CNameChars = CEngChars + CDigChars + CRusChars;
    begin
      Result := 
        RandomChar(CNameChars) +
        RandomStr(0, 14, CNameChars + [' ', '.']) + 
        RandomChar(CNameChars);
    end;

    function CreateRandomFile: string;

      procedure StreamFromStream(const ADst, ASrc: TStream; const ASize: Integer);
      var
        I, SrcSize: Integer;
      begin
        SrcSize := ASrc.Size;
        for I := 0 to Pred(ASize div SrcSize) do
        begin
          ASrc.Seek(0, 0);
          ADst.CopyFrom(ASrc, SrcSize);
        end;
        I := ASize mod SrcSize;
        if I > 0 then
        begin
          ASrc.Seek(0, 0);
          ADst.CopyFrom(ASrc, I);
        end;
      end;
    
      procedure SetFileRandomTime;
      var
        Handle: integer;
        LocalFileTime, FileTime: TFileTime;
        Age: integer;
      begin
        Age := DateTimeToFileDate(RandomDateTime(Now - 5 * 365.0, Now));
        Handle := SysUtils.FileOpen(Result, fmOpenWrite);
        if Handle <> -1 then
          try
            if DosDateTimeToFileTime(LongRec(Age).Hi, LongRec(Age).Lo, LocalFileTime) then
              if LocalFileTimeToFileTime(LocalFileTime, FileTime) then
                Windows.SetFileTime(Handle, @FileTime, nil, @FileTime);
          finally
            FileClose(Handle);
          end;
      end;

    const
      FileExts: packed array [1..3] of string = ('.TXT', '.BIN', '.ARC');
    var
      BinF: string;
      Buf: array [0..1024] of Char;
      Fs: TStream;
      I: Integer;
      FileSize: Integer;
      FileType: Integer;
      SrcStream: TStream;
    begin
      FileType := RandomInteger(1, 3);
      repeat
        Result := AddBk(ADir) + RandomFileName + FileExts[FileType];
      until not FileExists(Result);

      ATest.CheckTrue(IsChildPath(Result, ADir));

      FileSize := RandomInteger(APars.FileSize.Min, APars.FileSize.Max);

      Fs := TFileStream.Create(Result, fmCreate);
      try
        case FileType of
          1:
          begin
            // best compression text data
            SrcStream := TResourceStream.Create(hInstance,
              'Through_the_Looking_Glass', 'TXT');
            try
              StreamFromStream(Fs, SrcStream, FileSize);
            finally
              FreeAndNil(SrcStream);
            end;
          end;
          2:
          begin
            // middle compression text data
            BinF := GetBinFiles[RandomInteger(0, Pred(GetBinFiles.Count))];
            SrcStream := TFileStream.Create(BinF, fmOpenRead or fmShareDenyWrite);
            try
              StreamFromStream(Fs, SrcStream, FileSize);
            finally
              FreeAndNil(SrcStream);
            end;
          end;
        else {3: }
          begin
            // poor compression random data
            for I := 0 to Pred(FileSize div 1024) do
            begin
              RandomBuf(@Buf, 1024);
              Fs.Write(Buf, 1024);
            end;
            I := FileSize mod 1024;
            if I > 0 then
            begin
              RandomBuf(@Buf, I);
              Fs.Write(Buf, I);
            end;
          end;
        end;
      finally
        FreeAndNil(Fs);
      end;

      ATest.CheckTrue(FileExists(Result));
      ATest.CheckEquals(GetSizeFile(Result), FileSize);

      SetFileRandomTime;

      AFileList.Add(Result);

      Inc(AStat.Files);
      Inc(AStat.FileSize, FileSize);
    end;

    function CreateRandomDir: string;
    begin
      repeat
        Result := AddBk(ADir) + RandomFileName;
      until not FileExists(Result) and not DirExists(Result);

      Inc(AStat.SubDirs);

      ATest.CheckTrue(IsChildPath(Result, ADir));

      ATest.CheckTrue(CreateDirs(Result));

      DoGenerate(Result);
    end;

  var
    DirCnt: Integer;
    FileCnt: Integer;
    I: Integer;
  begin
    FileCnt := RandomInteger(APars.Files.Min, APars.Files.Max);
    for I := 1 to FileCnt do
    begin
      CreateRandomFile;
      if TimeExpired then
        Exit;
    end;

    if Depth < MaxDepth then
    begin
      Inc(Depth);
      try
        DirCnt := RandomInteger(APars.SubDirs.Min, APars.SubDirs.Max);
        for I := 1 to DirCnt do
        begin
          CreateRandomDir;
          if TimeExpired then
            Exit;
        end;
      finally
        Dec(Depth);
      end;
    end;
  end;

begin
  FillChar(AStat, SizeOf(AStat), #0);
  MaxDepth := RandomInteger(APars.Depth.Min, APars.Depth.Max);
  Depth := 0;
  BinFiles := nil;
  QueryPerformanceCounter(Start);
  DoGenerate(ARootDir);
end;


{ T7z_Test }

procedure T7z_Test.Setup;
begin
  inherited;
end;

procedure T7z_Test.TearDown;
begin
  inherited;
end;

procedure T7z_Test.Test_00_BasicUsage;
var
  ArchName: string;
  DstD: string;
  DstFileList: IStringList;
  DstFN: string;
  Err: string;
  NowTime: Int64;
  Freq: Int64;
  I: Integer;
  Ok: Boolean;
  Pars: TDirStructParams;
  SrcD: string;
  SrcFileList: IStringList;
  SrcFN: string;
  StartTime: Int64;
  Stat: TDirStructStat;
  TimeMs: Int64;
begin
  SrcD := CreateTempDirectory(True);
  try
    SrcFileList := MakeStringList(dupIgnore);

    Pars.Files.Min := 0;
    Pars.Files.Max := 16;
    Pars.FileSize.Min := 0;
    Pars.FileSize.Max := 2 * 1024 * 1024;
    Pars.SubDirs.Min := 0;
    Pars.SubDirs.Max := 2;
    Pars.Depth.Min := 1;
    Pars.Depth.Max := 4;
    Pars.MaxTime := 30 * 1000;
    
//    model test data: 
//    Pars.Files.Min := 16;
//    Pars.Files.Max := 16;
//    Pars.FileSize.Min := 1 * 1024 * 1024;
//    Pars.FileSize.Max := 1 * 1024 * 1024;
//    Pars.SubDirs.Min := 2;
//    Pars.SubDirs.Max := 2;
//    Pars.Depth.Min := 4;
//    Pars.Depth.Max := 4;
//    Pars.MaxTime := 30 * 1000;
//    files: 496, dirs: 30, size: 496.0 MB

//    # model test results:

//    ## SevenZipVcl (7za.dll):
//    compress time: | ratio: | saved:  | speed:    || decompress time: | speed: 
//    ---------------+--------+---------+-----------++------------------+-----------
//    142628 ms      | 2.71   | 63.03%  | 3646 kbps || 16066 ms         | 32372 kbps
//    148870 ms      | 2.59   | 61.39%  | 3493 kbps || 16621 ms         | 31291 kbps
//    154958 ms      | 2.91   | 65.67%  | 3356 kbps || 15034 ms         | 34594 kbps

//    ## sevenzip (7z.dll):
//    compress time: | ratio: | saved:  | speed:    || decompress time: | speed: 
//    ---------------+--------+---------+-----------++------------------+-----------
//    101008 ms      | 2.72   | 63.18%  | 5149 kbps || 12997 ms         | 40016 kbps
//    105985 ms      | 2.76   | 63.73%  | 4907 kbps || 12429 ms         | 41845 kbps
//    103518 ms      | 2.70   | 62.91%  | 5024 kbps || 12724 ms         | 40875 kbps
//    122404 ms      | 3.15   | 68.21%  | 4248 kbps || 12877 ms         | 40389 kbps
//    118105 ms      | 2.80   | 64.34%  | 4403 kbps || 12763 ms         | 40750 kbps

    GenerateDirStruct(Self, SrcD, Pars, SrcFileList.AsStrings, Stat);

    Status('files: %d, dirs: %d, size: %s', [Stat.Files, Stat.SubDirs, 
      SaySize(Stat.FileSize)]);


    ArchName := TempFileName('arc.7z');
    if FileExists(ArchName) then
      SafeDeleteFile(ArchName);
    CheckFalse(FileExists(ArchName));

    QueryPerformanceCounter(StartTime);

    Ok := CreateZipArchive(ArchName, SrcD, SrcD, Err, True, False);

    CheckTrue(Ok);

    QueryPerformanceCounter(NowTime);
    if QueryPerformanceFrequency(Freq) then
    begin
      TimeMs := 1000 * (NowTime - StartTime) div Freq;
      I := GetSizeFile(ArchName) + 1;
      Status('compress time: %d ms, ratio: %s (saved: %s%%), speed: %d kbps', [
        TimeMs,
        FloatToStrF(Stat.FileSize / I, ffFixed, 16, 2),
        FloatToStrF(100 * (1 - I / Stat.FileSize), ffFixed, 16, 2),
        Stat.FileSize div TimeMs]);
    end;

    DstD := CreateTempDirectory(True);
    try
      QueryPerformanceCounter(StartTime);

      Ok := UnzipFile2(ArchName, DstD, '', nil, Err);

      CheckTrue(Ok);

      QueryPerformanceCounter(NowTime);
      if QueryPerformanceFrequency(Freq) then
      begin
        TimeMs := 1000 * (NowTime - StartTime) div Freq;
        Status('decompress time: %d ms, speed: %d kbps', [TimeMs,
          Stat.FileSize div  TimeMs]);
      end;

      DstFileList := MakeStringList(dupIgnore);
      GetFileList(DstD, '', DstFileList.AsStrings, True);

      CheckEquals(SrcFileList.Count, DstFileList.Count);

      for I := 0 to Pred(SrcFileList.Count) do
      begin
        SrcFN := SrcFileList[I];
        DstFN := AddBk(DstD) + System.Copy(SrcFN, Succ(Length(SrcD)), MAX_PATH);

        CheckTrue(FileExists(DstFN));

        CheckTrue(
          BinFileComp(SrcFN, DstFN) = 0
        );

        CheckEquals(GetDateTimeFile(SrcFN), GetDateTimeFile(DstFN));
      end;
    finally
      DeleteTempDirectory(DstD);
    end;
  finally
    DeleteTempDirectory(SrcD);
  end;
end;

type
  T7zProgress = class(TInterfacedObject, IUnknown)
  private
    FArch: TSevenZip;
    FMaxProgress: Int64;
    FProgress: IFormProgress;
    procedure PreProgress(Sender: TObject; MaxProgress: Int64);
    procedure Progress(Sender: TObject; FileName: Widestring; 
      FilePosArc, FilePosFile: Int64);
  public
    constructor Create(const AArch: TSevenZip);
    destructor Destroy; override;
  end;

procedure T7z_Test.Test_01_PasswordAndProgress;
const
  SPassword = 'sigma-$oft';
var
  ArchName: string;
  DstD: string;
  DstFileList: IStringList;
  DstFN: string;
  NowTime: Int64;
  Freq: Int64;
  I: Integer;
  Ok: Boolean;
  Pars: TDirStructParams;
  Progress: IUnknown;
  SrcD: string;
  SrcFileList: IStringList;
  SrcFN: string;
  StartTime: Int64;
  Stat: TDirStructStat;
  TimeMs: Int64;
  Unzip7: TSevenZip;
  Zip7: TSevenZip;
begin
  SrcD := CreateTempDirectory(True);
  try
    SrcFileList := MakeStringList(dupIgnore);

    Pars.Files.Min := 0;
    Pars.Files.Max := 16;
    Pars.FileSize.Min := 0;
    Pars.FileSize.Max := 2 * 1024 * 1024;
    Pars.SubDirs.Min := 0;
    Pars.SubDirs.Max := 2;
    Pars.Depth.Min := 1;
    Pars.Depth.Max := 4;
    Pars.MaxTime := 30 * 1000;
    
    GenerateDirStruct(Self, SrcD, Pars, SrcFileList.AsStrings, Stat);

    Status('files: %d, dirs: %d, size: %s', [Stat.Files, Stat.SubDirs, 
      SaySize(Stat.FileSize)]);


    ArchName := TempFileName('arc.7z');
    if FileExists(ArchName) then
      SafeDeleteFile(ArchName);
    CheckFalse(FileExists(ArchName));

    QueryPerformanceCounter(StartTime);

    Zip7 := TSevenZip.Create(nil);
    try
      Zip7.AddRootDir := SrcD;
      Zip7.AddOptions := [AddSolid, AddRecurseDirs];
      Zip7.LZMACompressStrength := ULTRA;
      Zip7.SZFileName := ArchName;
      Zip7.Password := SPassword;

      Progress := T7zProgress.Create(Zip7);
      
      Ok := (Zip7.Add >= 0) and FileExists(ArchName);

      CheckTrue(Ok);

      Progress := nil;

    finally
      FreeAndNil(Zip7);
    end;
    
    QueryPerformanceCounter(NowTime);
    if QueryPerformanceFrequency(Freq) then
    begin
      TimeMs := 1000 * (NowTime - StartTime) div Freq;
      I := GetSizeFile(ArchName) + 1;
      Status('compress time: %d ms, ratio: %s (saved: %s%%), speed: %d kbps', [
        TimeMs,
        FloatToStrF(Stat.FileSize / I, ffFixed, 16, 2),
        FloatToStrF(100 * (1 - I / Stat.FileSize), ffFixed, 16, 2),
        Stat.FileSize div TimeMs]);
    end;

    DstD := CreateTempDirectory(True);
    try
      QueryPerformanceCounter(StartTime);

      Unzip7 := TSevenZip.Create(nil);
      try
        Unzip7.SZFileName := ArchName;
        Unzip7.ExtrBaseDir := DstD;
        Unzip7.ExtractOptions := [ExtractOverwrite];
        Unzip7.Password := SPassword;

        Progress := T7zProgress.Create(Unzip7);

        Unzip7.Extract;

        Progress := nil;
      finally
        Unzip7.Free;
      end;

      QueryPerformanceCounter(NowTime);
      if QueryPerformanceFrequency(Freq) then
      begin
        TimeMs := 1000 * (NowTime - StartTime) div Freq;
        Status('decompress time: %d ms, speed: %d kbps', [TimeMs,
          Stat.FileSize div  TimeMs]);
      end;

      DstFileList := MakeStringList(dupIgnore);
      GetFileList(DstD, '', DstFileList.AsStrings, True);

      CheckEquals(SrcFileList.Count, DstFileList.Count);

      for I := 0 to Pred(SrcFileList.Count) do
      begin
        SrcFN := SrcFileList[I];
        DstFN := AddBk(DstD) + System.Copy(SrcFN, Succ(Length(SrcD)), MAX_PATH);

        CheckTrue(FileExists(DstFN));

        CheckTrue(
          BinFileComp(SrcFN, DstFN) = 0
        );

        CheckEquals(GetDateTimeFile(SrcFN), GetDateTimeFile(DstFN));
      end;
    finally
      DeleteTempDirectory(DstD);
    end;
  finally
    DeleteTempDirectory(SrcD);
  end;
end;

{ T7zProgress }

constructor T7zProgress.Create(const AArch: TSevenZip);
begin
  inherited Create;
  FArch := AArch;
  FArch.OnPreProgress := PreProgress;
  FArch.OnProgress := Progress;
end;

destructor T7zProgress.Destroy;
begin
  FArch.OnProgress := nil;
  FArch.OnPreProgress := nil;
  FProgress := nil;
  inherited;
end;

procedure T7zProgress.PreProgress(Sender: TObject; MaxProgress: Int64);
begin
  FMaxProgress := MaxProgress;
  FProgress := MakeViewProgress('');
end;

procedure T7zProgress.Progress(Sender: TObject; FileName: Widestring;
  FilePosArc, FilePosFile: Int64);
begin
  if Assigned(FProgress) then
  begin
    FProgress.Play(FilePosArc, FMaxProgress);
    if FileName <> '' then
      FProgress.Info(FileName);
  end;
end;

initialization
  TestFramework.RegisterTests('Проверка SevenZipVCL', [T7z_Test.Suite]);
end.