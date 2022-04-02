Fragment of main.html  <!-- body { color: #000000; background-color: #FFFFFF; } .pas1-assembler { background-color: #FFFFFF; color: #000000; } .pas1-character { background-color: #FFFFFF; color: #000080; } .pas1-comment { background-color: #FFFFFF; color: #000080; font-style: italic; } .pas1-float { background-color: #FFFFFF; color: #000080; } .pas1-hexadecimal { background-color: #FFFFFF; color: #000080; } .pas1-identifier { background-color: #FFFFFF; color: #000000; } .pas1-number { background-color: #FFFFFF; color: #000080; } .pas1-preprocessor { background-color: #FFFFFF; color: #000080; font-style: italic; } .pas1-reservedword { background-color: #FFFFFF; color: #000000; font-weight: bold; } .pas1-space { background-color: #FFFFFF; color: #000000; } .pas1-string { background-color: #FFFFFF; color: #000080; } .pas1-symbol { background-color: #FFFFFF; color: #000000; } -->

> 7-zip Delphi API
> ================

This API use the 7-zip dll (7z.dll) to read and write all 7-zip supported archive formats.

\- Autor: Henri Gourvest <hgourvest@progdigy.com>  
\- Licence: MPL1.1  
\- Date: 15/04/2009  
\- Version: 1.1

Reading archive:
----------------

### Extract to path:
```
     with CreateInArchive(CLSID_CFormatZip) do 
     begin   
       OpenFile('c:\test.zip');   
       ExtractTo('c:\test'); 
     end;
```     

### Get file list:
```
     with CreateInArchive(CLSID_CFormat7z) do 
     begin   
       OpenFile('c:\test.7z');   
       for i := 0 to NumberOfItems - 1 do    
         if not ItemIsFolder[i] then      
           Writeln(ItemPath[i]); 
     end;
```

### Extract to stream
```
     with CreateInArchive(CLSID_CFormat7z) do 
     begin   
       OpenFile('c:\test.7z');   
       for i := 0 to NumberOfItems - 1 do     
         if not ItemIsFolder[i] then       
           ExtractItem(i, stream, false); 
     end;
```

### Extract "n" Items
```
    function GetStreamCallBack(sender: Pointer; index: Cardinal;  var outStream: ISequentialOutStream): HRESULT; stdcall;
    begin  
      case index of 
        ...    
        outStream := T7zStream.Create(aStream, soReference);  
        Result := S_OK;
    end;
    
    procedure TMainForm.ExtractClick(Sender: TObject);
    var  
      i: integer;  
      items: array[0..2] of Cardinal;
    begin  
      with CreateInArchive(CLSID_CFormat7z) do  
      begin    
        OpenFile('c:\test.7z');
```

### Open stream
```
     with CreateInArchive(CLSID_CFormatZip) do 
     begin   
       OpenStream(T7zStream.Create(TFileStream.Create('c:\test.zip', fmOpenRead), soOwned));   
       OpenStream(aStream, soReference);   
       ... 
     end;
```

### Progress bar
```
     function ProgressCallback(sender: Pointer; total: boolean; value: int64): HRESULT; stdcall; 
     begin   
       if total then     
         Mainform.ProgressBar.Max := value 
       else     
         Mainform.ProgressBar.Position := value;   
       Result := S_OK; 
     end; 
     
     procedure TMainForm.ExtractClick(Sender: TObject); 
     begin   
       with CreateInArchive(CLSID_CFormatZip) do   
       begin     
         OpenFile('c:\test.zip');     
         SetProgressCallback(nil, ProgressCallback);     
         ...   
       end; 
     end;
```

### Password
```
     function PasswordCallback(sender: Pointer; var password: WideString): HRESULT; stdcall; 
     begin
```

Writing archive
---------------
```
     procedure TMainForm.ExtractAllClick(Sender: TObject); 
     var   
       Arch: I7zOutArchive; 
     begin   
       Arch := CreateOutArchive(CLSID_CFormat7z);   
       // add a file   
       Arch.AddFile('c:\test.bin', 'folder\test.bin');   
       
       // add files using willcards and recursive search   
       Arch.AddFiles('c:\test', 'folder', '*.pas;*.dfm', true);   
       
       // add a stream   
       Arch.AddStream(aStream, soReference, faArchive, CurrentFileTime, CurrentFileTime, 'folder\test.bin', false, false);   
       
       // compression level   
       SetCompressionLevel(Arch, 5);   
       
       // compression method if <> LZMA   
       SevenZipSetCompressionMethod(Arch, m7BZip2);   
       
       // add a progress bar 
       ...   
       Arch.SetProgressCallback(...);   
       
       // set a password if necessary   
       Arch.SetPassword('password');   
       
       // Save to file   
       Arch.SaveToFile('c:\test.zip');   
       
       // or a stream   
       Arch.SaveToStream(aStream); 
     end;
```       
