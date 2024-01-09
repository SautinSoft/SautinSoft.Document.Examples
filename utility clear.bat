rem Delete all bin, obj, vs.
set DELETEPATH="d:\Work\Products\Document .Net\GitHub\SautinSoft.Document.Examples\CSharp" 

for /d /r "%DELETEPATH%" %%d in (bin,obj,.vs) do @if exist "%%d" rd /s/q "%%d"