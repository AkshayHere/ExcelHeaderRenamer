## Excel Header Renamer

Sample application created for updating the headers within an excel file

```bash
# Build the application for release
dotnet publish -c Release -r win-x64 --self-contained true ^
-p:PublishSingleFile=true ^
-p:IncludeNativeLibrariesForSelfExtract=true ^
-o publish

# Compress for `publish` in github
Compress-Archive -Path publish\* -DestinationPath releases\ExcelProcessorService{releaseTag}.zip
```
