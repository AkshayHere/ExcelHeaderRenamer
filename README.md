## Excel Header Renamer

Sample application created for updating the headers within an excel file

```bash
# Build the application for release
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -o "Releases\v1.0.0"

# Compress for `publish` in github
Compress-Archive -Path publish\* -DestinationPath Releases\ExcelProcessorService{releaseTag}.zip
```
