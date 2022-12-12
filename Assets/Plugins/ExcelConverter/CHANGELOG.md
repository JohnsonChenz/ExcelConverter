## CHANGELOG

### v1.1.1
- Fixed : Error on building player due to the wrong setup of the Script Assembly Definition for runtime scripts.
- Optimized code in FileExporter & ExcelConverter & SheetExportConfig related to sheet export process.

### v1.1.0

#### Changes on Json Config Generetor :
 
- Because changes below, you may need to remake your own json config.
- Added option to enable MainKey/Subkey
- Added **MainKeySelectType**, please checkout **README.md** to see how it works.
- More approprivately naming for some variables.

#### Others :
- Script optimization & simplfied on FileExporter & JsonConfig & SheetExportConfig

### v1.0.4
- Deleted none-use using at FileExporter.cs

### v1.0.3
- Fixed : Json-Formatting Toggle does not Enable/Disable properly according to the Export Type Option.
- Simplified path information of progress bar.
- Adjusted min size of Editor.

### v1.0.2
- Totally re-layout the ExcelConverter editor.

### v1.0.1
- Fixed : Overflow display of Progressbar's text.

### v1.0.0
- An editor tool that allows you convert excel(xlsx/xls) into json/bson files.