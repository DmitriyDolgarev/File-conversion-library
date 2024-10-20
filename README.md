# File Conversion Library

## Начало работы
Для работы с Word и PowerPoint файлами необходим установленный LibreOffice.\
Работа происходит при помощи приложения `soffice.exe`, который по умолчанию установлен по пути `C:\Program Files\LibreOffice\program`\
Если приложение установлено в другой папке, то нужно указать этот путь перед использованием библиотеки:
```csharp
FileConverter.sofficePath = "Путь до soffice.exe"
```
Чтобы убедиться, что библиотека находит файл `soffice.exe` можно написать:
```csharp
FileConverter.SofficeExists
```

## Объединение PDF
```csharp
FileConverter.MergePDFs(string[] pdfFiles, string pdfOutput)
```
- `pdfFiles` - Список PDF файлов для объединения
- `pdfOutput` - PDF файл c результатом

## Разделение PDF
```csharp
FileConverter.SplitPDF(string pdfInput, int pageSplitFrom, string pdf1Output, string pdf2Output)
```
- `pdfInput` - PDF файл, который нужно разделить
- `pageSplitFrom` - страница, с которой разделится файл, причём указанная страница будет добавлена во второй файл
- `pdf1Output` - 1-ый итоговый PDF файл
- `pdf2Output` - 2-ой итоговый PDF файл

```csharp
FileConverter.SplitPDF(string pdfInput, int pageSplitFrom)
```
В результате в той же директории, что и PDF файл, будут созданы 2 файла: `{pdfInput}_splitted1.pdf` и `{pdfInput}_splitted2.pdf`

## JPG в PNG
```csharp
FileConverter.JpgFileToPngFile(string jpgFileName, string pngFileName)
```
- `jpgFileName` - исходный JPG файл
- `pngFileName` - итоговый PNG файл

```csharp
FileConverter.JpgFileToPngFile(string jpgFileName)
```
PNG файл будет создан в той же директории и с тем же названием, что и JPG файл

## PNG в JPG
```csharp
FileConverter.PngFileToJpgFile(string pngFileName, string jpgFileName)
```
- `pngFileName` - исходный PNG файл
- `jpgFileName` - итоговый JPG файл

```csharp
FileConverter.PngFileToJpgFile(string pngFileName)
```
JPG файл будет создан в той же директории и с тем же названием, что и PNG файл

## JPG в PDF
```csharp
FileConverter.JpgFilesToPdfFile(string[] jpgFiles, string pdfFileName)
```
- `jpgFiles` - Список JPG файлов для объединения
- `pdfFileName` - Итоговый PDF файл

## PDF в JPG
```csharp
FileConverter.PdfFileToJpgFiles(string pdfFileName, string jpgFolderName, bool zip = false)
```
- `pdfFileName` - Исходный PDF файл
- `jpgFolderName` - Папка, в которой будут сохранены JPG файлы с названием `page_{i}.pdf`.
- `zip`: (по умолчанию - `false`)
    - Если `true` - то JPG файлы сохранятся в указанной папке
    - Если `false` - то JPG файлы сохранятся в указанном ZIP-архиве 
```csharp
FileConverter.PdfFileToJpgFiles(string pdfFileName, bool zip = false)
```
Папка (или архив) с JPG файлами будет создана в той же директории и с таким же названием, что и PDF файл 

## Word в PDF
```csharp
FileConverter.DocxFileToPdfFile(string wordFileName, string pdfFileFolder)
```
- `wordFileName` - Исходный Word файл
- `pdfFileFolder` - Папка, в которой будет создан PDF файл с тем же названием, что и Word файл

## PDF в Word
```csharp
FileConverter.PdfFileToDocxFile(string pdfFileName, string wordFileFolder)
```
- `pdfFileName` - Исходный PDF файл
- `wordFileFolder` - Папка, в которой будет создан Word файл с тем же названием, что и PDF файл

## PowerPoint в PDF
```csharp
FileConverter.PptxFileToPdfFile(string pptxFileName, string pdfFileFolder)
```
- `pptxFileName` - Исходный PowerPoint файл
- `pdfFileFolder` - Папка, в которой будет создан PDF файл с тем же названием, что и PowerPoint файл