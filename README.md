# File Conversion Library

- [**Описание библиотеки по преобразованию файлов**](#описание-библиотеки-по-преобразованию-файлов)
- [**Функционал библиотеки**](#функционал-библиотеки)
  - [**Модуль Images:**](#модуль-images)
    - [**Конвертация JPG в PNG:**](#конвертация-jpg-в-png)
    - [**Конвертация PNG в JPG:**](#конвертация-png-в-jpg)
  - [**Модуль LibreOffice:**](#модуль-libreoffice)
    - [**Конвертация Docx в PDF:**](#конвертация-docx-в-pdf)
    - [**Конвертация PDF в Docx:**](#конвертация-pdf-в-docx)
    - [**Конвертация PPTX в PDF:**](#конвертация-pptx-в-pdf)
  - [**Модуль MSOffice:**](#модуль-msoffice)
    - [**Конвертация Docx в PDF:**](#конвертация-docx-в-pdf-1)
    - [**Конвертация PDF в Docx:**](#конвертация-pdf-в-docx-1)
    - [**Конвертация PPTX в PDF:**](#конвертация-pptx-в-pdf-1)
  - [**Модуль PDF:**](#модуль-pdf)
    - [**Объединение PDF файлов:**](#объединение-pdf-файлов)
    - [**Разделение PDF файла:**](#разделение-pdf-файла)
    - [**Конвертация JPG в PDF:**](#конвертация-jpg-в-pdf)
    - [**Конвертация PDF в JPG:**](#конвертация-pdf-в-jpg)
- [**Работа с консольными приложениями**](#работа-с-консольными-приложениями)
  - [**Установка**](#установка)
  - [**Начало работы**](#начало-работы)
  - [**Конвертация JPG в Png:**](#конвертация-jpg-в-png-1)
  - [**Конвертация Png в JPG:**](#конвертация-png-в-jpg-1)
  - [**Преобразование файлов при помощи модуля MSOffice:**](#преобразование-файлов-при-помощи-модуля-msoffice)
    - [**Конвертация Docx в PDF:**](#конвертация-docx-в-pdf-2)
    - [**Конвертация PDF в Docx:**](#конвертация-pdf-в-docx-2)
    - [**Конвертация PPTX в PDF:**](#конвертация-pptx-в-pdf-2)
  - [**Преобразование файлов при помощи модуля LibreOffice:**](#преобразование-файлов-при-помощи-модуля-libreoffice)
    - [**Конвертация Docx в PDF:**](#конвертация-docx-в-pdf-3)
    - [**Конвертация PDF в Docx:**](#конвертация-pdf-в-docx-3)
    - [**Конвертация PPTX в PDF:**](#конвертация-pptx-в-pdf-3)
  - [**Объединение PDF файлов:**](#объединение-pdf-файлов-1)
  - [**Разделение PDF файла:**](#разделение-pdf-файла-1)
  - [**Конвертация JPG в PDF:**](#конвертация-jpg-в-pdf-1)
  - [**Конвертация PDF в JPG:**](#конвертация-pdf-в-jpg-1)
- [**Интеграция**](#интеграция)
  - [**Интеграция в проект на Node.js:**](#интеграция-в-проект-на-nodejs)
  - [**Интеграция в проект на PHP:**](#интеграция-в-проект-на-php)
  - [**Интеграция в проект на Python:**](#интеграция-в-проект-на-python)


# **Описание библиотеки по преобразованию файлов**
Библиотека состоит из 4 модулей:

- `Images` - модуль для работы с изображениями
- `PDF` - модуль для работы с PDF файлами.
- `LibreOffice` - модуль для работы с Docx, Pptx и PDF файлами.
- `MSOffice` - модуль для работы с Docx, Pptx и PDF файлами.

Модули `LibreOffice` и `MSOffice` предназначены для работы с одними и теми же типами файлов. Модуль `MSOffice` будет работать только на ОС Windows, используя установленный MS Office. Модуль `LibreOffice` будет работать как на ОС Windows, так и на ОС Linux, используя установленный LibreOffice.

Для возможности использовать библиотеку на разных языках программирования были созданы консольные приложения для каждого модуля библиотеки:

- `ImageConverter` - для работы с изображениями.
- `PDFConverter` - для работы с PDF файлами.
- `LibreOfficeConverter` - модуль для работы с Docx, Pptx и PDF файлами.
- `MSOfficeConverter` - для работы с Docx, Pptx и PDF файлами.

# **Функционал библиотеки**
## **Модуль Images:**

### **Конвертация JPG в PNG:**
Существует 2 метода: 

1. `JpgFileToPngFile(string jpgFileName, string pngFileName)`: 
- `string jpgFileName` - путь до JPG файла.
- `string pngFileName` - путь, куда положить PNG файл.\
После вызова метода появится PNG файл с соответствующими директорией и именем.

2. `JpgFileToPngFile(string jpgFileName)`:
- `string jpgFileName` - путь до JPG файла.\
После вызова метода появится PNG файл с таким же именем, что и JPG файл.

### **Конвертация PNG в JPG:**
Существует 2 метода: 

1. `PngFileToJpgFile(string pngFileName, string jpgFileName)`: 
- `string pngFileName` - путь до PNG файла.
- `string jpgFileName`- путь, куда положить JPG файл.\
После вызова метода появится JPG файл с соответствующими директорией и именем.

2. PngFileToJpgFile(string pngFileName):
- `string pngFileName` - путь до PNG файла.\
После вызова метода появится JPG файл с таким же именем, что и PNG файл.

## **Модуль LibreOffice:**
Перед началом работы необходимо задать путь до soffice.exe полю `sofficePath`: 

```csharp
public static string sofficePath = @"Путь до soffice.exe ";
```

### **Конвертация Docx в PDF:**
Существует 2 метода: 

1. `DocxFileToPdfFile(string wordFileName, string pdfFileName)`: 
- `string wordFileName` - путь до Docx файла.
- `string pdfFileName` - путь, куда положить PDF файл.\
После вызова метода появится PDF файл с соответствующими директорией и именем.

2. `DocxFileToPdfFile(string wordFileName)`:
- `string wordFileName` - путь до Docx файла\
После вызова метода появится PDF файл с таким же именем, что и Docx файл, в той же папке.

### **Конвертация PDF в Docx:**
Существует 2 метода: 

1. `PdfFileToDocxFile(string pdfFileName, string wordFileName)`: 
- `string pdfFileName` - путь до PDF файла.
- `string wordFileName` - путь, куда положить Docx файл.\
После вызова метода появится Docx файл с соответствующими директорией и именем.

2. `PdfFileToDocxFile(string pdfFileName)`:
- `string pdfFileName` - путь до PDF файла.\
После вызова метода появится Docx файл с таким же именем, что и PDF файл, в той же папке.

### **Конвертация PPTX в PDF:**
Существует 2 метода: 

1. `PptxFileToPdfFile(string pptxFileName, string pdfFileName)`: 
- `string pptxFileName` - путь до PPTX файла.
- `string pdfFileName` - путь, куда положить PDF файл.\
После вызова метода появится PPTX файл с соответствующими директорией и именем.

2. `PptxFileToPdfFile(string pptxFileName)`:
- `string pptxFileName` - путь до PPTX файла.\
После вызова метода появится PPTX файл с таким же именем, что и PDF файл, в той же папке.

## **Модуль MSOffice:**
### **Конвертация Docx в PDF:**
Существует 2 метода: 

1. `DocxFileToPdfFile(string wordFileName, string pdfFileName)`: 
- `string wordFileName` - путь до Docx файла.
- `string pdfFileName` - путь, куда положить PDF файл.\
После вызова метода появится PDF файл с соответствующими директорией и именем.

2. `DocxFileToPdfFile(string wordFileName)`:
- `string wordFileName` - путь до Docx файла.\
После вызова метода появится PDF файл с таким же именем, что и Docx файл, в той же папке.

### **Конвертация PDF в Docx:**
Существует 2 метода: 

1. `PdfFileToDocxFile(string pdfFileName, string wordFileName)`: 
- `string pdfFileName` - путь до PDF файла.
- `string wordFileName` - путь, куда положить Docx файл.\
После вызова метода появится Docx файл с соответствующими директорией и именем.

2. `PdfFileToDocxFile(string pdfFileName)`:
- `string pdfFileName` - путь до PDF файла.\
После вызова метода появится Docx файл с таким же именем, что и PDF файл, в той же папке.

### **Конвертация PPTX в PDF:**
Существует 2 метода: 

1. `PptxFileToPdfFile(string pptxFileName, string pdfFileName)`: 
- `string pptxFileName` - путь до PPTX файла.
- `string pdfFileName` - путь, куда положить PDF файл.\
После вызова метода появится PPTX файл с соответствующими директорией и именем.

2. `PptxFileToPdfFile(string pptxFileName)`:
- `string pptxFileName` - путь до PPTX файла.\
После вызова метода появится PPTX файл с таким же именем, что и PDF файл, в той же папке.

## **Модуль PDF:**
### **Объединение PDF файлов:**
`MergePDFs(string[] pdfFiles, string pdfOutput)`: 
- `string[] pdfFiles`- список путей до PDF файлов, которые необходимо объеденить.
- `string pdfOutput`- путь, куда положить PDF файл, который будет содержать в себе все необходимые PDF файлы.\
После вызова метода появится файл, содержащий в себе необходимый PDF файлы.

### **Разделение PDF файла:**
Существует 2 метода: 

1. `SplitPDF(string pdfInput, int pageSplitFrom, string pdf1Output, string pdf2Output)`: 
- `string pdfInput` - путь до PDF файла.
- `int pageSplitFrom` – номер страницы, по которую необходимо разделить файл.
- `string pdf1Output`- путь, куда положить выходной файл, содержащий первую часть входного файла.
- `string pdf2Output`- путь, куда положить выходной файл, содержащий вторую часть входного файла.\
После вызова метода появятся 2 PDF файла с соответствующими директориями и именами.

2. `PptxFileToPdfFile(string pptxFileName)`:
- `string pptxFileName` - имя PDF файла.\
После вызова метода появятся 2 PDF файла с именами в следущей форме: имя входного файла + "\_splitted1.pdf", имя входного файла + "\_splitted2.pdf"; файлы появятся в той же папке, что и входной файл.

### **Конвертация JPG в PDF:**
`JpgFilesToPdfFile(string[] jpgFiles, string pdfFileName)`:

- `string[] jpgFiles`- список путей до JPG файлов.
- `string pdfFileName`- путь, куда положить PDF файл, который будет содержать в себе JPG файлы.\
После вызова метода появится PDF файл, который будет содержать в себе введенные JPG файлы.

### **Конвертация PDF в JPG:**
Существует 2 метода: 

1. `PdfFileToJpgFiles(string pdfFileName, string jpgFolderName, bool zip = false)`: 
- `string pdfFileName` – путь до PDF файла.
- `string jpgFolderName`- путь, куда положить папку, в которую необходимо положить JPG файлы.
- `bool zip` – параметр, отвечающий за создание архива(по умолчанию `false`).\
Если вызвать метод с `zip = false`, то появится папка с соответствующим именем, содержащая JPG файлы.\
Если вызвать метод с `zip = true`, то появится архив с соответствующим именем, содержащий JPG файлы, архив будет расположен в той же папке, что и входной PDF файл.

2. `PdfFileToJpgFiles(string pdfFileName, bool zip = false)`:
- `string pdfFileName` - путь до PDF файла.
- `bool zip` - параметр, отвечающий за создание архива(по умолчанию `false`).\
Если вызвать метод с `zip = false`, то появится папка с именем по умолчанию, содержащая JPG файлы.
Если вызвать метод с `zip = true`, то появится архив с именем по умолчанию, содержащий JPG файлы, архив будет расположен в той же папке, что и входной PDF файл.

# **Работа с консольными приложениями**

## **Установка**
Скачайте архив с библиотекой и консольными приложениями и распакуйте его.

## **Начало работы**
Необходимо скомпилировать консольные приложения, после этого появятся исполняемые файлы ImageConverter, LibreOfficeConverter, MSOfficeConverter и PDFConverter.

## **Конвертация JPG в Png:**
У команды 2 аргумента: 

- `jpgFileName` - путь до JPG файла.
- `pngFileName` - путь, куда положить Png файл.

Второй параметр не является обязательным, в случае, если этот параметр не указать, Png файл будет сохранен в той же папке, что и JPG файл, с таким же именем.

Пример команды на ОС Windows:

```shell
ImageConverter.exe --method JpgFileToPngFile --jpgFileName input.jpg --pngFileName output.png
```

## **Конвертация Png в JPG:**
У команды 2 аргумента: 

- `pngFileName` - путь до Png файла.
- `jpgFileName` - путь, куда положить JPG файл.

Второй параметр не является обязательным, в случае, если этот параметр не указать, JPG файл будет сохранен в той же папке, что и Png файл, с таким же именем.

Пример команды на ОС Windows:

```shell
ImageConverter.exe --method PngFileToJpgFile --pngFileName input.png --jpgFileName output.jpg
```

## **Преобразование файлов при помощи модуля MSOffice:**
### **Конвертация Docx в PDF:**
У команды 2 аргумента: 

- `wordFileName` - путь до Docx файла.
- `pdfFileName` - путь, куда положить PDF файл.

Второй параметр не является обязательным, в случае, если этот параметр не указать, PDF файл будет сохранен в той же папке, что и Docx файл, с таким же именем.

Пример команды на ОС Windows:

```shell
MSOfficeConverter.exe --method DocxFileToPdfFile --wordFileName input.docx --pdfFileName output.pdf
```

### **Конвертация PDF в Docx:**
У команды 2 аргумента: 

- `pdfFileName` - путь до PDF файла.
- `wordFileName` - путь, куда положить Docx файл.

Второй параметр не является обязательным, в случае, если этот параметр не указать, Docx файл будет сохранен в той же папке, что и PDF файл, с таким же именем.

Пример команды на ОС Windows:

```shell
MSOfficeConverter.exe --method PdfFileToDocxFile --pdfFileName input.pdf --wordFileName output.docx
```

### **Конвертация PPTX в PDF:**
У команды 2 аргумента: 

- `pptxFileName` - путь до PPTX файла.
- `pdfFileName` - путь, куда положить PDF файл.

Второй параметр не является обязательным, в случае, если этот параметр не указать, PDF файл будет сохранен в той же папке, что и Pptx файл, с таким же именем.

Пример команды на ОС Windows:
```shell
MSOfficeConverter.exe --method PptxFileToPdfFile --pptxFileName input.pptx --pdfFileName output.pdf
```

## **Преобразование файлов при помощи модуля LibreOffice:**
### **Конвертация Docx в PDF:**
У команды 2 аргумента: 

- `sofficePath` - путь до soffice.exe
- `wordFileName` - путь до Docx файла.
- `pdfFileName` - путь, куда положить PDF файл.

Первый параметр не является обязательным, в случае, если команда soffice доступна из терминала
Третий параметр не является обязательным, в случае, если этот параметр не указать, PDF файл будет сохранен в той же папке, что и Docx файл, с таким же именем.

Пример команды на ОС Windows:

```shell
LibreOfficeConverter.exe --method DocxFileToPdfFile --sofficePath "C:\Program Files\LibreOffice\program" --wordFileName input.docx --pdfFileName output.pdf
```
### **Конвертация PDF в Docx:**
У команды 2 аргумента: 

- `sofficePath` - путь до soffice.exe
- `pdfFileName` - путь до PDF файла.
- `wordFileName` - путь, куда положить Docx файл.

У команды 3 аргумента: 
1. Путь до soffice.exe. Если команда soffice доступна из терминала, то параметр можно не указывать.
2. Путь до PDF файла и путь, куда положить Docx файл. 
3. Третий параметр не является обязательным, в случае, если этот параметр не указать, Docx файл будет сохранен в той же папке, что и PDF файл, с таким же именем.

Пример команды на ОС Windows:

```shell
LibreOfficeConverter.exe --method PdfFileToDocxFile --sofficePath "C:\Program Files\LibreOffice\program" --pdfFileName input.pdf --wordFileName output.docx
```

### **Конвертация PPTX в PDF:**
У команды 3 аргумента: 

- `sofficePath` - путь до soffice.exe
- `pptxFileName` - путь до PPTX файла.
- `pdfFileName` - путь, куда положить PDF файл.

Третий параметр не является обязательным, в случае, если этот параметр не указать, PDF файл будет сохранен в той же папке, что и Pptx файл, с таким же именем.

Пример команды на ОС Windows:

```shell
LibreOfficeConverter.exe --method PptxFileToPdfFile --sofficePath "C:\Program Files\LibreOffice\program" --pptxFileName input.pptx --pdfFileName output.pdf 
```

## **Объединение PDF файлов:**
У команды 2 аргумента: 

- `pdfFiles` - список PDF файлов.
- `pdfOutput` - имя файла, который будет состоять из объединенных файлов.

Список файлов записывается через пробел, если файл находится в другой папке, то путь до него указывается в кавычках.

Пример команды на ОС Windows:

```shell
PDFConverter.exe --method MergePDFs --pdfFiles firstFile.pdf "path/to/secondFile.pdf" thirdFile.pdf --pdfOutput output.pdf
```

## **Разделение PDF файла:**
У команды 4 аргумента: 

- `pdfInput` - PDF файл, который необходимо разделить.
- `pageSplitFrom` - номер страницы, по которую необходимо разделить файл.
- `pdf1Output` - имя файла, в котором будет первая часть разделенного файла.
- `pdf2Output` - имя файла, в котором будет вторая часть разделенного файла.

Последние 2 аргумента не являются обязательными.

Пример команды на ОС Windows:

```shell
PDFConverter.exe --method SplitPDF --pdfInput input.pdf --pageSplitFrom 5 --pdf1Output firstPart.pdf --pdf2Output secondPart.pdf
```

## **Конвертация JPG в PDF:**
У команды 2 аргумента: 

- `jpgFiles` - список JPG файлов.
- `pdfFileName` - имя PDF фала.

Список JPG файлов записывается через пробел если файл находится в другой папке, то путь до него указывается в кавычках.

Пример команды на ОС Windows:

```shell
PDFConverter.exe --method JpgFilesToPdfFile --jpgFiles firstFile.jpg "path/to/secondFile.jpg" thirdFile.jpg --pdfFileName output.pdf
```

## **Конвертация PDF в JPG:**
У команды 3 аргумента: 

- `pdfFileName` – имя PDF файла.
- `jpgFolderName` – имя папки, в которую будут сохраняться JPG файлы.
- `zip` - аргумент, отвечающий за то, создать архив с JPG файлами или нет.

Последние 2 аргумента не являются обязательными. Если не вводить аргумент `jpgFolderName`, то JPG файлы сохранятся в папку с именем по умолчанию, находящуюся в той же папке, что и PDF файл, если не вводить аргумент `zip`, то архив не будет создаваться. Если ввести в значение аргумента `zip` true, то архив будет создан в той же папке, что и PDF файл.

Пример команды на ОС Windows:

```shell
PDFConverter.exe --method PdfFileToJpgFiles --pdfFileName input.pdf --jpgFolderName JPGFolder --zip true
```

# **Интеграция**
## **Интеграция в проект на Node.js:**
``child_process.execFile()`` в Node.js - это функция, которая позволяет запускать внешние исполняемые файлы как дочерние процессы. Она является частью модуля `child_process`.

Пример использования конвертации JPG в Png:

```js
//Подключаем функцию execFile
const { execFile } = require('child_process');

//Прописываем путь до исполняемого файла
const command = 'path/to/exe/ImageConverter.exe';

//Прописываем необходимые аргументы
const args = [
'--method',
'JpgFileToPngFile',
'--jpgFileName',
'input.jpg',
'--pngFileName',
'output.png',
];

//Запускаем необходимый исполняемый файл
execFile(command, args, (error, stdout, stderr) => {
    if (error) {
        console.error(`exec error: ${error}`);
        return;
    }

    console.log(`stdout: ${stdout}`);
    console.error(`stderr: ${stderr}`);
});
```

После выполнения функции execFile запустится исполняемый файл, и в той же папке, что и исполняемый файл, появится файл `output.png`.

## **Интеграция в проект на PHP:**
`exec()` - это функция в PHP, которая выполняет внешнюю программу. Она позволяет PHP-скриптам взаимодействовать с базовой системой и запускать внешние программы или команды.

Пример использования конвертации JPG в Png:
```php
<?php
    // Используем exec для запуска команды с аргументами
    $result = exec("ImageConverter.exe --method JpgFileToPngFile --jpgFileName input.jpg --pngFileName output.png");
?>
```
После выполнения функции exec запустится исполняемый файл, и в той же папке, что и исполняемый файл, появится файл `output.png`.

## **Интеграция в проект на Python:**
Функция `subprocess.run()` в Python позволяет запускать новые процессы, подключаться к их входным/выходным/ошибочным каналам и получать их возвращаемые значения.

Пример использования конвертации JPG в Png:
```python
    import subprocess

    # Записываем команду в переменную
    command = 'ImageConverter.exe'

    # Записываем аргументы в массив
    args = [
        '--method',
        'JpgFileToPngFile',
        '--jpgFileName',
        'input.jpg',
        '--pngFileName',
        'output.png'
    ]

    # Запускаем процесс с помощью subprocess.run()
    process = subprocess.run([command, *args])
```
После выполнения функции run запустится исполняемый файл, и в той же папке, что и исполняемый файл, появится файл `output.png`.
