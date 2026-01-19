# Office to PDF Converter

A PHP 7.4 compatible module for converting office documents to PDF using LibreOffice command-line tools.

## Features

- ✅ PHP 7.4+ compatible
- ✅ No Composer dependencies required
- ✅ Auto-detects LibreOffice installation
- ✅ Supports multiple office formats:
  - Microsoft Office: `.doc`, `.docx`, `.xls`, `.xlsx`, `.ppt`, `.pptx`
  - OpenDocument: `.odt`, `.ods`, `.odp`
  - Other: `.rtf`, `.txt`
- ✅ Batch conversion support
- ✅ Cross-platform (Linux, macOS, Windows)

## Requirements

- PHP 7.4 or higher
- LibreOffice installed on the system
  - **Linux**: `sudo apt-get install libreoffice` (Ubuntu/Debian) or `sudo yum install libreoffice` (CentOS/RHEL)
  - **macOS**: Download from [LibreOffice website](https://www.libreoffice.org/download/download/) or `brew install --cask libreoffice`
  - **Windows**: Download installer from [LibreOffice website](https://www.libreoffice.org/download/download/)

## Installation

1. Copy `OfficeToPdfConverter.php` to your project directory
2. Include it in your PHP script:
   ```php
   require_once 'OfficeToPdfConverter.php';
   ```

## Usage

### Basic Conversion

```php
<?php
require_once 'OfficeToPdfConverter.php';

try {
    $converter = new OfficeToPdfConverter();
    $pdfPath = $converter->convertToPdf('document.docx');
    echo "PDF created: " . $pdfPath;
} catch (Exception $e) {
    echo "Error: " . $e->getMessage();
}
```

### Custom Output Path

```php
$converter = new OfficeToPdfConverter();
$pdfPath = $converter->convertToPdf('document.docx', '/path/to/output.pdf');
```

### Custom LibreOffice Path

```php
// macOS example
$converter = new OfficeToPdfConverter('/Applications/LibreOffice.app/Contents/MacOS/soffice');

// Linux example
$converter = new OfficeToPdfConverter('/usr/bin/libreoffice');

// Windows example
$converter = new OfficeToPdfConverter('C:\\Program Files\\LibreOffice\\program\\soffice.exe');
```

### Batch Conversion

```php
$converter = new OfficeToPdfConverter();
$files = ['file1.docx', 'file2.xlsx', 'file3.pptx'];
$pdfs = $converter->convertMultipleToPdf($files, '/path/to/output/dir');
```

### Check File Support

```php
if ($converter->isSupported('document.docx')) {
    $converter->convertToPdf('document.docx');
}
```

## Command Line Usage

You can also use the example script from command line:

```bash
php example.php document.docx
```

## API Reference

### `OfficeToPdfConverter` Class

#### Constructor

```php
new OfficeToPdfConverter($libreOfficePath = null, $outputDir = null)
```

- `$libreOfficePath`: Optional custom path to LibreOffice executable
- `$outputDir`: Optional default output directory for all conversions

#### Methods

##### `convertToPdf($inputFile, $outputFile = null)`

Convert a single office document to PDF.

- **Parameters:**
  - `$inputFile` (string): Path to input office document
  - `$outputFile` (string|null): Optional output PDF path
- **Returns:** Path to converted PDF file
- **Throws:** Exception if conversion fails

##### `convertMultipleToPdf($inputFiles, $outputDir = null)`

Convert multiple office documents to PDF.

- **Parameters:**
  - `$inputFiles` (array): Array of input file paths
  - `$outputDir` (string|null): Optional output directory
- **Returns:** Array of output PDF file paths
- **Throws:** Exception if any conversion fails

##### `isSupported($filePath)`

Check if file extension is supported.

- **Parameters:**
  - `$filePath` (string): Path to file
- **Returns:** bool

##### `getSupportedExtensions()`

Get list of supported file extensions.

- **Returns:** array

##### `setLibreOfficePath($path)`

Set LibreOffice path manually.

- **Parameters:**
  - `$path` (string): Path to LibreOffice executable
- **Throws:** Exception if path is not valid

##### `getLibreOfficePath()`

Get current LibreOffice path.

- **Returns:** string

## Troubleshooting

### LibreOffice Not Found

If you get an error that LibreOffice is not found:

1. Make sure LibreOffice is installed
2. Specify the path manually:
   ```php
   $converter = new OfficeToPdfConverter('/path/to/libreoffice');
   ```

### Permission Errors

Make sure:
- Input files are readable
- Output directory is writable
- PHP has permission to execute LibreOffice

### Conversion Fails

- Check that LibreOffice can open the file manually
- Verify file is not corrupted
- Check PHP error logs for detailed error messages
- Ensure sufficient disk space is available

## License

This module is provided as-is for use in your projects.

## Notes

- The converter uses LibreOffice's headless mode (no GUI)
- Temporary files are automatically cleaned up after conversion
- Large files may take some time to convert
- The module is compatible with PHP 7.4 and higher
