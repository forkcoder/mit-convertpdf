<?php
/**
 * OfficeToPdfConverter
 * 
 * A PHP 7.4 compatible module for converting office documents to PDF
 * using LibreOffice command-line tools.
 * 
 * Supports: .doc, .docx, .xls, .xlsx, .ppt, .pptx, .odt, .ods, .odp, .rtf, .txt
 * 
 * @author Your Name
 * @version 1.0.0
 */

class OfficeToPdfConverter
{
    /**
     * Path to LibreOffice executable
     * Common paths: 
     * - Linux: /usr/bin/libreoffice or /usr/bin/soffice
     * - macOS: /Applications/LibreOffice.app/Contents/MacOS/soffice
     * - Windows: C:\Program Files\LibreOffice\program\soffice.exe
     * 
     * @var string
     */
    private $libreOfficePath;
    
    /**
     * Output directory for converted PDFs
     * 
     * @var string
     */
    private $outputDir;
    
    /**
     * Supported office document extensions
     * 
     * @var array
     */
    private $supportedExtensions = [
        'doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx',
        'odt', 'ods', 'odp', 'rtf', 'txt'
    ];
    
    /**
     * Constructor
     * 
     * @param string|null $libreOfficePath Custom path to LibreOffice executable
     * @param string|null $outputDir Custom output directory (defaults to same directory as source file)
     */
    public function __construct($libreOfficePath = null, $outputDir = null)
    {
        $this->libreOfficePath = $libreOfficePath ?: $this->detectLibreOfficePath();
        $this->outputDir = $outputDir;
    }
    
    /**
     * Detect LibreOffice installation path
     * 
     * @return string
     * @throws Exception if LibreOffice is not found
     */
    private function detectLibreOfficePath()
    {
        $possiblePaths = [];
        
        // Detect OS
        $os = strtoupper(substr(PHP_OS, 0, 3));
        
        if ($os === 'WIN') {
            // Windows paths
            $possiblePaths = [
                'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
                'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe',
                'soffice.exe' // In PATH
            ];
        } elseif ($os === 'DAR') {
            // macOS paths
            $possiblePaths = [
                '/Applications/LibreOffice.app/Contents/MacOS/soffice',
                '/usr/local/bin/libreoffice',
                '/usr/local/bin/soffice',
                'libreoffice', // In PATH
                'soffice' // In PATH
            ];
        } else {
            // Linux/Unix paths
            $possiblePaths = [
                '/usr/bin/libreoffice',
                '/usr/bin/soffice',
                '/usr/local/bin/libreoffice',
                '/usr/local/bin/soffice',
                'libreoffice', // In PATH
                'soffice' // In PATH
            ];
        }
        
        // Check each path
        foreach ($possiblePaths as $path) {
            if ($this->isExecutable($path)) {
                return $path;
            }
        }
        
        throw new Exception(
            'LibreOffice not found. Please install LibreOffice or specify the path manually. ' .
            'Tried paths: ' . implode(', ', $possiblePaths)
        );
    }
    
    /**
     * Check if a file is executable
     * 
     * @param string $path
     * @return bool
     */
    private function isExecutable($path)
    {
        // If it's just a command name (no path separators), check if it's in PATH
        if (strpos($path, DIRECTORY_SEPARATOR) === false && strpos($path, '/') === false) {
            $result = shell_exec("which " . escapeshellarg($path) . " 2>/dev/null");
            return !empty(trim($result));
        }
        
        // Check if file exists and is executable
        return file_exists($path) && is_executable($path);
    }
    
    /**
     * Convert office document to PDF
     * 
     * @param string $inputFile Path to input office document
     * @param string|null $outputFile Optional output PDF path (if null, uses same name as input)
     * @return string Path to converted PDF file
     * @throws Exception if conversion fails
     */
    public function convertToPdf($inputFile, $outputFile = null)
    {
        // Validate input file
        if (!file_exists($inputFile)) {
            throw new Exception("Input file not found: {$inputFile}");
        }
        
        if (!is_readable($inputFile)) {
            throw new Exception("Input file is not readable: {$inputFile}");
        }
        
        // Check file extension
        $extension = strtolower(pathinfo($inputFile, PATHINFO_EXTENSION));
        if (!in_array($extension, $this->supportedExtensions)) {
            throw new Exception(
                "Unsupported file extension: {$extension}. " .
                "Supported: " . implode(', ', $this->supportedExtensions)
            );
        }
        
        // Determine output file path
        if ($outputFile === null) {
            $inputDir = dirname($inputFile);
            $inputBasename = pathinfo($inputFile, PATHINFO_FILENAME);
            $outputDir = $this->outputDir ?: $inputDir;
            
            // Ensure output directory exists
            if (!is_dir($outputDir)) {
                if (!mkdir($outputDir, 0755, true)) {
                    throw new Exception("Cannot create output directory: {$outputDir}");
                }
            }
            
            $outputFile = $outputDir . DIRECTORY_SEPARATOR . $inputBasename . '.pdf';
        }
        
        // Ensure output directory exists
        $outputDir = dirname($outputFile);
        if (!is_dir($outputDir)) {
            if (!mkdir($outputDir, 0755, true)) {
                throw new Exception("Cannot create output directory: {$outputDir}");
            }
        }
        
        // Prepare temporary directory for LibreOffice conversion
        $tempDir = sys_get_temp_dir() . DIRECTORY_SEPARATOR . 'libreoffice_convert_' . uniqid();
        if (!mkdir($tempDir, 0755, true)) {
            throw new Exception("Cannot create temporary directory: {$tempDir}");
        }
        
        try {
            // Build LibreOffice command
            // --headless: Run without GUI
            // --convert-to pdf: Convert to PDF format
            // --outdir: Output directory
            $command = sprintf(
                '%s --headless --convert-to pdf --outdir %s %s 2>&1',
                escapeshellarg($this->libreOfficePath),
                escapeshellarg($tempDir),
                escapeshellarg($inputFile)
            );
            
            // Execute conversion
            $output = [];
            $returnCode = 0;
            exec($command, $output, $returnCode);
            
            if ($returnCode !== 0) {
                throw new Exception(
                    "LibreOffice conversion failed with return code {$returnCode}. " .
                    "Output: " . implode("\n", $output)
                );
            }
            
            // Find the generated PDF file
            $inputBasename = pathinfo($inputFile, PATHINFO_FILENAME);
            $generatedPdf = $tempDir . DIRECTORY_SEPARATOR . $inputBasename . '.pdf';
            
            if (!file_exists($generatedPdf)) {
                // Sometimes LibreOffice adds spaces or changes the filename
                $files = glob($tempDir . DIRECTORY_SEPARATOR . '*.pdf');
                if (empty($files)) {
                    throw new Exception(
                        "PDF file was not generated. LibreOffice output: " . implode("\n", $output)
                    );
                }
                $generatedPdf = $files[0];
            }
            
            // Move PDF to final destination
            if (!copy($generatedPdf, $outputFile)) {
                throw new Exception("Failed to copy PDF to destination: {$outputFile}");
            }
            
            // Verify output file exists and is readable
            if (!file_exists($outputFile) || !is_readable($outputFile)) {
                throw new Exception("Output PDF file was not created successfully: {$outputFile}");
            }
            
            return $outputFile;
            
        } finally {
            // Clean up temporary directory
            $this->removeDirectory($tempDir);
        }
    }
    
    /**
     * Convert multiple office documents to PDF
     * 
     * @param array $inputFiles Array of input file paths
     * @param string|null $outputDir Optional output directory for all PDFs
     * @return array Array of output PDF file paths
     * @throws Exception if any conversion fails
     */
    public function convertMultipleToPdf($inputFiles, $outputDir = null)
    {
        if (!is_array($inputFiles)) {
            throw new Exception("Input must be an array of file paths");
        }
        
        $results = [];
        $errors = [];
        
        foreach ($inputFiles as $inputFile) {
            try {
                if ($outputDir) {
                    $basename = pathinfo($inputFile, PATHINFO_FILENAME);
                    $pdfPath = $outputDir . DIRECTORY_SEPARATOR . $basename . '.pdf';
                    $pdfPath = $this->convertToPdf($inputFile, $pdfPath);
                } else {
                    $pdfPath = $this->convertToPdf($inputFile);
                }
                $results[] = $pdfPath;
            } catch (Exception $e) {
                $errors[] = "Failed to convert {$inputFile}: " . $e->getMessage();
            }
        }
        
        if (!empty($errors)) {
            throw new Exception("Some conversions failed:\n" . implode("\n", $errors));
        }
        
        return $results;
    }
    
    /**
     * Check if file extension is supported
     * 
     * @param string $filePath
     * @return bool
     */
    public function isSupported($filePath)
    {
        $extension = strtolower(pathinfo($filePath, PATHINFO_EXTENSION));
        return in_array($extension, $this->supportedExtensions);
    }
    
    /**
     * Get list of supported file extensions
     * 
     * @return array
     */
    public function getSupportedExtensions()
    {
        return $this->supportedExtensions;
    }
    
    /**
     * Set LibreOffice path manually
     * 
     * @param string $path
     * @throws Exception if path is not valid
     */
    public function setLibreOfficePath($path)
    {
        if (!$this->isExecutable($path)) {
            throw new Exception("LibreOffice executable not found or not executable: {$path}");
        }
        $this->libreOfficePath = $path;
    }
    
    /**
     * Get current LibreOffice path
     * 
     * @return string
     */
    public function getLibreOfficePath()
    {
        return $this->libreOfficePath;
    }
    
    /**
     * Recursively remove directory
     * 
     * @param string $dir
     * @return bool
     */
    private function removeDirectory($dir)
    {
        if (!is_dir($dir)) {
            return false;
        }
        
        $files = array_diff(scandir($dir), ['.', '..']);
        foreach ($files as $file) {
            $path = $dir . DIRECTORY_SEPARATOR . $file;
            is_dir($path) ? $this->removeDirectory($path) : unlink($path);
        }
        
        return rmdir($dir);
    }
}
