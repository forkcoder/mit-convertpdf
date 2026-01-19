<?php
/**
 * Example usage of OfficeToPdfConverter
 */

require_once 'OfficeToPdfConverter.php';

try {
    // Initialize converter (auto-detects LibreOffice)
    $converter = new OfficeToPdfConverter();
    
    echo "LibreOffice path: " . $converter->getLibreOfficePath() . "\n";
    echo "Supported extensions: " . implode(', ', $converter->getSupportedExtensions()) . "\n\n";
    
    // Example 1: Convert a single file
    // Replace 'document.docx' with your actual file path
    if (isset($argv[1])) {
        $inputFile = $argv[1];
        
        if (!file_exists($inputFile)) {
            die("Error: File not found: {$inputFile}\n");
        }
        
        echo "Converting: {$inputFile}\n";
        
        try {
            $outputFile = $converter->convertToPdf($inputFile);
            echo "Success! PDF created: {$outputFile}\n";
        } catch (Exception $e) {
            echo "Error: " . $e->getMessage() . "\n";
            exit(1);
        }
    } else {
        echo "Usage: php example.php <path-to-office-document>\n";
        echo "Example: php example.php document.docx\n\n";
        
        // Example 2: Convert with custom output path
        echo "Example code:\n";
        echo "// Convert single file\n";
        echo "\$converter = new OfficeToPdfConverter();\n";
        echo "\$pdfPath = \$converter->convertToPdf('document.docx');\n";
        echo "// PDF will be saved as 'document.pdf' in the same directory\n\n";
        
        echo "// Convert with custom output path\n";
        echo "\$pdfPath = \$converter->convertToPdf('document.docx', '/path/to/output.pdf');\n\n";
        
        echo "// Convert multiple files\n";
        echo "\$files = ['file1.docx', 'file2.xlsx', 'file3.pptx'];\n";
        echo "\$pdfs = \$converter->convertMultipleToPdf(\$files, '/path/to/output/dir');\n\n";
        
        echo "// Check if file is supported\n";
        echo "if (\$converter->isSupported('document.docx')) {\n";
        echo "    \$converter->convertToPdf('document.docx');\n";
        echo "}\n";
    }
    
} catch (Exception $e) {
    echo "Error: " . $e->getMessage() . "\n";
    exit(1);
}
