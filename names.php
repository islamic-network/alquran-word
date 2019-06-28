<?php

require 'vendor/autoload.php';

use PhpOffice\PhpWord\Shared\Converter;
use PhpOffice\PhpWord\Style\TablePosition;
use ArabicNumbers\Numbers;
use PhpOffice\PhpWord\SimpleType\TextAlignment;

// Set writers
$writers = array('Word2007' => 'docx', 'ODText' => 'odt', 'RTF' => 'rtf', 'HTML' => 'html');

/**
 * Write documents
 *
 * @param \PhpOffice\PhpWord\PhpWord $phpWord
 * @param string $filename
 * @param array $writers
 *
 * @return string
 */
function write($phpWord, $filename, $writers)
{
    $result = '';
    // Write documents
    foreach ($writers as $format => $extension) {
        $result .= date('H:i:s') . " Write to {$format} format";
        if (null !== $extension) {
            $targetFile = __DIR__ . "/results/{$filename}.{$extension}";
            $phpWord->save($targetFile, $format);
        } else {
            $result .= ' ... NOT DONE!';
        }
        $result .= "\n";
    }
    $result .= getEndingNotes($writers, $filename);
    return $result;
}

/**
 * Get ending notes
 *
 * @param array $writers
 * @param mixed $filename
 * @return string
 */
function getEndingNotes($writers, $filename)
{
    $result = '';
    $result .= date('H:i:s') . ' Done writing file(s)' . "\n";
    $result .= date('H:i:s') . ' Peak memory usage: ' . (memory_get_peak_usage(true) / 1024 / 1024) . ' MB' . "\n";
    return $result;
}

// New Word Document
echo date('H:i:s'), ' Create new PhpWord object', "\n";
$phpWord = new \PhpOffice\PhpWord\PhpWord();
$phpWord->setDefaultFontName('Arial');
$section = $phpWord->addSection(['paperSize' => 'A3']);

$en = json_decode(file_get_contents('files/english.json'));
$ar = json_decode(file_get_contents('files/uthmani.json'));

$currentPage = 1;
foreach ($en->data->surahs as $surahKey => $surah) {
        $section->addText(str_replace('سورة', 'سُورَةُ', $surah->name), ['rtl' => true], ['rtl' => true]);
}

// Save file
echo write($phpWord, basename(__FILE__, '.php'), $writers);
