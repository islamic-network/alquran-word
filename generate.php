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

$table = $section->addTable();
$currentPage = 1;
foreach ($en->data->surahs as $surahKey => $surah) {
    foreach ($surah->ayahs as $ayahKey => $ayah) {
        if ($currentPage < $ayah->page) {
            $currentPage++;
            $section->addPageBreak();
            $table = $section->addTable();
        }
        if ($ayah->numberInSurah === 1) {
            $table->addRow();
            //$table->addCell()->addText('V', ['name' => 'AGA Islamic Phrases Regular', 'size' => 26]);
            $table->addCell(null, ['bgColor' => ''])->addText($ar->data->surahs[$surahKey]->name, ['name' => 'me_quran', 'size' => 20], ['align' => 'center']); //KFGQPC Uthmanic Script HAFS Regular
            //$table->addCell()->addText('V', ['name' => 'AGA Islamic Phrases Regular', 'size' => 26]);
            if ($surah->number !== 1 && $surah->number !== 9) {
                //$table->addRow();
                //$table->addCell();
                //$table->addCell(null, ['bgColor' => ''])->addText($ar->data->surahs[0]->ayahs[0]->text, ['name' => 'me_quran', 'size' => 28], ['align' => 'center']); //KFGQPC Uthmanic Script HAFS Regular
                //$table->addCell();
            }
        }
        $table->addRow();
        //$table->addCell();
        if ($ayah->numberInSurah === 1 && $surah->number !== 1 && $surah->number !== 9) {
            $table->addCell(null)->addText(mb_substr($ar->data->surahs[$surahKey]->ayahs[$ayahKey]->text, 39). '   ' . Numbers::latinToArabic($ayah->numberInSurah), ['name' => 'me_quran', 'size' => 28], ['align' => 'right']);
        } else {
            $table->addCell(null)->addText($ar->data->surahs[$surahKey]->ayahs[$ayahKey]->text . '   ' . Numbers::latinToArabic($ayah->numberInSurah), ['name' => 'me_quran', 'size' => 28], ['align' => 'right']);
        }
        //$table->addCell();
        $table->addRow();
        //$table->addCell(;
        $table->addCell(null)->addText($ayah->numberInSurah . '. ' .$ayah->text, ['size' => 14]);
        //$table->addCell();
    }
}
$table->addRow();
$table->addCell();

// Save file
echo write($phpWord, basename(__FILE__, '.php'), $writers);
