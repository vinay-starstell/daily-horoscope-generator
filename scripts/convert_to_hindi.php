<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Alignment;

$dotenv = new Symfony\Component\Dotenv\Dotenv();
$dotenv->usePutenv()->load(__DIR__ . '/../.env');

$claudeApiKey = $_ENV['CLAUDE_API_KEY'] ?? null;
if (!$claudeApiKey) {
    die("Error: CLAUDE_API_KEY not set\n");
}

$zodiacHindi = [
    "Aries" => "मेष", "Taurus" => "वृषभ", "Gemini" => "मिथुन",
    "Cancer" => "कर्क", "Leo" => "सिंह", "Virgo" => "कन्या",
    "Libra" => "तुला", "Scorpio" => "वृश्चिक", "Sagittarius" => "धनु",
    "Capricorn" => "मकर", "Aquarius" => "कुंभ", "Pisces" => "मीन"
];

function translateToHindi($englishPrediction, $claudeApiKey) {
    $prompt = "You are a Hindi horoscope writer. Summarize the following English horoscope prediction into Hindi.\n\nRules:\n- One compact paragraph only, 4-5 lines\n- Simple conversational Hindi\n- Blend all aspects naturally: Personal, Health, Profession, Emotions, Travel, Luck\n- Slightly predictive and advisory tone\n- NO bullet points, NO subheadings\n- About half the length of the original\n\nEnglish prediction:\n$englishPrediction\n\nWrite only the Hindi paragraph, nothing else:";

    $ch = curl_init('https://api.anthropic.com/v1/messages');
    curl_setopt($ch, CURLOPT_RETURNTRANSFER, 1);
    curl_setopt($ch, CURLOPT_POST, 1);
    curl_setopt($ch, CURLOPT_TIMEOUT, 30);
    curl_setopt($ch, CURLOPT_HTTPHEADER, [
        'Content-Type: application/json',
        "x-api-key: $claudeApiKey",
        'anthropic-version: 2023-06-01'
    ]);
    curl_setopt($ch, CURLOPT_POSTFIELDS, json_encode([
        'model' => 'claude-3-5-sonnet-20241022',
        'max_tokens' => 300,
        'messages' => [['role' => 'user', 'content' => $prompt]]
    ]));

    $response = curl_exec($ch);
    curl_close($ch);

    $data = json_decode($response, true);
    return trim($data['content'][0]['text'] ?? '');
}

if (!file_exists('daily_horoscope.xlsx')) {
    die("Error: daily_horoscope.xlsx not found\n");
}

$wbEn = \PhpOffice\PhpSpreadsheet\IOFactory::load('daily_horoscope.xlsx');
$wsEn = $wbEn->getActiveSheet();

$wbHi = new Spreadsheet();
$wsHi = $wbHi->getActiveSheet();
$wsHi->setTitle("Horoscope Data");

// Title row
$wsHi->mergeCells('A1:C1');
$wsHi->getCell('A1')->setValue("दैनिक राशिफल — StarsTell");
$wsHi->getStyle('A1')->applyFromArray([
    'font' => ['bold' => true, 'size' => 14, 'color' => ['argb' => 'FFFFFFFF']],
    'fill' => ['fillType' => Fill::FILL_SOLID, 'startColor' => ['argb' => 'FF1F3864']],
    'alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER]
]);
$wsHi->getRowDimension(1)->setRowHeight(30);

// Header row
$headers = ['A2' => 'तिथि', 'B2' => 'राशि', 'C2' => 'सारांश'];
foreach ($headers as $cell => $header) {
    $wsHi->getCell($cell)->setValue($header);
    $wsHi->getStyle($cell)->applyFromArray([
        'font' => ['bold' => true, 'color' => ['argb' => 'FFFFFFFF']],
        'fill' => ['fillType' => Fill::FILL_SOLID, 'startColor' => ['argb' => 'FF4472C4']],
        'alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER]
    ]);
}

// Column widths
$wsHi->getColumnDimension('A')->setWidth(14);
$wsHi->getColumnDimension('B')->setWidth(12);
$wsHi->getColumnDimension('C')->setWidth(90);

// Freeze header
$wsHi->freezePane('A3');

$rowNum = 3;
$totalRows = $wsEn->getHighestRow();
$idx = 0;

for ($i = 2; $i <= $totalRows; $i++) {
    $dateVal = $wsEn->getCell('B' . $i)->getValue();
    $signVal = $wsEn->getCell('C' . $i)->getValue();
    $predVal = $wsEn->getCell('D' . $i)->getValue();

    if (!$dateVal || !$signVal || !$predVal) continue;

    $idx++;
    echo "[$idx/" . ($totalRows - 1) . "] Translating $signVal for $dateVal...\n";

    $hindiSign = $zodiacHindi[$signVal] ?? $signVal;
    $hindiSummary = translateToHindi($predVal, $claudeApiKey);

    if (!$hindiSummary) {
        echo "Warning: Translation failed for $signVal on $dateVal, skipping\n";
        continue;
    }

    $fillColor = ($idx % 2 == 0) ? 'FFF2F2F2' : 'FFFFFFFF';

    $wsHi->getCell('A' . $rowNum)->setValue($dateVal);
    $wsHi->getCell('B' . $rowNum)->setValue($hindiSign);
    $wsHi->getCell('C' . $rowNum)->setValue($hindiSummary);

    $wsHi->getStyle("A{$rowNum}:C{$rowNum}")->applyFromArray([
        'fill' => ['fillType' => Fill::FILL_SOLID, 'startColor' => ['argb' => $fillColor]],
        'alignment' => ['wrapText' => true, 'vertical' => Alignment::VERTICAL_TOP]
    ]);

    $rowNum++;
}

$writer = new Xlsx($wbHi);
$writer->save('daily_horoscope_hindi.xlsx');
echo "✅ Hindi file saved: daily_horoscope_hindi.xlsx (" . ($rowNum - 3) . " rows)\n";
