<?php
/**
 * Claude Code Routine: Monthly Horoscope Hindi Pipeline
 * Triggers: 1st of each month at 6 AM
 * Integrations: Gmail, Slack
 */

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Alignment;

// Load environment from .env
$dotenv = new Symfony\Component\Dotenv\Dotenv();
$dotenv->usePutenv()->load(__DIR__ . '/.env');

$divineAuthToken = $_ENV['DIVINEAPI_AUTH_TOKEN'] ?? null;
$divineApiKey = $_ENV['DIVINEAPI_KEY'] ?? null;
$claudeApiKey = $_ENV['CLAUDE_API_KEY'] ?? null;
$slackToken = $_ENV['SLACK_BOT_TOKEN'] ?? null;
$slackUserId = $_ENV['SLACK_USER_ID'] ?? null;

if (!$divineAuthToken || !$divineApiKey || !$claudeApiKey || !$slackToken || !$slackUserId) {
    die("Error: Missing required environment variables\n");
}

echo "=== Starting Horoscope Pipeline ===\n";

// ============================================================
// PART 1: FETCH HOROSCOPE DATA FROM API
// ============================================================

function fetchHoroscopeData($date, $sign, $authToken, $apiKey) {
    $url = "https://astroapi-5.divineapi.com/api/v2/daily-horoscope-custom";

    $headers = ["Authorization: " . $authToken];
    $postFields = [
        "api_key" => $apiKey,
        "sign" => $sign,
        "day" => $date->format('d'),
        "month" => $date->format('m'),
        "year" => $date->format('Y'),
        "tzone" => "5.5",
        "lan" => "en"
    ];

    $ch = curl_init();
    curl_setopt($ch, CURLOPT_URL, $url);
    curl_setopt($ch, CURLOPT_RETURNTRANSFER, 1);
    curl_setopt($ch, CURLOPT_POST, 1);
    curl_setopt($ch, CURLOPT_HTTPHEADER, $headers);
    curl_setopt($ch, CURLOPT_POSTFIELDS, $postFields);
    curl_setopt($ch, CURLOPT_TIMEOUT, 10);

    $response = curl_exec($ch);
    $httpCode = curl_getinfo($ch, CURLINFO_HTTP_CODE);
    curl_close($ch);

    if ($httpCode !== 200) {
        return null;
    }

    $data = json_decode($response, true);

    if (isset($data['success']) && $data['success'] == 1) {
        $prediction = $data['data']['prediction'] ?? [];

        $formattedPrediction = "";
        foreach ($prediction as $key => $value) {
            if (is_array($value)) {
                $formattedPrediction .= ucfirst($key) . " - " . implode(", ", $value) . "\n";
            } else {
                $formattedPrediction .= ucfirst($key) . " - " . $value . "\n";
            }
        }

        return [
            'Date' => $date->format('Y-m-d'),
            'Sign' => $sign,
            'Prediction' => trim($formattedPrediction)
        ];
    }

    return null;
}

// Calculate date range: today to +31 days
$today = new DateTime();
$startDate = clone $today;
$endDate = (clone $today)->modify('+31 days');

echo "Fetching data from " . $startDate->format('Y-m-d') . " to " . $endDate->format('Y-m-d') . "\n";

// Initialize English Excel
$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle("Horoscope Data");
$sheet->setCellValue('A1', 'No');
$sheet->setCellValue('B1', 'Date');
$sheet->setCellValue('C1', 'Sign');
$sheet->setCellValue('D1', 'Prediction');

$signs = ["Aries", "Taurus", "Gemini", "Cancer", "Leo", "Virgo", "Libra", "Scorpio", "Sagittarius", "Capricorn", "Aquarius", "Pisces"];

$row = 2;
$index = 1;
$datesWithData = [];
$datesWithoutData = [];
$currentDate = clone $startDate;

while ($currentDate <= $endDate) {
    $dateString = $currentDate->format('Y-m-d');
    $hasDataForDay = false;

    foreach ($signs as $sign) {
        $data = fetchHoroscopeData($currentDate, $sign, $divineAuthToken, $divineApiKey);
        if ($data) {
            $sheet->setCellValue("A$row", $index);
            $sheet->setCellValue("B$row", $data['Date']);
            $sheet->setCellValue("C$row", $data['Sign']);
            $sheet->setCellValue("D$row", $data['Prediction']);
            $row++;
            $index++;
            $hasDataForDay = true;
        }
    }

    if ($hasDataForDay) {
        $datesWithData[] = $dateString;
    } else {
        $datesWithoutData[] = $dateString;
    }

    $currentDate->modify('+1 day');
}

if (empty($datesWithData)) {
    die("Error: No data found\n");
}

$writer = new Xlsx($spreadsheet);
$writer->save('daily_horoscope.xlsx');
echo "✅ Generated: daily_horoscope.xlsx (" . ($index - 1) . " entries)\n";

// ============================================================
// PART 2: TRANSLATE TO HINDI USING CLAUDE
// ============================================================

function translateToHindi($englishPrediction, $claudeApiKey) {
    $prompt = "Translate this horoscope to Hindi. One paragraph, 4-5 lines, simple Hindi. Blend all aspects: Personal, Health, Profession, Emotions, Travel, Luck. NO bullet points.\n\nEnglish:\n$englishPrediction\n\nHindi only:";

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

// Load English horoscope
$wbEn = \PhpOffice\PhpSpreadsheet\IOFactory::load('daily_horoscope.xlsx');
$wsEn = $wbEn->getActiveSheet();

// Create Hindi horoscope
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

// Headers
$headers = ['A2' => 'तिथि', 'B2' => 'राशि', 'C2' => 'सारांश'];
foreach ($headers as $cell => $header) {
    $wsHi->getCell($cell)->setValue($header);
    $wsHi->getStyle($cell)->applyFromArray([
        'font' => ['bold' => true, 'color' => ['argb' => 'FFFFFFFF']],
        'fill' => ['fillType' => Fill::FILL_SOLID, 'startColor' => ['argb' => 'FF4472C4']]
    ]);
}

// Column widths
$wsHi->getColumnDimension('A')->setWidth(14);
$wsHi->getColumnDimension('B')->setWidth(12);
$wsHi->getColumnDimension('C')->setWidth(90);
$wsHi->freezePane('A3');

// Zodiac mapping
$zodiacHindi = [
    "Aries" => "मेष", "Taurus" => "वृषभ", "Gemini" => "मिथुन",
    "Cancer" => "कर्क", "Leo" => "सिंह", "Virgo" => "कन्या",
    "Libra" => "तुला", "Scorpio" => "वृश्चिक", "Sagittarius" => "धनु",
    "Capricorn" => "मकर", "Aquarius" => "कुंभ", "Pisces" => "मीन"
];

// Translate each row
$rowNum = 3;
$idx = 0;
$totalRows = $wsEn->getHighestRow();

for ($i = 2; $i <= $totalRows; $i++) {
    $dateVal = $wsEn->getCell('B' . $i)->getValue();
    $signVal = $wsEn->getCell('C' . $i)->getValue();
    $predVal = $wsEn->getCell('D' . $i)->getValue();

    if (!$dateVal || !$signVal || !$predVal) continue;

    $idx++;
    echo "Translating [$idx/" . ($totalRows - 1) . "] $signVal on $dateVal...\n";

    $hindiSign = $zodiacHindi[$signVal] ?? $signVal;
    $hindiSummary = translateToHindi($predVal, $claudeApiKey);

    if (!$hindiSummary) {
        echo "⚠️  Skip: $signVal on $dateVal\n";
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
echo "✅ Generated: daily_horoscope_hindi.xlsx (" . ($rowNum - 3) . " rows)\n";

// ============================================================
// PART 3: SEND TO SLACK
// ============================================================

$slackMessage = "📊 **दैनिक राशिफल तैयार है!**\n";
$slackMessage .= "📅 Date Range: " . min($datesWithData) . " to " . max($datesWithData) . "\n";
$slackMessage .= "📈 Total Entries: " . ($rowNum - 3) . "\n";
if (!empty($datesWithoutData)) {
    $slackMessage .= "⚠️ Missing " . count($datesWithoutData) . " date(s)\n";
}
$slackMessage .= "\nइस महीने का दैनिक राशिफल तैयार है। कृपया संलग्न फ़ाइल देखें। 🪐";

// Upload to Slack
$ch = curl_init('https://slack.com/api/files.upload');
curl_setopt($ch, CURLOPT_RETURNTRANSFER, 1);
curl_setopt($ch, CURLOPT_POST, 1);
curl_setopt($ch, CURLOPT_TIMEOUT, 30);
curl_setopt($ch, CURLOPT_HTTPHEADER, [
    "Authorization: Bearer $slackToken"
]);

$postData = [
    'channels' => $slackUserId,
    'initial_comment' => $slackMessage,
    'file' => new CURLFile('daily_horoscope_hindi.xlsx')
];

curl_setopt($ch, CURLOPT_POSTFIELDS, $postData);
$response = curl_exec($ch);
curl_close($ch);

$data = json_decode($response, true);

if ($data['ok']) {
    echo "✅ Sent to Slack successfully\n";
} else {
    echo "❌ Slack error: " . ($data['error'] ?? 'Unknown') . "\n";
}

echo "=== Pipeline Complete ===\n";