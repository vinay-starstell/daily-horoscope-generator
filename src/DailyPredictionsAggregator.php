<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\Exception;

// Load environment variables
$dotenv = new Symfony\Component\Dotenv\Dotenv();
$dotenv->usePutenv()->load(__DIR__ . '/../.env');

$authToken = $_ENV['DIVINEAPI_AUTH_TOKEN'] ?? null;
$apiKey = $_ENV['DIVINEAPI_KEY'] ?? null;
$gmailAddress = $_ENV['GMAIL_ADDRESS'] ?? null;
$gmailPassword = $_ENV['GMAIL_APP_PASSWORD'] ?? null;

if (!$authToken || !$apiKey || !$gmailAddress || !$gmailPassword) {
    die("Error: Missing required environment variables.\n");
}

function fetchHoroscopeData($date, $sign, $authToken, $apiKey)
{
    $url = "https://astroapi-5.divineapi.com/api/v2/daily-horoscope-custom";

    $headers = [
        "Authorization: " . $authToken,
    ];

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

// Start from today, fetch next 31 days
$today = new DateTime();
$startDate = clone $today;
$endDate = (clone $today)->modify('+31 days');

echo "Attempting to fetch data from " . $startDate->format('Y-m-d') . " to " . $endDate->format('Y-m-d') . "\n";

// Initialize Excel file
$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle("Horoscope Data");
$sheet->setCellValue('A1', 'No');
$sheet->setCellValue('B1', 'Date');
$sheet->setCellValue('C1', 'Sign');
$sheet->setCellValue('D1', 'Prediction');

// Zodiac signs
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
        $data = fetchHoroscopeData($currentDate, $sign, $authToken, $apiKey);
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
    die("Error: No data found for the requested date range.\n");
}

// Add summary sheet
$summarySheet = $spreadsheet->createSheet();
$summarySheet->setTitle("Summary");
$summarySheet->setCellValue('A1', 'Data Summary');
$summarySheet->setCellValue('A2', 'Date Range');
$summarySheet->setCellValue('B2', min($datesWithData) . ' to ' . max($datesWithData));
$summarySheet->setCellValue('A3', 'Total Entries');
$summarySheet->setCellValue('B3', $index - 1);
$summarySheet->setCellValue('A4', 'Dates with Data');
$summarySheet->setCellValue('B4', count($datesWithData));
$summarySheet->setCellValue('A5', 'Missing Dates');
$summarySheet->setCellValue('B5', count($datesWithoutData));

if (!empty($datesWithoutData)) {
    $summarySheet->setCellValue('A6', 'Missing Date List');
    $summarySheet->setCellValue('B6', implode(", ", $datesWithoutData));
}

// Save Excel file
$filename = 'daily_horoscope.xlsx';
$writer = new Xlsx($spreadsheet);
$writer->save($filename);

echo "✅ Horoscope data saved to $filename\n";
echo "📊 Data Range: " . min($datesWithData) . " to " . max($datesWithData) . "\n";
echo "📈 Total Entries: " . ($index - 1) . "\n";
if (!empty($datesWithoutData)) {
    echo "⚠️  Missing " . count($datesWithoutData) . " date(s): " . implode(", ", $datesWithoutData) . "\n";
}

// Send email with summary
$summary = "📊 Daily Horoscope Data Generated\n\n";
$summary .= "📅 Data Range: " . min($datesWithData) . " to " . max($datesWithData) . "\n";
$summary .= "📈 Total Entries: " . ($index - 1) . "\n";
if (!empty($datesWithoutData)) {
    $summary .= "⚠️ Missing Dates (" . count($datesWithoutData) . "): " . implode(", ", $datesWithoutData) . "\n";
}
$summary .= "\nPlease check the Summary sheet in the attached file for details.";

echo "Sending email to $gmailAddress...\n";

$mail = new PHPMailer(true);

try {
    $mail->isSMTP();
    $mail->Host = 'smtp.gmail.com';
    $mail->SMTPAuth = true;
    $mail->Username = $gmailAddress;
    $mail->Password = $gmailPassword;
    $mail->SMTPSecure = PHPMailer::ENCRYPTION_STARTTLS;
    $mail->Port = 587;

    $mail->setFrom($gmailAddress, 'StarsTell Horoscope Generator');
    $mail->addAddress($gmailAddress);
    $mail->Subject = 'Daily Horoscope - ' . date('Y-m-d');
    $mail->Body = $summary;

    $mail->addAttachment($filename);
    $mail->send();
    echo "✅ Email sent successfully\n";
} catch (Exception $e) {
    die("Email failed: {$mail->ErrorInfo}\n");
}