<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\Exception;

// Load environment variables
$dotenv = new Symfony\Component\Dotenv\Dotenv();
$dotenv->usePutenv()->load(__DIR__ . '/../.env');

// Get credentials from environment variables
$authToken = $_ENV['DIVINEAPI_AUTH_TOKEN'] ?? null;
$apiKey = $_ENV['DIVINEAPI_KEY'] ?? null;
$gmailAddress = $_ENV['GMAIL_ADDRESS'] ?? null;
$gmailPassword = $_ENV['GMAIL_APP_PASSWORD'] ?? null;

// Validate credentials
if (!$authToken || !$apiKey || !$gmailAddress || !$gmailPassword) {
    die("Error: Missing required environment variables. Check .env file.\n");
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
        echo "Error: API returned HTTP $httpCode for $sign on " . $date->format('Y-m-d') . "\n";
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

// Calculate dynamic date range
// Run on 20th: fetch from 20th of current month to 20th of next month
$today = new DateTime();
$currentDay = (int)$today->format('d');

// If script runs between 20-30 of month, fetch current month 20th to next month 20th
// If script runs before 20th, fetch previous month 20th to current month 20th
if ($currentDay >= 20) {
    $startDate = new DateTime($today->format('Y-m-20'));
    $endDate = (clone $startDate)->modify('+1 month');
} else {
    $startDate = (clone $today)->modify('-1 month')->format('Y-m-20');
    $startDate = new DateTime($startDate);
    $endDate = new DateTime($today->format('Y-m-20'));
}

echo "Fetching horoscope data from " . $startDate->format('Y-m-d') . " to " . $endDate->format('Y-m-d') . "\n";

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
$errorCount = 0;

while ($startDate <= $endDate) {
    foreach ($signs as $sign) {
        $data = fetchHoroscopeData($startDate, $sign, $authToken, $apiKey);
        if ($data) {
            $sheet->setCellValue("A$row", $index);
            $sheet->setCellValue("B$row", $data['Date']);
            $sheet->setCellValue("C$row", $data['Sign']);
            $sheet->setCellValue("D$row", $data['Prediction']);
            $row++;
            $index++;
        } else {
            $errorCount++;
        }
    }
    $startDate->modify('+1 day');
}

if ($errorCount > 0) {
    echo "Warning: $errorCount API calls failed. Check credentials and API limits.\n";
}

// Save Excel file
$filename = 'daily_horoscope.xlsx';
$writer = new Xlsx($spreadsheet);
$writer->save($filename);

echo "Horoscope data saved to $filename\n";

// Send email
echo "Sending email to $gmailAddress...\n";

$mail = new PHPMailer(true);

try {
    // SMTP configuration
    $mail->isSMTP();
    $mail->Host = 'smtp.gmail.com';
    $mail->SMTPAuth = true;
    $mail->Username = $gmailAddress;
    $mail->Password = $gmailPassword;
    $mail->SMTPSecure = PHPMailer::ENCRYPTION_STARTTLS;
    $mail->Port = 587;

    // Email details
    $mail->setFrom($gmailAddress, 'StarsTell Horoscope Generator');
    $mail->addAddress($gmailAddress);
    $mail->Subject = 'Daily Horoscope - ' . date('Y-m-d');
    $mail->Body = "Please find attached the daily horoscope file.";

    // Attach file
    $mail->addAttachment($filename);

    $mail->send();
    echo "Email sent successfully to $gmailAddress\n";
} catch (Exception $e) {
    die("Email failed: {$mail->ErrorInfo}\n");
}