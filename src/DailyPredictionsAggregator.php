<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

function fetchHoroscopeData($date, $sign)
{
      $url = "https://astroapi-5.divineapi.com/api/v2/daily-horoscope-custom";

      $headers = [
            "Authorization: eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpc3MiOiJodHRwczovL2FzdHJvYXBpLTEuZGl2aW5lYXBpLmNvbS9hcGkvYXV0aC1hcGktdXNlciIsImlhdCI6MTczNzAyMjg3OCwibmJmIjoxNzM3MDIyODc4LCJqdGkiOiJEZnVySk5RNkVQWXJ5RHJuIiwic3ViIjoiMTY4MiIsInBydiI6ImU2ZTY0YmIwYjYxMjZkNzNjNmI5N2FmYzNiNDY0ZDk4NWY0NmM5ZDcifQ.NNctV0gztAlGSykDqE7Tu84G73PYC0mW3r0fmeYSgwo",  // Replace with your actual token
      ];

      $postFields = [
            "api_key" => '6a81681a7af700c6385d36577ebec359', // Replace with your API key
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

      $response = curl_exec($ch);
      curl_close($ch);

      $data = json_decode($response, true);

      if (isset($data['success']) && $data['success'] == 1) {
            $prediction = $data['data']['prediction'] ?? [];

            // Formatting predictions with tags
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

// Initialize Excel file
$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle("Horoscope Data");
$sheet->setCellValue('A1', 'No');
$sheet->setCellValue('B1', 'Date');
$sheet->setCellValue('C1', 'Sign');
$sheet->setCellValue('D1', 'Prediction');

// Zodiac signs list
$signs = ["Aries", "Taurus", "Gemini", "Cancer", "Leo", "Virgo", "Libra", "Scorpio", "Sagittarius", "Capricorn", "Aquarius", "Pisces"];
// $startDate = "new DateTime();" // Today's date
// $endDate = (clone $startDate)->modify('+30 days'); // Next one month

$startDate = new DateTime('2026-03-23'); // Set start date
$endDate = new DateTime('2026-04-23'); // Set end date

$row = 2;
$index = 1;

while ($startDate <= $endDate) {
      foreach ($signs as $sign) {
            $data = fetchHoroscopeData($startDate, $sign);
            if ($data) {
                  $sheet->setCellValue("A$row", $index);
                  $sheet->setCellValue("B$row", $data['Date']);
                  $sheet->setCellValue("C$row", $data['Sign']);
                  $sheet->setCellValue("D$row", $data['Prediction']);
                  $row++;
                  $index++;
            }
      }
      $startDate->modify('+1 day');
}

// Save Excel file
$writer = new Xlsx($spreadsheet);
$filename = 'daily_horoscope.xlsx';
$writer->save($filename);

echo "Horoscope data saved to $filename";