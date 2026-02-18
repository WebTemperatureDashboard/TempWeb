<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Chart\Chart;
use PhpOffice\PhpSpreadsheet\Chart\Title;
use PhpOffice\PhpSpreadsheet\Chart\Legend;
use PhpOffice\PhpSpreadsheet\Chart\PlotArea;
use PhpOffice\PhpSpreadsheet\Chart\DataSeries;
use PhpOffice\PhpSpreadsheet\Chart\DataSeriesValues;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Border;

$channels = [
    ["id"=>3011291, "temp"=>1, "hum"=>2, "name"=>"Channel 1"],
    ["id"=>3016529, "temp"=>1, "hum"=>2, "name"=>"Channel 2"],
    ["id"=>345678,  "temp"=>1, "hum"=>2, "name"=>"Channel 3"]
];

$spreadsheet = new Spreadsheet();
$spreadsheet->removeSheetByIndex(0);

function createSheet($spreadsheet, $title, $fieldKey, $unit, $channels) {

    $sheet = $spreadsheet->createSheet();
    $sheet->setTitle($title);

    // ===== Header ISO =====
    $sheet->setCellValue("A1", "FACTORY ENVIRONMENT MONITORING REPORT");
    $sheet->setCellValue("A2", "Report Period: Last 7 Days");
    $sheet->setCellValue("A3", "Generated: ".date("Y-m-d H:i:s"));
    $sheet->getStyle("A1:A3")->getFont()->setBold(true);

    // ===== Table Header =====
    $sheet->setCellValue("A5", "DateTime");
    $col = "B";

    foreach($channels as $ch){
        $sheet->setCellValue($col."5", $ch["name"]." ($unit)");
        $col++;
    }

    $sheet->getStyle("A5:D5")->getFont()->setBold(true);
    $sheet->getStyle("A5:D5")->getFill()
        ->setFillType(Fill::FILL_SOLID)
        ->getStartColor()->setARGB("D9E1F2");

    $row = 6;

    foreach($channels as $index=>$ch){

        $url = "https://api.thingspeak.com/channels/".$ch["id"]."/feeds.json?days=7";
        $json = json_decode(file_get_contents($url), true);

        if(!isset($json["feeds"])) continue;

        foreach($json["feeds"] as $i=>$feed){

            if($index == 0){
                $sheet->setCellValue("A".$row,
                    date("Y-m-d H:i:s", strtotime($feed["created_at"])));
            }

            $value = $feed["field".$ch[$fieldKey]];
            $sheet->setCellValue(chr(66+$index).$row, $value);
            $row++;
        }

        $row = 6;
    }

    foreach(range('A','D') as $col){
        $sheet->getColumnDimension($col)->setAutoSize(true);
    }

    // ===== Chart =====
    $labels = [
        new DataSeriesValues('String', $title.'!$A$6:$A$'.($row-1), null, 10)
    ];

    $dataSeriesLabels = [];
    $dataSeriesValues = [];

    for($i=0;$i<count($channels);$i++){
        $col = chr(66+$i);
        $dataSeriesLabels[] =
            new DataSeriesValues('String', $title.'!$'.$col.'$5', null, 1);

        $dataSeriesValues[] =
            new DataSeriesValues('Number', $title.'!$'.$col.'$6:$'.$col.'$'.($row-1), null, 10);
    }

    $series = new DataSeries(
        DataSeries::TYPE_LINECHART,
        DataSeries::GROUPING_STANDARD,
        range(0,count($dataSeriesValues)-1),
        $dataSeriesLabels,
        $labels,
        $dataSeriesValues
    );

    $plotArea = new PlotArea(null, [$series]);
    $legend = new Legend(Legend::POSITION_RIGHT, null, false);

    $chart = new Chart(
        'chart_'.$title,
        new Title($title." Trend"),
        $legend,
        $plotArea
    );

    $chart->setTopLeftPosition("F5");
    $chart->setBottomRightPosition("N20");
    $sheet->addChart($chart);
}

createSheet($spreadsheet, "Temperature", "temp", "°C", $channels);
createSheet($spreadsheet, "Humidity", "hum", "%", $channels);

$filename = "Factory_Report_".date("Y-m-d").".xlsx";

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header("Content-Disposition: attachment;filename=\"$filename\"");
header('Cache-Control: max-age=0');

$writer = new Xlsx($spreadsheet);
$writer->setIncludeCharts(true);
$writer->save("php://output");
exit;
