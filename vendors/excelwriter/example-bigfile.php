<?php

set_time_limit(300);
error_reporting(E_ALL ^ E_NOTICE);
ini_set('display_errors', 0);
//$start = microtime_float();

require_once "class.writeexcel_workbookbig.inc.php";
require_once "class.writeexcel_worksheet.inc.php";

$fname = dirname(__FILE__) . "/bigfile.xls";
$workbook = new writeexcel_workbookbig($fname);
$worksheet = $workbook->addworksheet();

$worksheet->set_column(0, 20, 18);

for ($col=0;$col<20;$col++) {
    for ($row=0;$row<10000;$row++) {
        $worksheet->write($row, $col, "ROW:$row COL:$col");
    }
}

$workbook->close();

header("Content-Type: application/x-msexcel; name=\"example-bigfile.xls\"");
header("Content-Disposition: inline; filename=\"example-bigfile.xls\"");
$fh=fopen($fname, "rb");
fpassthru($fh);


$end = microtime_float();
file_put_contents( dirname(__FILE__) . '/' . 'patient_targeting_profilling.txt', sprintf( '[%s] Export took  %s seconds' . "\r\n", date('Y-m-d H:i:s'),$end - $start ), FILE_APPEND);
file_put_contents( dirname(__FILE__) .  '/' . 'patient_targeting_profilling.txt', sprintf('[%s] memory usage: %s ' . "\r\n" , date('Y-m-d H:i:s'), sprintf('%s M', memory_get_usage(true) / (1024*1024) ) ), FILE_APPEND);
file_put_contents( dirname(__FILE__) .  '/' . 'patient_targeting_profilling.txt', sprintf('[%s] memory peak usage: %s ' . "\r\n" , date('Y-m-d H:i:s'), sprintf('%s M', memory_get_peak_usage(true) / (1024*1024) ) ), FILE_APPEND);


/**
 * Simple function to replicate PHP 5 behaviour
 */
function microtime_float()
{
    list($usec, $sec) = explode(" ", microtime());
    return ((float)$usec + (float)$sec);
}

?>
