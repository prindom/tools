<?php

test('exceltocsv', function () {
    $excelToCSV = new \App\Commands\ExcelToCSV();
    // check if getNameFromNumberIndexed method called with 1 returns A
    $this->assertEquals('A', $excelToCSV->getNameFromNumberIndexed(1));
    // check if getNameFromNumberIndexed method called with 27 returns AA
    $this->assertEquals('AA', $excelToCSV->getNameFromNumberIndexed(27));
    // check if getNameFromNumberIndexed method called with 28 returns AB
    $this->assertEquals('AB', $excelToCSV->getNameFromNumberIndexed(28));

    // check if getNameFromNumber method called with 0 returns A
    $this->assertEquals('A', $excelToCSV->getNameFromNumber(0));
    // check if getNameFromNumber method called with 26 returns AA
    $this->assertEquals('AA', $excelToCSV->getNameFromNumber(26));
    // check if getNameFromNumber method called with 27 returns AB
    $this->assertEquals('AB', $excelToCSV->getNameFromNumber(27));

});
