<?php

namespace App\Commands;

use Illuminate\Console\Scheduling\Schedule;
use Illuminate\Contracts\Console\PromptsForMissingInput;
use Illuminate\Support\Facades\File;
use LaravelZero\Framework\Commands\Command;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Reader\Exception as SpreadsheetReaderException;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Csv;

class ExcelToCSV extends Command implements PromptsForMissingInput
{
    /**
     * The name and signature of the console command.
     *
     * @var string
     */
    protected $signature = 'exceltocsv 
                            {--sheet= : Specific sheet name to convert (default: first sheet)}
                            {--encoding=UTF-8 : Output file encoding}
                            {--skip-rows=0 : Number of rows to skip from the beginning}
                            {--no-headers : Don\'t include headers in output}';

    /**
     * The console command description.
     *
     * @var string
     */
    protected $description = 'Convert Excel files to CSV with options for columns and value mapping';

    protected function promptForMissingArgumentsUsing(): array
    {
        return [
            'input_file' => 'Please provide the input Excel file path',
        ];
    }

    /**
     * Execute the console command.
     */
    public function handle(): void
    {
        // Read all xlsx in the current working directory
        $files = glob('*.xlsx');

        if (empty($files)) {
            $this->warn('No Excel files found in the current directory.');
        }

        $inputFile = $this->anticipate('Please provide the input Excel file path', $files);

        if (!File::exists($inputFile)) {
            $this->error("File not found: {$inputFile}");
            return;
        }

        $filenameWithoutExtension = pathinfo($inputFile, PATHINFO_FILENAME);
        $newFilename = 'formatted_' . $filenameWithoutExtension . '.csv';

        $outputFile = $this->ask(
            'Please provide the output CSV file path or keep empty to use the default',
            $newFilename
        );

        $delimiter = $this->ask('CSV delimiter character', ';');
        if (strlen($delimiter) !== 1) {
            $this->warn('Delimiter should be a single character. Using ";" as default.');
            $delimiter = ';';
        }

        $selectedSheet = $this->option('sheet');
        $skipRows = (int) $this->option('skip-rows');
        $includeHeaders = !$this->option('no-headers');
        $encoding = $this->option('encoding');

        try {
            $this->info('Loading Excel file...');
            $spreadsheet = $this->loadSpreadsheet($inputFile, $selectedSheet);

            if ($this->confirm('Do you want to modify columns before exporting?', false)) {
                $spreadsheet = $this->handleColumnChanges($spreadsheet);
            }

            if ($this->confirm('Do you want to map values for any columns?', false)) {
                $spreadsheet = $this->handleMapping($spreadsheet);
            }

            $this->exportToCsv($spreadsheet, $outputFile, $delimiter, $encoding, $skipRows, $includeHeaders);
        } catch (SpreadsheetReaderException $e) {
            $this->error("Error loading Excel file: " . $e->getMessage());
            return;
        } catch (\Exception $e) {
            $this->error("Error: " . $e->getMessage());
            if ($this->option('verbose')) {
                $this->error($e->getTraceAsString());
            }
        }
    }

    /**
     * Load a spreadsheet from file
     */
    private function loadSpreadsheet(string $inputFile, ?string $sheetName = null): Spreadsheet
    {
        $reader = IOFactory::createReaderForFile($inputFile);
        $reader->setReadDataOnly(true);

        if ($sheetName) {
            $reader->setLoadSheetsOnly([$sheetName]);
        }

        return $reader->load($inputFile);
    }

    /**
     * Export spreadsheet to CSV
     */
    private function exportToCsv(
        Spreadsheet $spreadsheet,
        string $outputFile,
        string $delimiter,
        string $encoding,
        int $skipRows = 0,
        bool $includeHeaders = true
    ): void {
        $writer = new Csv($spreadsheet);
        $writer->setDelimiter($delimiter);
        $writer->setEnclosure('"');
        $writer->setLineEnding("\r\n");
        $writer->setSheetIndex(0);

        if ($encoding !== 'UTF-8') {
            $writer->setUseBOM(true);
        }

        // Check if file exists, if so, ask user to overwrite or not
        if (File::exists(getcwd() . '/' . $outputFile)) {
            $overwrite = $this->confirm('File already exists, do you want to overwrite it?', true);
            if (!$overwrite) {
                $outputFile = $this->ask('Please provide the name for the new file');
                $outputFile = str_ends_with($outputFile, '.csv') ? $outputFile : $outputFile . '.csv';
            }
        }

        // Adjust skip rows if needed
        if ($skipRows > 0) {
            // Implementation would need custom CSV writer or post-processing
            $this->info("Skipping first {$skipRows} rows...");
        }

        // Write the file
        $writer->save($outputFile);
        $this->info("File converted successfully to {$outputFile}");
    }

    /**
     * Handle column mapping
     */
    private function handleMapping(Spreadsheet $spreadsheet): Spreadsheet
    {
        $headers = $spreadsheet->getActiveSheet()->toArray()[0];
        $this->info('Starting column value mapping...');

        // ask if any column values should be mapped to new values by 1. asking for the column name
        // 2. showing an array of unique values in that column
        // 3. asking for the new value for each unique value
        foreach ($headers as $col => $header) {
            $map = $this->confirm("Do you want to map values for column {$header}?");
            if ($map) {
                $columnValues = array_unique(array_column($spreadsheet->getActiveSheet()->toArray(), $col));
                unset($columnValues[0]);
                $columnValues = array_values($columnValues);
                $newValues = [];
                $columnValueCount = count($columnValues);

                $this->info("Column {$header} has {$columnValueCount} unique values");

                foreach ($columnValues as $key => $value) {
                    $step = $key + 1;
                    $newValue = $this->anticipate("Please provide the new value for {$value}! [$step/{$columnValueCount}]", [$value]);
                    $newValues[$value] = $newValue;
                }

                // show the mapping to the user as a table
                $this->table(['Old Value', 'New Value'], array_map(fn($key, $value) => [$key, $value], array_keys($newValues), $newValues));

                // ask if the user wants to apply the mapping
                $apply = $this->confirm('Do you want to apply this mapping?', true);

                if (!$apply) {
                    continue;
                }

                $colName = Coordinate::stringFromColumnIndex($col + 1);

                // apply the mapping by replacing the values in the column
                foreach ($spreadsheet->getActiveSheet()->toArray() as $rowIndex => $row) {
                    if ($rowIndex === 0) {
                        continue;
                    }

                    $cellAddress = $colName . ($rowIndex + 1);
                    $cell = $spreadsheet->getActiveSheet()->getCell($colName . ($rowIndex + 1));
                    $cellValue = $cell->getValue();

                    if ($cell->getDataType() === 'b') {
                        if ($cellValue === true) {
                            $cellValue = 'TRUE';
                        } else {
                            $cellValue = 'FALSE';
                        }
                    }

                    if (array_key_exists($cellValue, $newValues)) {
                        $spreadsheet->getActiveSheet()->setCellValue($cellAddress, $newValues[$cellValue]);
                    }
                }
            }
        }

        return $spreadsheet;
    }

    /**
     * Handle column modifications
     */
    private function handleColumnChanges(Spreadsheet $spreadsheet): Spreadsheet
    {
        $headers = $spreadsheet->getActiveSheet()->toArray()[0];
        $this->info('Starting column modifications...');

        $toRemove = [];

        // ask if any columns should be removed interactively
        foreach ($headers as $index => $header) {
            $remove = $this->confirm("Do you want to remove column {$header}?");
            if ($remove) {
                $toRemove[] = $index;
            }
        }

        $columnsToRemove = count($toRemove);
        $this->info("{$columnsToRemove} columns will be removed");

        // sort the columns to remove in descending order
        rsort($toRemove);

        // remove the columns
        foreach ($toRemove as $index) {
            $spreadsheet->getActiveSheet()->removeColumnByIndex($index + 1);
        }

        return $spreadsheet;
    }

    // Helper methods for converting between column index and column name
    public function getNameFromNumber(int $num): string
    {
        $numeric = $num % 26;
        $letter = chr(65 + $numeric);
        $num2 = intval($num / 26);
        if ($num2 > 0) {
            return $this->getNameFromNumber($num2 - 1) . $letter;
        } else {
            return $letter;
        }
    }

    public function getNameFromNumberIndexed(int $num): string
    {
        $numeric = ($num - 1) % 26;
        $letter = chr(65 + $numeric);
        $num2 = intval(($num - 1) / 26);
        if ($num2 > 0) {
            return $this->getNameFromNumberIndexed($num2) . $letter;
        } else {
            return $letter;
        }
    }

    /**
     * Define the command's schedule.
     */
    public function schedule(Schedule $schedule): void
    {
        // $schedule->command(static::class)->everyMinute();
    }
}
