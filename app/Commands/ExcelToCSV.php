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
                            {--config= : Path to a saved configuration file}
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

    /**
     * Configuration storage
     * 
     * @var array
     */
    protected array $config = [];

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
        // Check if config file was provided
        $configFilePath = $this->option('config');

        if ($configFilePath && File::exists($configFilePath)) {
            $this->info("Loading configuration from: {$configFilePath}");
            if ($this->loadConfigFromJson($configFilePath)) {
                $this->processWithSavedConfig();
                return;
            }
            $this->warn("Invalid configuration file. Proceeding with interactive mode.");
        }

        // Interactive mode
        $this->processInteractively();
    }

    /**
     * Process conversion using saved configuration
     */
    private function processWithSavedConfig(): void
    {
        try {
            $this->info('Processing with saved configuration...');

            // Load configuration values
            $inputFile = $this->config['input_file'];
            $outputFile = $this->config['output_file'];
            $delimiter = $this->config['delimiter'];
            $selectedSheet = $this->config['sheet'] ?? null;
            $skipRows = (int) ($this->config['skip_rows'] ?? 0);
            $includeHeaders = $this->config['include_headers'] ?? true;
            $encoding = $this->config['encoding'] ?? 'UTF-8';

            if (!File::exists($inputFile)) {
                $this->error("Input file not found: {$inputFile}");
                return;
            }

            $spreadsheet = $this->loadSpreadsheet($inputFile, $selectedSheet);

            // Apply column removals if specified
            if (!empty($this->config['columns_to_remove'])) {
                $this->applyColumnRemovals($spreadsheet, $this->config['columns_to_remove']);
            }

            // Apply value mappings if specified
            if (!empty($this->config['value_mappings'])) {
                $this->applyValueMappings($spreadsheet, $this->config['value_mappings']);
            }

            $this->exportToCsv($spreadsheet, $outputFile, $delimiter, $encoding, $skipRows, $includeHeaders);
        } catch (SpreadsheetReaderException $e) {
            $this->error("Error loading Excel file: " . $e->getMessage());
        } catch (\Exception $e) {
            $this->error("Error: " . $e->getMessage());
            if ($this->option('verbose')) {
                $this->error($e->getTraceAsString());
            }
        }
    }

    /**
     * Process conversion interactively
     */
    private function processInteractively(): void
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

        // Store basic configuration
        $this->config = [
            'input_file' => $inputFile,
            'output_file' => $outputFile,
            'delimiter' => $delimiter,
            'sheet' => $selectedSheet,
            'skip_rows' => $skipRows,
            'include_headers' => $includeHeaders,
            'encoding' => $encoding,
            'columns_to_remove' => [],
            'value_mappings' => [],
        ];

        try {
            $this->info('Loading Excel file...');
            $spreadsheet = $this->loadSpreadsheet($inputFile, $selectedSheet);

            $columnsToRemove = [];
            if ($this->confirm('Do you want to modify columns before exporting?', false)) {
                [$spreadsheet, $columnsToRemove] = $this->handleColumnChanges($spreadsheet);
                $this->config['columns_to_remove'] = $columnsToRemove;
            }

            $valueMappings = [];
            if ($this->confirm('Do you want to map values for any columns?', false)) {
                [$spreadsheet, $valueMappings] = $this->handleMapping($spreadsheet);
                $this->config['value_mappings'] = $valueMappings;
            }

            $this->exportToCsv($spreadsheet, $outputFile, $delimiter, $encoding, $skipRows, $includeHeaders);

            // Ask if the user wants to save this configuration
            if ($this->confirm('Would you like to save this configuration for future use?', true)) {
                $this->saveConfigToJson();
            }
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
     * Save current configuration to a JSON file
     */
    private function saveConfigToJson(): void
    {
        $configName = $this->ask('Enter a name for this configuration', 'excel_config_' . date('YmdHis'));
        $configName = str_ends_with($configName, '.json') ? $configName : $configName . '.json';

        try {
            $jsonContent = json_encode($this->config, JSON_PRETTY_PRINT);
            File::put($configName, $jsonContent);
            $this->info("Configuration saved to: {$configName}");
            $this->line("You can use it with: ./data-formatter-cli exceltocsv --config={$configName}");
        } catch (\Exception $e) {
            $this->error("Failed to save configuration: " . $e->getMessage());
        }
    }

    /**
     * Load configuration from a JSON file
     */
    private function loadConfigFromJson(string $path): bool
    {
        try {
            $content = File::get($path);
            $config = json_decode($content, true);

            if (json_last_error() !== JSON_ERROR_NONE) {
                return false;
            }

            // Validate required fields
            if (!isset($config['input_file'], $config['output_file'], $config['delimiter'])) {
                return false;
            }

            $this->config = $config;
            return true;
        } catch (\Exception $e) {
            return false;
        }
    }

    /**
     * Apply saved column removals to a spreadsheet
     */
    private function applyColumnRemovals(Spreadsheet $spreadsheet, array $columnsToRemove): void
    {
        if (empty($columnsToRemove)) {
            return;
        }

        $this->info('Applying column removals...');

        // Sort in descending order to avoid index shifting issues
        rsort($columnsToRemove);

        foreach ($columnsToRemove as $index) {
            $spreadsheet->getActiveSheet()->removeColumnByIndex($index + 1);
        }
    }

    /**
     * Apply saved value mappings to a spreadsheet
     */
    private function applyValueMappings(Spreadsheet $spreadsheet, array $valueMappings): void
    {
        if (empty($valueMappings)) {
            return;
        }

        $this->info('Applying value mappings...');

        foreach ($valueMappings as $mapping) {
            $columnIndex = $mapping['column_index'];
            $mappings = $mapping['mappings'];

            $colName = Coordinate::stringFromColumnIndex($columnIndex + 1);

            foreach ($spreadsheet->getActiveSheet()->toArray() as $rowIndex => $row) {
                if ($rowIndex === 0) {
                    continue;
                }

                $cellAddress = $colName . ($rowIndex + 1);
                $cell = $spreadsheet->getActiveSheet()->getCell($cellAddress);
                $cellValue = $cell->getValue();

                // Handle boolean values
                if ($cell->getDataType() === 'b') {
                    $cellValue = $cellValue ? 'TRUE' : 'FALSE';
                }

                if (array_key_exists($cellValue, $mappings)) {
                    $spreadsheet->getActiveSheet()->setCellValue($cellAddress, $mappings[$cellValue]);
                }
            }
        }
    }

    /**
     * Handle column mapping and return the mappings for potential storage
     */
    private function handleMapping(Spreadsheet $spreadsheet): array
    {
        $headers = $spreadsheet->getActiveSheet()->toArray()[0];
        $this->info('Starting column value mapping...');

        $valueMappings = [];

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

                // Store this mapping configuration
                $valueMappings[] = [
                    'column_index' => $col,
                    'column_name' => $header,
                    'mappings' => $newValues
                ];

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

        return [$spreadsheet, $valueMappings];
    }

    /**
     * Handle column modifications and return the columns to remove for potential storage
     */
    private function handleColumnChanges(Spreadsheet $spreadsheet): array
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

        return [$spreadsheet, $toRemove];
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
