<?php

return [
    'writer' => env('SPREADSHEET_WRITER', 'PHPSpreadsheet'), // either BoxSpout or PHPExcel (default)
    'reader' => env('SPREADSHEET_READER', 'PHPSpreadsheet'), // either BoxSpout or PHPExcel (default)
];