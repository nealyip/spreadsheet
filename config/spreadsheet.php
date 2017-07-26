<?php

return [
    'writer' => env('SPREADSHEET_WRITER', 'PHPExcel'), // either BoxSpout or PHPExcel (default)
    'reader' => env('SPREADSHEET_READER', 'PHPExcel'), // either BoxSpout or PHPExcel (default)
];