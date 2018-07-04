## Description ##
Spreadsheet read writer abstraction for box/spout and phpexcel, support Laravel

## Updates ##
v1.1.0 Add PHPSpreadsheet. PHPExcel was deprecated. Allow use of class with namespace for SPREADSHEET_WRITER and SPREADSHEET_READER from .env

## Installation ##
```
composer require nealyip/spreadsheet
```
Add this provider to config/app.php
```php
\Nealyip\Spreadsheet\SpreadsheetServiceProvider::class,
```

### Configuration ##

Publish config

```
php artisan vendor:publish --provider="Nealyip\Spreadsheet\SpreadsheetServiceProvider"
```

Simply change config/spreadsheet.php to select one spreadsheet data provider  
PHPSpreadsheet is used by default  

or configure on the .env file
```dotenv
SPREADSHEET_WRITER=PHPSpreadsheet  
SPREADSHEET_READER=BoxSpout  
```
Or if you implement your own Writer or Reader, you may use full class name here.  
```dotenv
SPREADSHEET_WRITER=App\Spreadsheet\CustomerWriter
```
be remember to implement  
```
Nealyip\Spreadsheet\Writer
```  

## How to use ##
Dependency Injection

### Reader ###
```php
use Nealyip\Spreadsheet\Reader;
class Sth{
    protected $_reader;

    public function __construct(Reader $reader) {
        $this->_reader = $reader;
    }

    public function readFile($filename){
        $data = $this->_reader->toKeyValueArray($filename);

    }
```

### Reader using generator ###
```php
use Nealyip\Spreadsheet\Reader;
class Sth{
    protected $_reader;

    public function __construct(Reader $reader) {
        $this->_reader = $reader;
    }

    public function readFile($filename){
        $data = $this->_reader->toKeyValueArray($filename);
        foreach ($this->_reader->read($filename) as $item){
            // $item is a row in array form        
        } 
    }

```

### Writer ###

```php
use Nealyip\Spreadsheet\Writer;
class Sth{
    protected $_writer;

    public function __construct(Writer $writer) {
        $this->_writer = $writer;
    }

    public function writeFile($filename){

        $headers = ['Name', 'Gender', 'Age'];

        $this->_writer
            ->setup("report.xlsx")
            ->useSheet('Report')
            ->writeArray([['Tom','M','20'], ['Ann','F','24']], $headers)
            ->save();

    }
```

### Writer using generator ###
```php
use Nealyip\Spreadsheet\Writer;
class Sth{
    protected $_writer;

    public function __construct(Writer $writer) {
        $this->_writer = $writer;
    }
    
    /**    
     * Data source from DB/API etc
     * 
     * @return \Generator
     */
    protected function _data(){
        $data = [['Tom','M','20'], ['Ann','F','24']];
        foreach ($data as $d) {
            yield $d;
        }
    }


    public function writeFile($filename){

        $headers = ['Name', 'Gender', 'Age'];

        $this->_writer
            ->setup("report.xlsx")
            ->useSheet('Report')
            ->write($this->_data(), $headers)
            ->save();
    }

```

### Write to local file ###
```php

use Nealyip\Spreadsheet\Writer;
class Sth{
    protected $_writer;

    public function __construct(Writer $writer) {
        $this->_writer = $writer;
    }
    
    /**    
     * Data source from DB/API etc
     * 
     * @return \Generator
     */
    protected function _data(){
        $data = [['Tom','M','20'], ['Ann','F','24']];
        foreach ($data as $d) {
            yield $d;
        }
    }


    public function writeFile($filename){

        $headers = ['Name', 'Gender', 'Age'];

        $this->_writer
            ->setup("report.xlsx", false)
            ->useSheet('Report')
            ->write($this->_data(), $headers)
            ->save();
    }

```


## Memory limit and execution timeout ##
If you have encounter memory exhaust problem, you may tune the memory limit by  
```php
ini_set('memory_limit', '1000M');
```
or for execution timeout  
```php
ini_set('max_execution_time', 300);
```