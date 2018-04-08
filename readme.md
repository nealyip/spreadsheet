## Description ##
Spreadsheet read writer abstraction for box/spout and phpexcel, support Laravel

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
PHPExcel is used by default


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
            ->writeArray([['Tom','M','20'], ['Ann','F','24']], $headers);

    }
```



