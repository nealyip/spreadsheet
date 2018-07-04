<?php
/**
 * Created by PhpStorm.
 * User: neal.yip
 * Date: 29/6/2017
 * Time: 15:13
 */

namespace Nealyip\Spreadsheet;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

class PHPSpreadsheetReader implements Reader
{

    /**
     * @var string
     */
    protected $_ext;

    use ReaderConvertArrayTrait;

    /**
     * @param string $file
     *
     * @return Spreadsheet
     * @throws WriterWrongFileFormatException
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    protected function _phpSpreadsheet($file)
    {

        if (!in_array($this->_ext, ['csv', 'xlsx', 'xls'])) {
            throw new WriterWrongFileFormatException();
        }

        $type = IOFactory::identify($file);

        $reader = IOFactory::createReader($type);

        return $reader->load($file);
    }

    /**
     * @inheritdoc
     */
    public function read($file, $sheetIndex = 0, $extension = null)
    {

        $this->_ext = $extension;

        if (is_null($extension)) {
            $this->_ext = strtolower(pathinfo($file, PATHINFO_EXTENSION));
        }



        try {
            $phpSpreadsheet = $this->_phpSpreadsheet($file);

            $sheet = $phpSpreadsheet->getSheet($sheetIndex);

            foreach ($sheet->getRowIterator() as $iterator) {
                $results = [];
                foreach ($iterator->getCellIterator() as $cell) {
                    /**
                     * @var \PhpOffice\PhpSpreadsheet\Cell\Cell $cell
                     */
                    $results[] = $cell->getValue();
                }
                yield $results;
            }
        } catch (\PhpOffice\PhpSpreadsheet\Exception $e) {
            throw new GenericException($e);
        }
    }
}