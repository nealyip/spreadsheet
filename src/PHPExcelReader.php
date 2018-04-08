<?php
/**
 * Created by PhpStorm.
 * User: neal.yip
 * Date: 29/6/2017
 * Time: 15:13
 */

namespace Nealyip\Spreadsheet;

use PHPExcel_IOFactory;

class PHPExcelReader implements Reader
{

    /**
     * @var string
     */
    protected $_ext;

    use ReaderConvertArrayTrait;

    /**
     * @param string $file
     *
     * @return \PHPExcel
     * @throws WriterWrongFileFormatException
     * @throws \PHPExcel_Reader_Exception
     */
    protected function _phpExcel($file)
    {

        if (!in_array($this->_ext, ['csv', 'xlsx', 'xls'])) {
            throw new WriterWrongFileFormatException();
        }

        $type = PHPExcel_IOFactory::identify($file);

        $reader = PHPExcel_IOFactory::createReader($type);

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
            $phpExcel = $this->_phpExcel($file);

            $sheet = $phpExcel->getSheet($sheetIndex);

            foreach ($sheet->getRowIterator() as $iterator) {
                $results = [];
                foreach ($iterator->getCellIterator() as $cell) {
                    /**
                     * @var \PHPExcel_Cell $cell
                     */
                    $results[] = $cell->getValue();
                }
                yield $results;
            }
        } catch (\PHPExcel_Exception $e) {
            throw new GenericException($e);
        }
    }
}