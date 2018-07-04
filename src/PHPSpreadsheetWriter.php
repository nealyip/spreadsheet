<?php
/**
 * Created by PhpStorm.
 * User: neal.yip
 * Date: 4/7/2018
 * Time: 11:22
 */

namespace Nealyip\Spreadsheet;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Writer\Xls;
use PhpOffice\PhpSpreadsheet\Writer\Csv;


class PHPSpreadsheetWriter implements Writer
{

    /**
     * @var string
     */
    private $_filename;

    /**
     * @var Spreadsheet
     */
    private $_spreadsheet;

    /**
     * @var string
     */
    private $_ext;

    /**
     * @var Worksheet
     */
    private $_current;

    /**
     * @var bool
     */
    protected $_download;

    /**
     * Writer constructor.
     *
     * @param string $filename
     * @param bool   $download
     *
     * @return static
     * @throws GenericException
     * @throws WriterWrongFileFormatException
     */
    public function setup($filename, $download = true)
    {
        $this->_ext = strtolower(pathinfo($filename, PATHINFO_EXTENSION));
        if (!in_array($this->_ext, ['csv', 'xlsx', 'xls'])) {
            throw new WriterWrongFileFormatException();
        }

        $this->_filename    = $filename;
        $this->_download    = $download;
        $this->_spreadsheet = new Spreadsheet();

        return $this;
    }

    /**
     * Write to buffer
     *
     * @param \Generator $generator
     * @param array      $headers The 1st row. If a 2D array is given, it will be used as multi-line headers
     * @param callable   $filter  Filter that applied to each row
     *
     * @return static
     * @throws GenericException
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function write(\Generator $generator, array $headers = [], callable $filter = null)
    {
        try {
            $beginRow = $this->_beforeWrite($headers);

            foreach ($generator as $k => $item) {
                if (is_callable($filter)) {
                    $filter($item);
                }
                if (method_exists($item, 'toArray')) {
                    $item = $item->toArray();
                }
                $this->_current->fromArray($this->_numberSafeToExcel($item), null, 'A' . ($k + 1 + $beginRow));
            }

            return $this->_afterWrite();
        } catch (\PHPExcel_Exception $e) {
            throw new GenericException($e);
        }
    }

    /**
     * Write array to buffer
     *
     * @param array $data
     * @param array $headers The 1st row. If a 2D array is given, it will be used as multi-line headers
     *
     * @return static
     * @throws GenericException
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function writeArray(array $data, array $headers = [])
    {
        try {

            $beginRow = $this->_beforeWrite($headers);

            foreach (array_values($data) as $k => $item) {
                $this->_current->fromArray($this->_numberSafeToExcel($item), null, 'A' . ($k + 1 + $beginRow));
            }

            return $this->_afterWrite();
        } catch (\PHPExcel_Exception $e) {
            throw new GenericException($e);
        }
    }

    /**
     * Use this sheet
     *
     * @param string $name
     * @param int    $index
     *
     * @return static
     * @throws GenericException
     */
    public function useSheet($name = null, $index = 0)
    {
        try {
            if (!is_null($name) && $this->_spreadsheet->sheetNameExists($name)) {
                $sheet = $this->_spreadsheet->setActiveSheetIndexByName($name);
            } else {
                $sheet = $this->_spreadsheet->setActiveSheetIndex($this->_ext !== 'csv' ? $index : 0);
            }

            if (!is_null($name)) {
                $sheet->setTitle($name);
            }

            return $this;
        } catch (\PhpOffice\PhpSpreadsheet\Exception $e) {
            throw new GenericException($e);
        }
    }

    /**
     * Create a new sheet and use
     *
     * @param string $name
     *
     * @return static
     * @throws GenericException
     */
    public function newSheet($name = null)
    {

        try {
            if ($this->_ext !== 'csv') {
                $new = $this->_spreadsheet->createSheet();
                $this->_spreadsheet->setActiveSheetIndex($this->_spreadsheet->getIndex($new));
            } else {
                $new = $this->_spreadsheet->getActiveSheet();
            }

            if (!is_null($name)) {
                $new->setTitle($name);
            }

            return $this;
        } catch (\PhpOffice\PhpSpreadsheet\Exception $e) {
            throw new GenericException($e);
        }
    }

    /**
     * Only PHPExcel writer support this action
     * Xlsx only
     *
     * @param array $cellLists eg, ['C1:D1', 'E1:F1']
     *
     * @return static
     * @throws GenericException
     */
    public function mergeCells(array $cellLists = [])
    {
        try {
            if ($this->_ext !== 'csv') {
                $current = $this->_spreadsheet->getActiveSheet();

                foreach ($cellLists as $cell) {
                    $current->mergeCells($cell);
                }
            }
            return $this;
        } catch (\PhpOffice\PhpSpreadsheet\Exception $e) {
            throw new GenericException($e);
        }
    }


    /**
     * @inheritdoc
     */
    public function save()
    {

        try {

            $this->_spreadsheet->setActiveSheetIndex(0);
            $writer = $this->_getWriter();
            if ($this->_download) {
                header('Content-Disposition: attachment; filename="' . $this->_filename . '"');
                $writer->save('php://output');
                exit;
            } else {
                $writer->save($this->_filename);
            }
        } catch (\PhpOffice\PhpSpreadsheet\Exception $e) {
            throw new GenericException($e);
        }
    }

    /**
     * Microsoft excel will auto case numeric string to number and make large number incorrectly displayed
     *
     * @param array $data
     *
     * @return array
     */
    private function _numberSafeToExcel(array $data)
    {
//        if ($this->_ext === 'csv') {
        foreach ($data as &$d) {
            if (is_numeric($d) && strlen((string)$d) > 11) {
                $d .= "\t";
            }
        }
        unset($d);
//        }
        return $data;
    }


    /**
     * @return Csv|Xls|Xlsx
     */
    private function _getWriter()
    {
        switch ($this->_ext) {
            case 'xlsx':
                if ($this->_download) {
                    header('Content-type: application/vnd.ms-excel');
                }
                return new Xlsx($this->_spreadsheet);
            case 'xls':
                if ($this->_download) {
                    header('Content-type: application/vnd.ms-excel');
                }
                return new Xls($this->_spreadsheet);
            default:
                if ($this->_download) {
                    header('Content-type: text/csv; UTF-8');
                }
                $csv = new Csv($this->_spreadsheet);
                $csv->setUseBOM(true);
                return $csv;
        }
    }

    /**
     * @param $headers
     *
     * @return int
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    protected function _beforeWrite($headers)
    {

        $this->_current = $this->_spreadsheet->getActiveSheet();

        $beginRow = $this->_current->getHighestRow() - 1;

        if (count($headers)) {
            reset($headers);
            if (!is_array(current($headers))) {
                $headers = [$headers];
            }
            foreach ($headers as $row) {
                $this->_current->fromArray($row, null, 'A' . ++$beginRow);
            }
        }

        return $beginRow;
    }


    /**
     * After write
     *
     * @return $this
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    protected function _afterWrite()
    {
        for ($i = 'A'; $i != $this->_current->getHighestColumn(); $i++) {
            $this->_current->getColumnDimension($i)->setAutoSize(true);
        }
        $this->_current->insertNewRowBefore($this->_current->getHighestRow() + 1, 1);

        return $this;
    }
}