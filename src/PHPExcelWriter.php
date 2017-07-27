<?php

namespace Nealyip\Spreadsheet;


class PHPExcelWriter implements Writer
{

    /**
     * @var string
     */
    private $_filename;

    /**
     * @var \PHPExcel
     */
    private $_phpExcel;

    /**
     * @var string
     */
    private $_ext;

    /**
     * @var \PHPExcel_Worksheet
     */
    private $_current;

    /**
     * @inheritdoc
     */
    public function setup($filename)
    {
        $this->_ext = strtolower(pathinfo($filename, PATHINFO_EXTENSION));
        if (!in_array($this->_ext, ['csv', 'xlsx', 'xls'])) {
            throw new WriterWrongFileFormatException();
        }
        $this->_filename = $filename;
        $this->_phpExcel = new \PHPExcel();

        return $this;
    }

    /**
     * @inheritdoc
     */
    public function write(\Generator $generator, array $headers = [], callable $filter = null)
    {
        $beginRow = $this->_beforeWrite($headers);

        foreach ($generator as $k => $item) {
            if (is_callable($filter)) {
                $filter($item);
            }
            $this->_current->fromArray($this->_numberSafeToExcel($item->toArray()), null, 'A' . ($k + 1 + $beginRow));
        }

        return $this->_afterWrite();
    }

    /**
     * @inheritDoc
     */
    public function writeArray(array $data, array $headers = [])
    {
        $beginRow = $this->_beforeWrite($headers);

        foreach (array_values($data) as $k => $item) {
            $this->_current->fromArray($this->_numberSafeToExcel($item), null, 'A' . ($k + 1 + $beginRow));
        }

        return $this->_afterWrite();
    }

    /**
     * @inheritDoc
     */
    public function useSheet($name = null, $index = 0)
    {

        if (!is_null($name) && $this->_phpExcel->sheetNameExists($name)) {
            $sheet = $this->_phpExcel->setActiveSheetIndexByName($name);
        } else {
            $sheet = $this->_phpExcel->setActiveSheetIndex($this->_ext !== 'csv' ? $index : 0);
        }

        if (!is_null($name)) {
            $sheet->setTitle($name);
        }

        return $this;
    }

    /**
     * @inheritDoc
     */
    public function newSheet($name = null)
    {

        if ($this->_ext !== 'csv') {
            $new = $this->_phpExcel->createSheet();
            $this->_phpExcel->setActiveSheetIndex($this->_phpExcel->getIndex($new));
        } else {
            $new = $this->_phpExcel->getActiveSheet();
        }

        if (!is_null($name)) {
            $new->setTitle($name);
        }

        return $this;
    }


    /**
     * @inheritDoc
     */
    public function save()
    {
        $this->_phpExcel->setActiveSheetIndex(0);
        $writer = $this->_getWriter();
        header('Content-Disposition: attachment; filename="' . $this->_filename . '"');
        $writer->save('php://output');
        exit;
    }


    /**
     * Get writer
     *
     * @return \PHPExcel_Writer_CSV|\PHPExcel_Writer_Excel2007|\PHPExcel_Writer_Excel5
     */
    private function _getWriter()
    {
        switch ($this->_ext) {
            case 'xlsx':
                header('Content-type: application/vnd.ms-excel');
                return new \PHPExcel_Writer_Excel2007($this->_phpExcel);
            case 'xls':
                header('Content-type: application/vnd.ms-excel');
                return new \PHPExcel_Writer_Excel5($this->_phpExcel);
            default:
                header('Content-type: text/csv; UTF-8');
                $csv = new \PHPExcel_Writer_CSV($this->_phpExcel);
                $csv->setUseBOM(true);
                return $csv;
        }
    }

    /**
     * Before write
     *
     * @param array $headers
     *
     * @return int
     */
    protected function _beforeWrite($headers)
    {

        $this->_current = $this->_phpExcel->getActiveSheet();

        $beginRow = $this->_current->getHighestRow();

        $this->_current->fromArray($headers, null, 'A' . $beginRow);

        return $beginRow;
    }

    /**
     * After write
     *
     * @return $this
     */
    protected function _afterWrite()
    {
        for ($i = 'A'; $i != $this->_current->getHighestColumn(); $i++) {
            $this->_current->getColumnDimension($i)->setAutoSize(true);
        }
        $this->_current->insertNewRowBefore($this->_current->getHighestRow() + 1, 1);

        return $this;
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
}