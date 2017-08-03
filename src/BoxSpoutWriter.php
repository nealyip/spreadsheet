<?php

namespace Nealyip\Spreadsheet;

use Box\Spout\Writer\AbstractMultiSheetsWriter;
use Box\Spout\Writer\WriterFactory;
use Box\Spout\Writer\WriterInterface;

class BoxSpoutWriter implements Writer
{
    /**
     * @var WriterInterface|AbstractMultiSheetsWriter
     */
    private $_box;

    /**
     * @var string
     */
    private $_ext;

    /**
     * @var string
     */
    private $_filename;

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
        $this->_box      = $this->_getWriter($filename);

        return $this;
    }

    /**
     * @inheritdoc
     */
    public function write(\Generator $generator, array $headers = [], callable $filter = null)
    {
        if (count($headers)){
            $this->_setHeaders($headers);
        }

        foreach ($generator as $k => $item) {
            if (is_callable($filter)) {
                $filter($item);
            }
            $this->_box->addRow($this->_numberSafeToExcel($item->toArray()));
        }

        return $this;
    }

    /**
     * @inheritDoc
     */
    public function writeArray(array $data, array $headers = [])
    {

        if (count($headers)){
            $this->_setHeaders($headers);
        }

        foreach (array_values($data) as $k => $item) {
            $this->_box->addRow($this->_numberSafeToExcel($item));
        }

        return $this;
    }

    /**
     * @inheritDoc
     */
    public function useSheet($name = null, $index = 0)
    {

        if ($this->_box instanceof AbstractMultiSheetsWriter) {

            $sheet = $this->_box->getSheets()[$index]->setName($name);

            $this->_box->setCurrentSheet($sheet);

        }


        return $this;
    }

    /**
     * @inheritDoc
     */
    public function newSheet($name = null)
    {

        if ($this->_box instanceof AbstractMultiSheetsWriter) {

            $this->_box->addNewSheetAndMakeItCurrent();

            if (!is_null($name)) {
                $this->_box->getCurrentSheet()->setName($name);
            }
        }

        return $this;
    }


    /**
     * @inheritDoc
     */
    public function save()
    {
        $this->_box->close();
        exit;
    }


    /**
     * Set the 1st row
     *
     * @param array $headers
     */
    private function _setHeaders($headers)
    {
        $this->_box->addRow($headers);
    }

    /**
     * Get writer
     *
     * @param string $filename
     *
     * @return WriterInterface|AbstractMultiSheetsWriter
     */
    private function _getWriter($filename)
    {
        $ext    = $this->_ext;
        $writer = WriterFactory::create($ext);
        $writer->openToBrowser($filename);
        if ($ext === 'csv') {
            $writer->setShouldAddBOM(true);
        }
        return $writer;
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