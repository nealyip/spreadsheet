<?php

namespace Nealyip\Spreadsheet;

use Box\Spout\Common\Exception\SpoutException;
use Box\Spout\Writer\AbstractMultiSheetsWriter;
use Box\Spout\Writer\WriterFactory;
use Box\Spout\Writer\WriterInterface;

use \Box\Spout\Common\Exception as BoxException;

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
     * @var bool
     */
    protected $_download;

    /**
     * @inheritdoc
     */
    public function setup($filename, $download = true)
    {

        $this->_ext = strtolower(pathinfo($filename, PATHINFO_EXTENSION));
        if (!in_array($this->_ext, ['csv', 'xlsx', 'xls'])) {
            throw new WriterWrongFileFormatException();
        }
        $this->_filename = $filename;
        $this->_download = $download;
        $this->_box      = $this->_getWriter($filename, $download);

        return $this;
    }

    /**
     * @inheritdoc
     */
    public function write(\Generator $generator, array $headers = [], callable $filter = null)
    {
        try {

            if (count($headers)) {
                $this->_setHeaders($headers);
            }

            foreach ($generator as $k => $item) {
                if (is_callable($filter)) {
                    $filter($item);
                }
                if (method_exists($item, 'toArray')){
                    $item = $item->toArray();
                }
                $this->_box->addRow($this->_numberSafeToExcel($item));
            }

            return $this;
        } catch (SpoutException $e) {
            throw new GenericException($e);
        }
    }

    /**
     * @inheritDoc
     */
    public function writeArray(array $data, array $headers = [])
    {

        try {

            if (count($headers)) {
                $this->_setHeaders($headers);
            }

            foreach (array_values($data) as $k => $item) {
                $this->_box->addRow($this->_numberSafeToExcel($item));
            }

            return $this;
        } catch (SpoutException $e) {
            throw new GenericException($e);
        }
    }

    /**
     * @inheritDoc
     */
    public function useSheet($name = null, $index = 0)
    {

        try {

            if ($this->_box instanceof AbstractMultiSheetsWriter) {

                $sheet = $this->_box->getSheets()[$index]->setName($name);

                $this->_box->setCurrentSheet($sheet);

            }


            return $this;
        } catch (BoxException\SpoutException $e) {
            throw new GenericException($e);
        }

    }

    /**
     * @inheritDoc
     */
    public function newSheet($name = null)
    {

        try {
            if ($this->_box instanceof AbstractMultiSheetsWriter) {

                $this->_box->addNewSheetAndMakeItCurrent();

                if (!is_null($name)) {
                    $this->_box->getCurrentSheet()->setName($name);
                }
            }
            return $this;
        } catch (BoxException\SpoutException $e) {
            throw new GenericException($e);
        }
    }

    /**
     * Not support this action on BoxSpout Writer
     *
     * @inheritDoc
     */
    public function mergeCells(array $cellLists = [])
    {
        return $this;
    }


    /**
     * @inheritDoc
     */
    public function save()
    {
        $this->_box->close();

        if ($this->_download) {
            exit;
        }
    }


    /**
     * Set the 1st row
     *
     * @param array $headers
     *
     * @throws BoxException\IOException
     * @throws BoxException\SpoutException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
     */
    protected function _setHeaders($headers)
    {
        reset($headers);
        if (!is_array(current($headers))) {
            $headers = [$headers];
        }
        foreach ($headers as $row) {
            $this->_box->addRow($row);
        }
    }

    /**
     * Get writer
     *
     * @param string $filename
     *
     * @param bool   $download
     *
     * @return WriterInterface|AbstractMultiSheetsWriter
     * @throws WriterWrongFileFormatException
     * @throws GenericException
     */
    private function _getWriter($filename, $download)
    {
        $ext = $this->_ext;
        try {
            $writer = WriterFactory::create($ext);

            if ($download) {
                $writer->openToBrowser($filename);
            } else {
                $writer->openToFile($filename);
            }

            if ($ext === 'csv') {
                $writer->setShouldAddBOM(true);
            }
            return $writer;
        } catch (BoxException\UnsupportedTypeException $unsupportedTypeException) {
            throw new WriterWrongFileFormatException();
        } catch (BoxException\SpoutException $exception) {
            throw new GenericException($exception);
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
            } else if (is_numeric($d) && is_string($d)) {
                $d = floatval($d);
            }
        }
        unset($d);
//        }
        return $data;
    }
}