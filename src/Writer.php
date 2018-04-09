<?php

namespace Nealyip\Spreadsheet;

interface Writer
{

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
    public function setup($filename, $download = true);

    /**
     * Write to buffer
     *
     * @param \Generator $generator
     * @param array      $headers The 1st row. If a 2D array is given, it will be used as multi-line headers
     * @param callable   $filter  Filter that applied to each row
     *
     * @return static
     * @throws GenericException
     */
    public function write(\Generator $generator, array $headers = [], callable $filter = null);

    /**
     * Write array to buffer
     *
     * @param array       $data
     * @param array       $headers The 1st row. If a 2D array is given, it will be used as multi-line headers
     *
     * @return static
     * @throws GenericException
     */
    public function writeArray(array $data, array $headers = []);

    /**
     * Use this sheet
     *
     * @param string $name
     * @param int $index
     *
     * @return static
     * @throws GenericException
     */
    public function useSheet($name = null, $index = 0);

    /**
     * Create a new sheet and use
     *
     * @param string $name
     *
     * @return static
     * @throws GenericException
     */
    public function newSheet($name = null);

    /**
     * Only PHPExcel writer support this action
     * Xlsx only
     *
     * @param array $cellLists eg, ['C1:D1', 'E1:F1']
     *
     * @return static
     * @throws GenericException
     */
    public function mergeCells(array $cellLists = []);

    /**
     * Close writer and export to browser
     *
     * @return void exit()
     * @throws GenericException
     */
    public function save();
}