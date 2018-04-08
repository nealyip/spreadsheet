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
     * @param array      $headers The 1st row
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
     * @param array       $headers The 1st row
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
     * Close writer and export to browser
     *
     * @return mixed
     * @throws GenericException
     */
    public function save();
}