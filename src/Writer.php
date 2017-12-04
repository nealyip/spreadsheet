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
     */
    public function write(\Generator $generator, array $headers = [], callable $filter = null);

    /**
     * Write array to buffer
     *
     * @param array       $data
     * @param array       $headers The 1st row
     *
     * @return static
     */
    public function writeArray(array $data, array $headers = []);

    /**
     * Use this sheet
     *
     * @param string $name
     * @param int $index
     *
     * @return static
     */
    public function useSheet($name = null, $index = 0);

    /**
     * Create a new sheet and use
     *
     * @param string $name
     *
     * @return static
     */
    public function newSheet($name = null);

    /**
     * Close writer and export to browser
     *
     * @return mixed
     */
    public function save();
}