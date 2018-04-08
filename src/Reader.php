<?php
/**
 * Created by PhpStorm.
 * User: neal.yip
 * Date: 29/6/2017
 * Time: 12:42
 */

namespace Nealyip\Spreadsheet;


interface Reader
{

    /**
     * Read file and return row Generator
     *
     * @param string $file      The file type is guessed by file extension, but for example, upload file, the file name is without extension
     * @param int    $sheetIndex
     * @param null   $extension The file extension is auto read by $file, but can be overridden by this parameter
     *
     * @return \Generator
     * @throws WriterWrongFileFormatException
     * @throws GenericException
     */
    public function read($file, $sheetIndex = 0, $extension = null);

    /**
     * Read file and convert to key value array
     *
     * @param string $file             File name
     * @param int    $sheetIndex       0
     * @param bool   $firstColIsHeader Is first header a header column
     * @param array  $columns          Custom column for the json
     * @param null   $extension        Force file extension
     *
     * @return array
     */
    public function toKeyValueArray($file, $sheetIndex = 0, $firstColIsHeader = true, $columns = [], $extension = null);

    /**
     * Read file and convert to key value json
     *
     * @param string $file             File name
     * @param int    $sheetIndex       0
     * @param bool   $firstColIsHeader Is first header a header column
     * @param array  $columns          Custom column for the json
     * @param null   $extension        Force file extension
     *
     * @return mixed
     */
    public function toJson($file, $sheetIndex = 0, $firstColIsHeader = true, $columns = [], $extension = null);
}