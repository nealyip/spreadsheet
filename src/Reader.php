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
     */
    public function read($file, $sheetIndex = 0, $extension = null);

}