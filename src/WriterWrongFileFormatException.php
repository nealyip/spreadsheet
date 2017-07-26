<?php
/**
 * Created by PhpStorm.
 * User: neal.yip
 * Date: 11/4/2017
 * Time: 15:12
 */

namespace Nealyip\Spreadsheet;


class WriterWrongFileFormatException extends \Exception
{

    public function __construct($message = '')
    {
        parent::__construct($message ?: 'Export file not support');
    }

}