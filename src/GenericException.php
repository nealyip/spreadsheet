<?php
/**
 * Created by PhpStorm.
 * User: Neal
 * Date: 8/4/2018
 * Time: 5:58 PM
 */

namespace Nealyip\Spreadsheet;


use Throwable;

class GenericException extends \Exception
{

    /**
     * @var Throwable
     */
    protected $_exception;

    public function __construct(Throwable $previous = null)
    {

        parent::__construct($previous->getMessage(), $previous->getCode(), $previous);

        $this->_exception = $previous;
    }

    /**
     * @return Throwable
     */
    public function previous(){
        return $this->_exception;
    }
}