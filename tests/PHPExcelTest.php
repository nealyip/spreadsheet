<?php
/**
 * Created by PhpStorm.
 * User: neal.yip
 * Date: 28/7/2017
 * Time: 16:53
 */

namespace Nealyip\Spreadsheet\Test;

use Nealyip\Spreadsheet\PHPExcelReader;

class PHPExcelTest extends TestCase
{

    protected $_datafile = __DIR__ . DIRECTORY_SEPARATOR . 'resources' . DIRECTORY_SEPARATOR . 'testdata.xlsx';

    public function testToKeyValueArray()
    {

        $phpexcel = new PHPExcelReader();

        $result = $phpexcel->toKeyValueArray($this->_datafile, 0, true);

        $this->assertEquals(
            [
                [
                    'name'   => 'Lee Ho',
                    'gender' => 'F',
                    'tel'    => '92121211'
                ],
                [
                    'name'   => 'chan tai man',
                    'gender' => 'M',
                    'tel'    => ''
                ]
            ], $result
        );

        $result = $phpexcel->toKeyValueArray($this->_datafile, 0, false, ['name', 'gender', 'tel']);

        $this->assertEquals(
            [
                [
                    'name'   => 'name',
                    'gender' => 'gender',
                    'tel'    => 'tel'
                ],
                [
                    'name'   => 'Lee Ho',
                    'gender' => 'F',
                    'tel'    => '92121211'
                ],
                [
                    'name'   => 'chan tai man',
                    'gender' => 'M',
                    'tel'    => ''
                ]
            ], $result
        );
    }

}