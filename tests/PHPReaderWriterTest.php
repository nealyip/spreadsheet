<?php
/**
 * Created by PhpStorm.
 * User: neal.yip
 * Date: 10/4/2018
 * Time: 13:02
 */

namespace Nealyip\Spreadsheet\Test;

use Nealyip\Spreadsheet\PHPExcelWriter;

class PHPReaderWriterTest extends TestCase
{

    use ProtectedMethod;

    /**
     * @test
     * @throws \Nealyip\Spreadsheet\GenericException
     * @throws \Nealyip\Spreadsheet\WriterWrongFileFormatException
     * @throws \ReflectionException
     */
    public function beforeWriteHeader1DTest()
    {

        $phpexcel = new PHPExcelWriter();

        $phpexcel->setup('./test.xlsx', false);

        $beforeWrite = $this->_protectedMethod($phpexcel, '_beforeWrite', [['a', 'b', 'c']]);

        $this->assertEquals($beforeWrite, 1);

        /**
         * @var \PHPExcel_Worksheet $ws
         */
        $ws      = $this->_protectedProperty($phpexcel, '_current');
        $results = [];
        foreach ($ws->getRowIterator() as $iterator) {
            $result = [];
            foreach ($iterator->getCellIterator() as $cell) {
                /**
                 * @var \PHPExcel_Cell $cell
                 */
                $result[] = $cell->getValue();
            }
            $results[] = $result;
        }

        $this->assertEquals([['a', 'b', 'c']], $results);
    }

    /**
     * @test
     * @throws \Nealyip\Spreadsheet\GenericException
     * @throws \Nealyip\Spreadsheet\WriterWrongFileFormatException
     * @throws \ReflectionException
     */
    public function beforeWriteHeader2DTest()
    {

        $phpexcel = new PHPExcelWriter();

        $phpexcel->setup('./test.xlsx', false);

        $beforeWrite = $this->_protectedMethod($phpexcel, '_beforeWrite', [[['a', 'b', 'c'], ['d', 'e', '', 'f']]]);

        $this->assertEquals($beforeWrite, 2);

        /**
         * @var \PHPExcel_Worksheet $ws
         */
        $ws      = $this->_protectedProperty($phpexcel, '_current');
        $results = [];
        foreach ($ws->getRowIterator() as $iterator) {
            $result = [];
            foreach ($iterator->getCellIterator() as $cell) {
                /**
                 * @var \PHPExcel_Cell $cell
                 */
                $result[] = $cell->getValue();
            }
            $results[] = $result;
        }

        $this->assertEquals([['a', 'b', 'c', ''], ['d', 'e', '', 'f']], $results);
    }

    /**
     * @test
     * @throws \Nealyip\Spreadsheet\GenericException
     * @throws \Nealyip\Spreadsheet\WriterWrongFileFormatException
     * @throws \ReflectionException
     */
    public function beforeWriteHeader2DWithAssociativeKeysTest()
    {

        $phpexcel = new PHPExcelWriter();

        $phpexcel->setup('./test.xlsx', false);

        $beforeWrite = $this->_protectedMethod($phpexcel, '_beforeWrite', [[[1 => 'a', 7 => 'b', 3 => 'c'], ['d', 'e', '', 'f']]]);

        $this->assertEquals($beforeWrite, 2);

        /**
         * @var \PHPExcel_Worksheet $ws
         */
        $ws      = $this->_protectedProperty($phpexcel, '_current');
        $results = [];
        foreach ($ws->getRowIterator() as $iterator) {
            $result = [];
            foreach ($iterator->getCellIterator() as $cell) {
                /**
                 * @var \PHPExcel_Cell $cell
                 */
                $result[] = $cell->getValue();
            }
            $results[] = $result;
        }

        $this->assertEquals([['a', 'b', 'c', ''], ['d', 'e', '', 'f']], $results);
    }
}
