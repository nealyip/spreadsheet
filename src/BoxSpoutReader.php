<?php
/**
 * Created by PhpStorm.
 * User: neal.yip
 * Date: 29/6/2017
 * Time: 12:43
 */

namespace Nealyip\Spreadsheet;

use Box\Spout\Reader\CSV\Sheet;
use Box\Spout\Reader\ReaderFactory;

use Box\Spout\Common\Exception as BoxException;

class BoxSpoutReader implements Reader
{

    /**
     * @var string
     */
    protected $_ext;

    /**
     * @var ReaderFactory
     */
    protected $_readerFactory;

    use ReaderConvertArrayTrait;

    public function __construct(ReaderFactory $readerFactory)
    {

        $this->_readerFactory = $readerFactory;

    }

    /**
     * @param string $file
     *
     * @return \Box\Spout\Reader\ReaderInterface
     * @throws WriterWrongFileFormatException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Common\Exception\IOException
     */
    protected function _boxSpout($file)
    {

        if (!in_array($this->_ext, ['csv', 'xlsx', 'xls'])) {
            throw new WriterWrongFileFormatException();
        }

        $reader = $this->_readerFactory->create($this->_ext);
        $reader->open($file);

        return $reader;
    }

    /**
     * @inheritdoc
     */
    public function read($file, $sheetIndex = 0, $extension = null)
    {

        $this->_ext = $extension;

        if (is_null($extension)) {
            $this->_ext = strtolower(pathinfo($file, PATHINFO_EXTENSION));
        }

        try {
            $reader = $this->_boxSpout($file);

            $iterator = $reader->getSheetIterator();

            switch ($this->_ext) {
                case 'csv':
                    break;
                default:
                    while ($sheetIndex-- !== 0) {
                        $iterator->next();
                        if (!$iterator->valid()) {
                            throw new GenericException(new \Exception('Sheet not found'));
                        }
                    }
            }

            /**
             * @var Sheet $sheet
             */
            $iterator->rewind();
            $sheet = $iterator->current();
            /**
             * @todo this yield a iterator instead of a generator, allows rewind,
             *       but phpexcel implementation yield a generator, this should be changed to foreach yield instead
             */
            yield from $sheet->getRowIterator();
        } catch (BoxException\SpoutException $e) {
            throw new GenericException($e);
        }

    }
}