<?php
/**
 * Created by PhpStorm.
 * User: neal.yip
 * Date: 28/7/2017
 * Time: 16:45
 */

namespace Nealyip\Spreadsheet;


trait ReaderConvertArrayTrait
{

    /**
     * @inheritdoc
     */
    public function toKeyValueArray($file, $sheetIndex = 0, $firstColIsHeader = true, $columns = [], $extension = null)
    {
        $rows = $this->read($file, $sheetIndex, $extension);

        if ($firstColIsHeader) {

            if (!count($columns)) {
                $columns = $rows->current();
            }

            $rows->next();
        }

        $results = [];

        while ($rows->valid()) {
            $row       = $rows->current();
            $results[] = array_combine($columns, array_map('strval', array_slice($row, 0, count($columns))));
            $rows->next();
        }

        return $results;
    }

    /**
     * @inheritDoc
     */
    public function toJson($file, $sheetIndex = 0, $firstColIsHeader = true, $columns = [], $extension = null)
    {
        return json_encode($this->toKeyValueArray($file, $sheetIndex, $firstColIsHeader, $columns, $extension));
    }

}