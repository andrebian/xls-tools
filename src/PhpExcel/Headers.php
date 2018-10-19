<?php

namespace Andrebian\XlsTools\PhpExcel;

use Andrebian\XlsTools\ColumnsHelper;
use PHPExcel;

/**
 * Class Headers
 * @package Andrebian\XlsTools\PhpExcel
 */
class Headers
{
    /**
     * @var PhpExcel
     */
    public $phpExcel;

    public function __construct(PHPExcel $phpExcel)
    {
        $this->phpExcel = $phpExcel;
    }

    /**
     * @param array $headers
     * @param int $sheetIndex
     * @return PHPExcel
     * @throws \PHPExcel_Exception
     */
    public function setHeaders(array $headers, $sheetIndex = 1)
    {
        $columnsHelper = new ColumnsHelper();
        $count = 0;
        foreach ($headers as $header) {
            $letter = $columnsHelper->getNameFromNumber($count);

            if (is_array($header)
                && isset($header['letter'])
                && !empty($header['letter'])
            ) {
                $letter = $header['letter'];
            }

            $this->phpExcel
                ->setActiveSheetIndex($sheetIndex)
                ->setCellValue($letter . '1', $header);

            $count++;
        }

        return $this->phpExcel;
    }
}
