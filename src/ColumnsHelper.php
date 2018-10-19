<?php

namespace Andrebian\XlsTools;

/**
 * Class ColumnsHelper
 * @package Andrebian\XlsTools
 */
class ColumnsHelper
{
    /**
     * @param $number
     * @return string
     */
    public function getNameFromNumber($number)
    {
        $numeric = $number % 26;
        $letter = chr(65 + $numeric);
        $num2 = intval($number / 26);
        if ($num2 > 0) {
            return $this->getNameFromNumber($num2 - 1) . $letter;
        }

        return $letter;
    }
}
