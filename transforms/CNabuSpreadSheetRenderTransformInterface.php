<?php

/*  Copyright 2009-2011 Rafael Gutierrez Martinez
 *  Copyright 2012-2013 Welma WEB MKT LABS, S.L.
 *  Copyright 2014-2016 Where Ideas Simply Come True, S.L.
 *  Copyright 2017 nabu-3 Group
 *
 *  Licensed under the Apache License, Version 2.0 (the "License");
 *  you may not use this file except in compliance with the License.
 *  You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 *  Unless required by applicable law or agreed to in writing, software
 *  distributed under the License is distributed on an "AS IS" BASIS,
 *  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *  See the License for the specific language governing permissions and
 *  limitations under the License.
 */

namespace providers\nabu\spreadsheet\transforms;

use nabu\core\CNabuEngine;

use nabu\data\lang\CNabuLanguage;
use nabu\render\adapters\CNabuRenderTransformInterfaceAdapter;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

/**
 * Class to dump HTML rendered as PDF as HTTP response.
 * @author Rafael Gutierrez <rgutierrez@nabu-3.com>
 * @since 0.0.1
 * @version 0.0.1
 * @package \providers\nabu\pdf\renders
 */
class CNabuSpreadSheetRenderTransformInterface extends CNabuRenderTransformInterfaceAdapter
{
    public function transform($source)
    {
        $spreadsheet = new Spreadsheet();

        if (strlen($source) > 0 &&
            is_array($json = json_decode($source, true)) &&
            count($json) > 0
        ) {
            $sheet = $spreadsheet->setActiveSheetIndex(0);
            $headers = array_shift($json);
            $column = 1;
            foreach ($headers as $header) {
                $sheet->setCellValue($this->numberToColumn($column) . '1', $header);
                $column++;
            }
            if (count($json) > 0) {
                $row = 2;
                foreach ($json as $data) {
                    $column = 1;
                    foreach ($headers as $key => $header) {
                        if (array_key_exists($key, $data)) {
                            $sheet->setCellValue($this->numberToColumn($column) . $row, $data[$key]);
                        }
                        $column++;
                    }
                    $row++;
                }
            }
        }

        $writer = IOFactory::createWriter($spreadsheet, 'Xls');
        $writer->save('php://output');
    }

    private function numberToColumn(int $num)
    {
        $num--;
        $column = '';

        do {
            $rest = $num % 26;
            if (strlen($column) == 0) {
                $column = chr(65 + $rest);
            } else {
                $column = chr(64 + $rest) . $column;
            }
            $num = intdiv($num, 26);
        } while ($num > 0);

        return $column;
    }
}
