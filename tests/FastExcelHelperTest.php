<?php

declare(strict_types=1);

namespace avadim\FastExcelHelper;

use PHPUnit\Framework\TestCase;

final class FastExcelHelperTest extends TestCase
{
    public function testExcelHelper()
    {
        $this->assertEquals(1, Helper::rowNumber('1'));
        $this->assertEquals(0, Helper::rowNumber('a'));
        $this->assertEquals(2, Helper::rowNumber('az2'));
        $this->assertEquals(3, Helper::rowNumber('AZ3'));
        $this->assertEquals(4, Helper::rowNumber('$az4'));
        $this->assertEquals(5, Helper::rowNumber('az$5'));
        $this->assertEquals(6, Helper::rowNumber('$az$6'));
        $this->assertEquals(7, Helper::rowNumber('$az$7:BI24'));
        $this->assertEquals(8, Helper::rowNumber('$az$8:$BI24'));
        $this->assertEquals(9, Helper::rowNumber('$az$9:BI$24'));
        $this->assertEquals(10, Helper::rowNumber('$az$10:$BI$24'));

        $this->assertEquals(1, Helper::colNumber('A'));
        $this->assertEquals(52, Helper::colNumber('az'));
        $this->assertEquals(Helper::EXCEL_2007_MAX_COL, Helper::colNumber('XFD'));

        $this->assertEquals(0, Helper::colIndex('A'));
        $this->assertEquals(51, Helper::colIndex('az'));

        $this->assertEquals('', Helper::colLetter(0));
        $this->assertEquals('A', Helper::colLetter(1));
        $this->assertEquals('XFD', Helper::colLetter(Helper::EXCEL_2007_MAX_COL));
        $this->assertEquals('', Helper::colLetter(Helper::EXCEL_2007_MAX_COL + 1));

        $this->assertEquals([], Helper::rangeArray('a'));
        $arr = [
            'min_col_letter' => 'B',
            'min_col_num' => 2,
            'min_row_num' => 2,
            'min_col_abs' => '',
            'min_row_abs' => '',
            'min_cell' => 'B2',
            'min_cell_abs' => 'B2',
            'range' => 'B2:B2',
            'range_abs' => 'B2:B2',
        ];
        $this->assertEquals($arr, Helper::rangeArray('b2'));

        $arr = [
            'min_col_letter' => 'B',
            'min_col_num' => 2,
            'min_row_num' => 2,
            'min_col_abs' => '',
            'min_row_abs' => '$',
            'min_cell' => 'B2',
            'min_cell_abs' => 'B$2',
            'range' => 'B2:B2',
            'range_abs' => 'B$2:B$2',
        ];
        $this->assertEquals($arr, Helper::rangeArray('b$2'));

        $arr = [
            'min_col_letter' => 'B',
            'min_col_num' => 2,
            'min_row_num' => 2,
            'min_col_abs' => '',
            'min_row_abs' => '',
            'min_cell' => 'B2',
            'min_cell_abs' => 'B2',
            'max_col_letter' => 'D',
            'max_col_num' => 4,
            'max_row_num' => 4,
            'max_col_abs' => '',
            'max_row_abs' => '',
            'max_cell' => 'D4',
            'max_cell_abs' => 'D4',
            'range' => 'B2:D4',
            'range_abs' => 'B2:D4'
        ];
        $this->assertEquals($arr, Helper::rangeArray('b2:d4'));

        $arr['min_col_abs'] = '$';
        $arr['min_cell_abs'] = '$B2';
        $arr['range_abs'] = '$B2:D4';
        $this->assertEquals($arr, Helper::rangeArray('$b2:D4'));

        $arr['min_col_abs'] = '';
        $arr['min_row_abs'] = '$';
        $arr['min_cell_abs'] = 'B$2';
        $arr['max_row_abs'] = '$';
        $arr['max_cell_abs'] = 'D$4';
        $arr['range_abs'] = 'B$2:D$4';
        $this->assertEquals($arr, Helper::rangeArray('b$2:d$4'));

        $arr['min_col_abs'] = '$';
        $arr['min_row_abs'] = '$';
        $arr['min_cell_abs'] = '$B$2';
        $arr['max_col_abs'] = '$';
        $arr['max_row_abs'] = '$';
        $arr['max_cell_abs'] = '$D$4';
        $arr['range_abs'] = '$B$2:$D$4';
        $this->assertEquals($arr, Helper::rangeArray('$b$2:$d$4'));

        $this->assertEquals(['A', 'B', 'C'], Helper::colLetterRange([0, 1, 2]));
        $this->assertEquals(['', 'A', 'B'], Helper::colLetterRange([0, 1, 2], 1));
        $this->assertEquals(['B', 'E', 'F'], Helper::colLetterRange('B, E, F'));
        $this->assertEquals(['B', 'C', 'D', 'E', 'F'], Helper::colLetterRange('B-E, F'));
        $this->assertEquals(['B', 'C', 'D', 'E', 'F'], Helper::colLetterRange(['B-E', 'F']));
        $this->assertEquals(['B', 'C', 'D', 'E'], Helper::colLetterRange('B1-E8'));
        $this->assertEquals(['B:E'], Helper::colLetterRange('B1:E8'));
        $this->assertEquals(['B:E'], Helper::colLetterRange('$B1:E$8'));

        $this->assertEquals('AP38', Helper::cellAddress(38, 42));
        $this->assertEquals('$AP$38', Helper::cellAddress(38, 42, true));
        $this->assertEquals('$AP38', Helper::cellAddress(38, 42, true, false));
        $this->assertEquals('AP$38', Helper::cellAddress(38, 42, false, true));
        $this->assertEquals('AP38', Helper::cellAddress(38, 42, false, false));
        $this->assertEquals('XFD' . Helper::EXCEL_2007_MAX_ROW, Helper::cellAddress(Helper::EXCEL_2007_MAX_ROW, Helper::EXCEL_2007_MAX_COL));

        $this->assertEquals('#bf8f00', Helper::correctColor('#FFC000', -0.249977111117893));

        $this->assertEquals('C', Helper::colLetterNext(2));
        $this->assertEquals('C', Helper::colLetterNext('2'));
        $this->assertEquals('AC', Helper::colLetterNext('AB'));
        $this->assertEquals('FB34', Helper::colLetterNext('FA34'));

        $this->assertTrue(Helper::inRange('c3', 'B2:D4'));
        $this->assertTrue(Helper::inRange('B2', 'B2:D4'));
        $this->assertTrue(Helper::inRange('D4', 'B2:D4'));
        $this->assertFalse(Helper::inRange('A3', 'B2:D4'));

        $this->assertEquals('B2:D5', Helper::addToRange('b2', 'd5'));
        $this->assertEquals('B3:D3', Helper::addToRange('C3', 'B3:D3'));
        $this->assertEquals('B3:F5', Helper::addToRange('F5', 'B3:D3'));

        $this->assertEquals([0, 0, 0, 0], Helper::rangeRelOffsets('RC', $absolute));
        $this->assertEquals([1, 1, 1, 1], Helper::rangeRelOffsets('R1C1', $absolute));
        $this->assertEquals([true, true, true, true], $absolute);
        $this->assertEquals([-1, 1, 1, 3], Helper::rangeRelOffsets('R[-1]C1:r1c[3]', $absolute));
        $this->assertEquals([false, true, true, false], $absolute);

        $this->assertEquals('D5:D5', Helper::addToRange('d5', 'RC:RC'));
        $this->assertEquals('D5:D7', Helper::addToRange('d5', 'RC:R2C'));
        $this->assertEquals('D4:G7', Helper::addToRange('d5', 'R[-1]C:R2C3'));

        $this->assertEquals('R[5]C[2]', Helper::A1toRC('D8', 'b3'));
        $this->assertEquals('RC:R[5]C[2]', Helper::A1toRC('B3:D8', 'b3'));
        $this->assertEquals('R[-1]C:R[5]C[2]', Helper::A1toRC('B2:D8', 'b3'));

        $this->assertEquals('R8C4', Helper::A1toRC('$D$8'));
        $this->assertEquals('R[-2]C1', Helper::A1toRC('$A1', 'c3'));
        $this->assertEquals('R1C[-2]', Helper::A1toRC('A$1', 'c3'));
        $this->assertEquals('R1C1', Helper::A1toRC('$A$1', 'c3'));
        $this->assertEquals('R[-2]C[-2]', Helper::A1toRC('A1', 'c3'));

        $this->assertEquals('$B$5', Helper::RCtoA1('R5C2'));
        $this->assertEquals('$B$5', Helper::RCtoA1('R5C2', 'b3'));
        $this->assertEquals('B3:$B$5', Helper::RCtoA1('RC:R5C2', 'b3'));
        $this->assertEquals('B2:$B$5', Helper::RCtoA1('R[-1]C:R5C2', 'b3'));
        $this->assertEquals('$A2:$B8', Helper::RCtoA1('R[-1]C1:R[5]C2', 'b3'));

    }

}

