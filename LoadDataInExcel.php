<?php
namespace moonland\phpexcel;

use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class LoadDataInExcel
{
    /** @var TableNavigator|null */
    public ?TableNavigator $tn = null;

    public array $colunWidth = [];

    public array $styleTitle = [
        'font' => [
            'bold' => true,
            'size' => 14
        ],
        'alignment' => [
            'horizontal' => Alignment::HORIZONTAL_CENTER,
        ],
    ];

    public array $classStyle = [];

    /**
     * LoadDataInExcel constructor.
     * @param Worksheet $sheet
     * @param int $x
     * @param int $y
     */
    public function __construct(Worksheet $sheet, int $x = 1, int $y = 1)
    {
        $this->tn = new TableNavigator($x, $y);
        $this->sheet = $sheet;
    }

    public ?Worksheet $sheet = null;


    public array $thStyle = [
        'font' => [
            'bold' => true
        ],
        'alignment' => [
            'horizontal' => Alignment::HORIZONTAL_CENTER,
            'vertical' => Alignment::VERTICAL_TOP
        ],
        'fill' => [
            'fillType' => Fill::FILL_SOLID,
            'color' => ['rgb' => 'EEEEEE']
        ],
    ];

    /**
     * @param string $tableTitle
     * @throws Exception
     */
    public function setTableTitle(string $tableTitle): void
    {
        $this
            ->sheet
            ->mergeCells([
                $this->tn->minX,
                $this->tn->minY,
                $this->tn->maxX,
                $this->tn->minY
            ]);

        $this
            ->sheet
            ->setCellValue(
                [$this->tn->minX, $this->tn->minY],
                $tableTitle
            );

        $this
            ->sheet
            ->getStyle([$this->tn->minX, $this->tn->minY])
            ->applyFromArray($this->styleTitle);
    }

    /**
     * @param TableTdCell|null $cell
     * @throws Exception
     */
    public function fillCell(?TableTdCell $cell): void
    {
        if (!$cell) {
            return;
        }
        if ($cell->colspan !== 1) {
            $this->tn->setColspan($cell->colspan);
            $this->sheet->mergeCells([
                $this->tn->x,
                $this->tn->y,
                $this->tn->x + $cell->colspan - 1,
                $this->tn->y + $cell->rowspan - 1
            ]);
        }
        if ($cell->rowspan !== 1) {
            $this->tn->setRowspan($cell->rowspan);
            $this->sheet->mergeCells([
                $this->tn->x,
                $this->tn->y,
                $this->tn->x + $cell->colspan - 1,
                $this->tn->y + $cell->rowspan - 1
            ]);
        }

        $classStyle = [];

        if ($cell->tag === 'th') {
            $style = $this->thStyle;
            if (!$cell->nowrap) {
                $style['alignment']['wrapText'] = true;
            }
            $classStyle[] = $style;
        }
        if ($cell->numberDecimals !== false) {
            if ($cell->numberDecimals === 0) {
                $numberFormat = NumberFormat::FORMAT_NUMBER;
            } else {
                $numberFormat = '#,##0.' . str_repeat('0', $cell->numberDecimals);
            }
            $this
                ->sheet
                ->getStyle([$this->tn->x, $this->tn->y])
                ->getNumberFormat()
                ->setFormatCode($numberFormat);
            $classStyle[] = [
                'alignment' => [
                    'horizontal' => $cell->horizontalAlign?:Alignment::HORIZONTAL_CENTER,
                ],
            ];
        }

        $this
            ->sheet
            ->setCellValue(
                [$this->tn->x,$this->tn->y],
                $cell->getValueForExcel($this->tn->x, $this->tn->y)
            );

        if ($cell->width) {
            $this->colunWidth[$this->tn->x] = $cell->width;
        }
        if ($cell->class) {
            $cellClass = $cell->class;
            if (!is_array($cellClass)) {
                $cellClass = [$cellClass];
            }
            foreach ($cellClass as $className) {
                if (!isset($this->classStyle[$className])) {
                    continue;
                }
                $classStyle[] = $this->classStyle[$className];
            }
        }

        if ($classStyle) {
            $classStyle = array_merge(...$classStyle);
            $this
                ->sheet
                ->getStyle([
                    $this->tn->x,
                    $this->tn->y,
                    $this->tn->x + $cell->colspan - 1,
                    $this->tn->y + $cell->rowspan - 1
                ])
                ->applyFromArray($classStyle);
        }
    }

    /**
     * @param TableTdCell[] $row
     * @throws Exception
     */
    public function fillRow(array &$row): void
    {
        /**
         * get last cell key
         */
        end($row);
        $lastKey = key($row);
        reset($row);
        foreach ($row as $cellKey => $cell) {
            if ($cell !== null) {
                $this->fillCell($cell);
            }
            if ($cellKey !== $lastKey) {
                $this->tn->nextCell();
            }
        }
    }

    public function setColumnsWidths(): void
    {
        foreach ($this->colunWidth as $column => $size) {
            $this
                ->sheet
                ->getColumnDimensionByColumn($column)
                ->setAutoSize(false)
                ->setWidth($size);
        }
    }
}
