<?php

namespace App\Exports;

use Maatwebsite\Excel\Concerns\FromCollection;
use Maatwebsite\Excel\Concerns\WithHeadings;
use Maatwebsite\Excel\Concerns\WithMapping;
use Maatwebsite\Excel\Concerns\WithStyles;
use Maatwebsite\Excel\Concerns\ShouldAutoSize;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Fill;

class MonthlyResultsExport implements FromCollection, WithHeadings, WithMapping, WithStyles, ShouldAutoSize
{
    protected $monthlyRewards;

    public function __construct($monthlyRewards)
    {
        $this->monthlyRewards = $monthlyRewards;
    }

    public function collection()
    {
        return collect($this->monthlyRewards);
    }

    public function headings(): array
    {
        // Dinamik sınav başlıkları ve toplam ödül sütunu
        return array_merge(
            ['Adı Soyadı', 'Görevi', 'Bölge', 'Hastane'],
            array_unique(collect($this->monthlyRewards)->pluck('exams')->flatten()->pluck('exam.title')->toArray()),
            ['Toplam Ödül Hakedişi']
        );
    }

    public function map($reward): array
    {
        $exams = collect($reward['exams'])->pluck('score', 'exam.title')->toArray();
        $row = [
            $reward['user']->name . ' ' . $reward['user']->surname,
            $reward['user']->job_title,
            $reward['user']->region,
            $reward['user']->hospital,
        ];

        $total_reward = 0;

        // Sınav başlıklarına göre dinamik olarak puan ekliyoruz
        $exam_titles = array_unique(collect($this->monthlyRewards)->pluck('exams')->flatten()->pluck('exam.title')->toArray());

        foreach ($exam_titles as $title) {
            $score = $exams[$title] ?? 0; // Eğer o sınavdan alınan puan varsa, yoksa 0
            $row[] = $score;

            // Eğer puan 70 ve üzeriyse ödül hesapla (her 70 üzeri sınav için 150 TL)
            if ($score >= 70) {
                $total_reward += 150;
            }
        }

        // Toplam ödülü satıra ekliyoruz
        $row[] = $total_reward;

        return $row;
    }

    public function styles(Worksheet $sheet)
    {
        // Başlıklar için stil
        $sheet->getStyle('A1:' . $sheet->getHighestColumn() . '1')->applyFromArray([
            'font' => [
                'bold' => true,
                'color' => ['rgb' => 'FFFFFF'],
            ],
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => ['rgb' => '4F81BD'], // Mavi renk başlık arka planı
            ],
        ]);

        // Sınav sonuçları renkleri (dinamik olarak renklendirme yapabiliriz)
        foreach ($sheet->getRowIterator(2) as $row) {
            foreach ($sheet->getColumnIterator('E', $sheet->getHighestColumn()) as $column) {
                $cellCoordinate = $column->getColumnIndex() . $row->getRowIndex();
                $cellValue = $sheet->getCell($cellCoordinate)->getValue();

                if (is_numeric($cellValue) && $cellValue >= 70) {
                    $sheet->getStyle($cellCoordinate)->applyFromArray([
                        'fill' => [
                            'fillType' => Fill::FILL_SOLID,
                            'startColor' => ['rgb' => '00FF00'], // Yeşil renk
                        ],
                    ]);
                }
            }
        }

        // Hücre kenarları için stil
        $sheet->getStyle('A1:' . $sheet->getHighestColumn() . $sheet->getHighestRow())->applyFromArray([
            'borders' => [
                'allBorders' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                    'color' => ['rgb' => '000000'],
                ],
            ],
        ]);

        return [
            'A1:' . $sheet->getHighestColumn() . '1' => ['font' => ['bold' => true]], // Başlıkları kalın yap
        ];
    }
}
