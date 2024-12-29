<?php

namespace App\Exports;

use App\Models\ExamResult;
use Maatwebsite\Excel\Concerns\FromCollection;
use Maatwebsite\Excel\Concerns\WithHeadings;
use Maatwebsite\Excel\Concerns\WithMapping;
use Maatwebsite\Excel\Concerns\WithStyles;
use Maatwebsite\Excel\Concerns\ShouldAutoSize;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Style\Fill;

class ExamResultsExport implements FromCollection, WithHeadings, WithMapping, WithStyles, ShouldAutoSize
{
    /**
     * Verileri toplar.
     */
    public function collection()
    {
        return ExamResult::with('user')->get();
    }

    /**
     * Excel başlıkları
     */
    public function headings(): array
    {
        return [
            'Ad',        // Name
            'Soyad',     // Surname
            'Görev',     // Job Title
            'Bölge',     // Region
            'Hastane',   // Hospital
            'Durum',     // Status (Başarılı / Başarısız)
            'Puan',      // Score
            'Ödül'       // Reward
        ];
    }

    /**
     * Excel'deki satırları nasıl sıralamak istediğinizi belirtir
     */
    public function map($result): array
    {
        return [
            $result->user->name,              // Ad
            $result->user->surname,           // Soyad
            $result->user->job_title,         // Görev
            $result->user->region,            // Bölge
            $result->user->hospital,          // Hastane
            $result->score >= 70 ? 'Başarılı' : 'Başarısız', // Durum
            $result->score,                   // Puan
            $result->reward                   // Ödül
        ];
    }

    /**
     * Excel hücre stillerini belirler
     */
    public function styles(Worksheet $sheet)
    {
        // Başlık satırı için stil
        $sheet->getStyle('A1:H1')->applyFromArray([
            'font' => [
                'bold' => true,
                'color' => ['rgb' => 'FFFFFF'],
            ],
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => ['rgb' => '4F81BD'], // Mavi renk başlık arka planı
            ],
        ]);

        // Her bir satır için stil
        foreach ($sheet->getRowIterator(2) as $row) {
            $scoreCell = 'G' . $row->getRowIndex(); // Puan hücresi
            $statusCell = 'F' . $row->getRowIndex(); // Durum hücresi
            
            // Puan renklendirmesi
            $cellValue = $sheet->getCell($scoreCell)->getValue();
            if (is_numeric($cellValue) && $cellValue >= 70) {
                $sheet->getStyle($scoreCell)->applyFromArray([
                    'fill' => [
                        'fillType' => Fill::FILL_SOLID,
                        'startColor' => ['rgb' => '00FF00'], // Yeşil renk
                    ],
                ]);
            }

            // Durum renklendirmesi
            if ($sheet->getCell($statusCell)->getValue() === 'Başarılı') {
                $sheet->getStyle($statusCell)->applyFromArray([
                    'fill' => [
                        'fillType' => Fill::FILL_SOLID,
                        'startColor' => ['rgb' => '00FF00'], // Yeşil renk
                    ],
                ]);
            } else {
                $sheet->getStyle($statusCell)->applyFromArray([
                    'fill' => [
                        'fillType' => Fill::FILL_SOLID,
                        'startColor' => ['rgb' => 'FF0000'], // Kırmızı renk
                    ],
                ]);
            }
        }

        // Hücre kenarları için stil
        $sheet->getStyle('A1:H100')->applyFromArray([
            'borders' => [
                'allBorders' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                    'color' => ['rgb' => '000000'],
                ],
            ],
        ]);
    }
}
