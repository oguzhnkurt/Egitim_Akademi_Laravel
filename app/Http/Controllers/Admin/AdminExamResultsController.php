<?php

namespace App\Http\Controllers\Admin;

use App\Http\Controllers\Controller;
use Illuminate\Http\Request;
use App\Models\Exam;
use App\Models\ExamResult;
use Illuminate\Support\Facades\Auth;
use App\Exports\ExamResultsExport;
use App\Exports\MonthlyResultsExport;
use Maatwebsite\Excel\Facades\Excel;
use Spipu\Html2Pdf\Html2Pdf;



class AdminExamResultsController extends Controller
{
    public function index(Request $request)
    {
        $exams = Exam::all();
        $results = ExamResult::with('user', 'exam')->get();

        return view('admin.exam_results.index', compact('exams', 'results'));
    }

    public function submit(Request $request, $examId)
    {
        $userId = Auth::id(); // Giriş yapmış kullanıcı ID'si
        $exam = Exam::find($examId); // Sınavı bul
        if (!$exam) {
            return redirect()->back()->with('error', 'Sınav bulunamadı.');
        }

        $questions = $exam->questions; // Sınavın soruları
        $score = 0; // Başlangıçta puan 0

        // Her bir soru için cevapları kontrol et
        foreach ($questions as $question) {
            $correctAnswer = $question->correct_option; // Doğru cevap
            $selectedAnswer = $request->input('questions.' . $question->id); // Kullanıcının seçtiği cevap

            // Eğer seçilen cevap doğru ise puanı artır
            if ($selectedAnswer == $correctAnswer) {
                $score += $question->max_score; // Sorunun maksimum puanı kadar puan ekle
            }
        }

        // Sonuçları veritabanına kaydet
        $examResult = ExamResult::updateOrCreate(
            ['user_id' => $userId, 'exam_id' => $examId],
            ['score' => $score]
        );

        // Kullanıcıyı sınav sonucunu gösterecek sayfaya yönlendirin
        return redirect()->route('exam.result.show', ['examId' => $examId]);
    }

    public function showUserResult($examId)
    {
        $userId = Auth::id(); // Giriş yapmış kullanıcının ID'si
        $examResult = ExamResult::where('exam_id', $examId)
            ->where('user_id', $userId)
            ->first(); // Kullanıcının sınav sonucunu al

        if (!$examResult) {
            return redirect()->back()->with('error', 'Sonuç bulunamadı.');
        }

        // Tüm mevcut sınavları al
        $exams = Exam::all();

        return view('exams.user_results', compact('examResult', 'exams'));
    }


    public function filter(Request $request)
    {
        $examId = $request->get('exam_id');
        $exams = Exam::all();

        if ($examId) {
            $results = ExamResult::with('user', 'exam')->where('exam_id', $examId)->get();
        } else {
            $results = ExamResult::with('user', 'exam')->get();
        }

        return view('admin.exam_results.index', compact('exams', 'results'));
    }

    public function results($examId)
    {
        $exams = Exam::all();
        $results = ExamResult::with('user', 'exam')->where('exam_id', $examId)->get();

        return view('admin.exam_results.index', compact('exams', 'results'));
    }

    public function getExamHistory($userId)
    {
        $results = ExamResult::with('exam')
            ->where('user_id', $userId)
            ->get(['exam_id', 'score', 'created_at']);

        $formattedResults = $results->map(function ($result) {
            return [
                'exam_name' => $result->exam->name,
                'score' => $result->score,
                'date' => $result->created_at->format('d-m-Y H:i')
            ];
        });

        return response()->json($formattedResults->toArray()); // Ensure this is an array
    }

    public function exportExcel()
    {
        return \Maatwebsite\Excel\Facades\Excel::download(new ExamResultsExport, 'sınav_sonuçları.xlsx');
    }

    public function exportPDF()
    {
        $results = ExamResult::with('user')->get(); // Sınav sonuçlarını alın

        // HTML içeriğini oluşturun
        $htmlContent = view('admin.exam_results.pdf', compact('results'))->render();

        // PDF oluşturma
        $html2pdf = new Html2Pdf();
        $html2pdf->writeHTML($htmlContent);
        return $html2pdf->output('sinav_sonuclari.pdf'); // PDF'yi tarayıcıda göster
    }


    // Bu metot, aylık sınav sonuçlarını sayfada göstermek için kullanılır
    public function monthlyResults(Request $request)
    {
        $startDate = now()->startOfMonth();
        $endDate = now()->endOfMonth();

        $results = ExamResult::with('user', 'exam')
            ->whereBetween('created_at', [$startDate, $endDate])
            ->get();

        $monthlyRewards = [];
        foreach ($results as $result) {
            if (!$result->user || !$result->user->id) {
                continue;
            }

            if (!isset($monthlyRewards[$result->user->id])) {
                $monthlyRewards[$result->user->id] = [
                    'user' => $result->user,
                    'total_reward' => 0,
                    'exams' => []
                ];
            }

            $monthlyRewards[$result->user->id]['total_reward'] += is_numeric($result->reward) ? $result->reward : 0;

            $monthlyRewards[$result->user->id]['exams'][] = $result;
        }

        return view('admin.exam_results.monthly', compact('monthlyRewards'));
    }

    // Bu metot, Excel indirme işlemi için verileri hazırlar
    public function getMonthlyResults()
    {
        $startDate = now()->startOfMonth();
        $endDate = now()->endOfMonth();

        $results = ExamResult::with('user', 'exam')
            ->whereBetween('created_at', [$startDate, $endDate])
            ->get();

        $monthlyRewards = [];
        foreach ($results as $result) {
            if (!$result->user || !$result->user->id) {
                continue;
            }

            if (!isset($monthlyRewards[$result->user->id])) {
                $monthlyRewards[$result->user->id] = [
                    'user' => $result->user,
                    'total_reward' => 0,
                    'exams' => []
                ];
            }

            $monthlyRewards[$result->user->id]['total_reward'] += is_numeric($result->reward) ? $result->reward : 0;

            $monthlyRewards[$result->user->id]['exams'][] = $result;
        }

        return $monthlyRewards;
    }

    // Bu metot Excel dosyasını indirir
    public function downloadExcel()
    {

        $monthlyRewards = $this->getMonthlyResults();
        return Excel::download(new MonthlyResultsExport($monthlyRewards), 'aylik_sinav_sonuclari.xlsx');
    }
}
