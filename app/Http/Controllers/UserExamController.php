<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use App\Models\Exam;
use App\Models\Question;
use App\Models\ExamResult;

class UserExamController extends Controller
{
    // List available exams
    public function index()
    {
        $exams = Exam::all();
        return view('exams.index', compact('exams'));
    }

    // Start an exam
    public function start(Request $request, Exam $exam)
    {
        $questions = Question::where('exam_id', $exam->id)->inRandomOrder()->limit(10)->get();
        return view('exams.exam_page', compact('exam', 'questions'));
    }

    // Submit answers and calculate the score
    

    // Display results (admin only)
    public function results()
    {
        if (!auth()->user()->role) { // hasRole yerine role kontrolü yapıldı
            return redirect()->route('user.exam-portal.index')->with('error', 'Bu sayfaya erişim izniniz yok.');
        }

        $results = ExamResult::with('exam', 'user')->orderBy('score', 'desc')->get();
        return view('exams.results', compact('results'));
    }


    public function finishExam(Request $request, $examId)
    {
        // Sınavı bitirme işlemi burada yapılır
        $exam = Exam::find($examId);

        // Sınavın bitirildiği işlemi veritabanına kaydedilir, kullanıcıya tamamlandı mesajı gösterilir.
        // Örneğin:
        $exam->status = 'completed'; // Örneğin, sınavın durumunu 'tamamlandı' yapabiliriz
        $exam->save();

        return redirect()->route('user.exam-portal')->with('success', 'Sınav başarıyla tamamlandı.');
    }
}

