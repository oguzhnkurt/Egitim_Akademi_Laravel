<?php

namespace App\Http\Controllers;

use App\Models\Exam;
use App\Models\Question;
use App\Models\ExamResult;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Log;
use App\Models\UserAnswer;
use App\Models\User;


class ExamPortalController extends Controller
{
    public function submit(Request $request, $examId)
    {
        if (!auth()->check()) {
            return redirect()->route('login')->with('error', 'Lütfen sınava girmeden önce giriş yapınız.');
        }

        $exam = Exam::findOrFail($examId);
        $questions = $request->input('questions'); // [question_id => user_answer]

        foreach ($questions as $questionId => $userAnswer) {
            \App\Models\UserAnswer::updateOrCreate(
                [
                    'user_id' => auth()->id(),
                    'exam_id' => $examId,
                    'question_id' => $questionId,
                ],
                ['answer' => $userAnswer['answer']]
            );
        }

        $user = User::findOrFail(auth()->id());
        $totalScore = $this->evaluateUserExam($user, $exam);

        // Kullanıcının sınav sonucunu ve ödül durumunu kaydet
        ExamResult::create([
            'exam_id' => $exam->id,
            'user_id' => $user->id,
            'score' => $totalScore,
            'status' => $totalScore >= 70 ? 'Başarılı' : 'Başarısız',
            'reward' => $totalScore >= 70 ? 'Evet' : 'Hayır' // Ödül durumu
        ]);

        return redirect()->route('user.exam.result', ['examId' => $examId])
            ->with('success', 'Sınav başarıyla tamamlandı.');
    }


    public function showUserResult($examId)
    {
        $userId = auth()->id(); // Giriş yapmış kullanıcının ID'sini alın
        $examResult = ExamResult::where('exam_id', $examId)
            ->where('user_id', $userId)
            ->first(); // Kullanıcının sınav sonucunu al

        // Eğer sınav sonucu bulunamazsa hata mesajı ile geri dön
        if (!$examResult) {
            return redirect()->route('exams.index')->with('error', 'Sonuç bulunamadı.');
        }

        // Kullanıcıya sonuçları göster
        return view('exams.user_results', compact('examResult'));
    }


    public function start(Request $request, $examId)
    {
        $exam = Exam::find($examId);
        if (!$exam) {
            return redirect()->route('exams.index')->with('error', 'Sınav bulunamadı.');
        }

        // Soruları şıklarıyla birlikte getir
        $questions = Question::where('exam_id', $examId)
            ->inRandomOrder()
            ->limit(10)
            ->get();

        return view('exams.exam_page', compact('exam', 'questions'));
    }


    public function end($examId)
    {
        $exam = Exam::findOrFail($examId);

        $participants = $exam->participants; // Sınava katılan kullanıcılar
        foreach ($participants as $participant) {
            $user = User::findOrFail($participant->id); // Her katılımcı için User nesnesi alınır
            $this->evaluateUserExam($user, $exam);
        }

        $exam->update(['status' => 'completed']);

        return redirect()->route('admin.exams.index')->with('success', 'Sınav başarıyla bitirildi.');
    }


    private function evaluateUserExam(User $user, Exam $exam)
    {
        $totalScore = 0;

        $userAnswers = \App\Models\UserAnswer::where('exam_id', $exam->id)
            ->where('user_id', $user->id)
            ->get();

        foreach ($exam->questions as $question) {
            $userAnswer = $userAnswers->where('question_id', $question->id)->first();
            if ($userAnswer && trim($userAnswer->answer) === trim($question->correct_option)) {
                $totalScore += $question->max_score; // Doğru cevapsa puanı ekle
            }
        }

        // Sınav sonucunu güncelle
        \App\Models\ExamResult::updateOrCreate(
            ['user_id' => $user->id, 'exam_id' => $exam->id],
            [
                'score' => $totalScore,
                'status' => $totalScore >= 70 ? 'Başarılı' : 'Başarısız',
                'reward' => $totalScore >= 70 ? 'Evet' : 'Hayır' // Ödül durumu
            ]
        );

        return $totalScore;
    }


    public function show(Exam $exam)
    {
        $questions = $exam->questions()->with('options')->get();
        return view('admin.exams.exam_page', compact('exam', 'questions'));
    }

    public function results($examId)
    {
        // Sınav ve sonuçları çek
        $exam = Exam::findOrFail($examId);
        $results = ExamResult::where('exam_id', $examId)
            ->orderBy('score', 'desc')
            ->get();

        return view('admin.exam_results.index', compact('exam', 'results'));
    }
}

