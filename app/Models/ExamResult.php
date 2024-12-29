<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class ExamResult extends Model
{
    use HasFactory;

    protected $fillable = ['user_id', 'exam_id', 'score', 'reward', 'job_title', 'hospital'];

    public static function calculateScore($userId, $examId, $answers)
{
    // Exam bulunamadığında işlem yapma
    $exam = Exam::find($examId);
    if (!$exam) {
        return ['score' => 0, 'reward' => 'Hayır'];
    }

    // Eğer answers null ise boş bir dizi olarak işleme al
    if (is_null($answers)) {
        $answers = [];
    }

    $score = 0;
    $reward = 'Hayır'; // Varsayılan ödül durumu

    foreach ($answers as $questionId => $selectedOption) {
        $question = $exam->questions()->find($questionId);

        if ($question) {
            // Soru nesnesinin doğru cevabını al
            $correctOption = $question->correct_option;

            // Kullanıcının seçtiği cevap doğru mu?
            if ($correctOption == $selectedOption) {
                $score += $question->max_score; // Her doğru cevap için puanı ekle
            }
        }
    }

    // Ödül kontrolü
    if ($score >= 70) {
        $reward = 'Evet';
    }

    return ['score' => $score, 'reward' => $reward];
}

    // Eloquent model event
    protected static function boot()
    {
        parent::boot();

        static::creating(function ($examResult) {
            $user = $examResult->user;
            $examResult->job_title = $user->job_title;
            $examResult->hospital = $user->hospital;

            // Puan 70 ve üzeriyse ödülü ayarla
            if ($examResult->score >= 70) {
                $examResult->reward = '150 TL';
            }
        });

        static::updating(function ($examResult) {
            // Puan 70 ve üzeriyse ödülü ayarla
            if ($examResult->score >= 70) {
                $examResult->reward = '150 TL';
            } else {
                $examResult->reward = null; // 70 altı puanlar için ödülü sıfırla
            }
        });
    }

    public function user()
    {
        return $this->belongsTo(User::class); // Her sınav sonucunun bir kullanıcıya ait olduğunu belirtir
    }

    public function exam()
    {
        return $this->belongsTo(Exam::class);
    }
}
