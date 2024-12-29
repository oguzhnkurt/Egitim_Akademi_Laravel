<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class Question extends Model
{
    use HasFactory;

    protected $fillable = [
        'exam_id', 'question_text', 'max_score', 'options',  'correct_option', 'order', 'image' ,'option_images'
    ];

    protected $casts = [
        'options' => 'array',
        'option_images' => 'array'
    ];

    public function exam()
    {
        return $this->belongsTo(Exam::class);
    }
}
