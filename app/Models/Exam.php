<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class Exam extends Model
{
    use HasFactory;

    protected $fillable = [
        'title',
        'description',
        'start_date',
        'end_date',
        'duration',
    ];

    public function questions()
    {
        return $this->hasMany(Question::class);
    }

    public function examResults()
    {
        return $this->hasMany(ExamResult::class);
    }
}