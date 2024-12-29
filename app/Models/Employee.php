<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class Employee extends Model
{
    use HasFactory;

    protected $fillable = [
        'tc_kimlik', 'adi_soyadi', 'ise_giris_tarihi', 'isten_ayrilis_tarihi', 'bolge_adi', 
        'gorev_yeri', 'gorevi', 'tel_no', 'email', 'foto', 'hesap_no'
    ];

    public function examResults()
    {
        return $this->hasMany(ExamResult::class);
    }

    public function trainings()
    {
        return $this->belongsToMany(Training::class, 'training_attendance');
    }

    public function meetings()
    {
        return $this->belongsToMany(Meeting::class, 'meeting_attendance');
    }
}
