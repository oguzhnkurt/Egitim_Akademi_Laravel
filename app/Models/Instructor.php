<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use App\Models\User;

class Instructor extends Model
{
    use HasFactory;

    protected $fillable = ['user_id', 'name', 'surname', 'email', 'external'];

    protected static function boot()
{
    parent::boot();

    static::creating(function ($instructor) {
        if ($instructor->user_id) {
            // İlişkiyi önceden yükleyin
            $user = User::find($instructor->user_id);
            if ($user) {
                $instructor->name = $user->name;
                $instructor->surname = $user->surname;
                $instructor->email = $user->email;
                $instructor->external = false; // İç eğitmen olarak işaretle
            }
        } else {
            if (!$instructor->name || !$instructor->surname || !$instructor->email) {
                throw new \Exception('Dış eğitmen bilgileri eksik. Lütfen tüm bilgileri doldurun.');
            }
            $instructor->external = true; // Dış eğitmen olarak işaretle
        }
    });

    static::updating(function ($instructor) {
        if ($instructor->user_id) {
            // İlişkiyi önceden yükleyin
            $user = User::find($instructor->user_id);
            if ($user) {
                $instructor->name = $user->name;
                $instructor->surname = $user->surname;
                $instructor->email = $user->email;
                $instructor->external = false; // İç eğitmen olarak işaretle
            }
        } else {
            if (!$instructor->name || !$instructor->surname || !$instructor->email) {
                throw new \Exception('Dış eğitmen bilgileri eksik. Lütfen tüm bilgileri doldurun.');
            }
            $instructor->external = true; // Dış eğitmen olarak işaretle
        }
    });
}

    
    // İlişkiyi tanımlayın
    public function user()
    {
        return $this->belongsTo(User::class);
    }

    // Eğitmen ile eğitim takvimi ilişkisi
    public function schedules()
    {
        return $this->hasMany(Schedule::class, 'instructor_id');
    }
    public function videos()
    {
        return $this->hasMany(Video::class);
    }
}

