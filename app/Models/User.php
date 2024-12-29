<?php

namespace App\Models;

use Illuminate\Contracts\Auth\MustVerifyEmail;
use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Foundation\Auth\User as Authenticatable;
use Illuminate\Notifications\Notifiable;
use Laravel\Sanctum\HasApiTokens;
use Spatie\Permission\Traits\HasRoles;

class User extends Authenticatable
{
    use HasApiTokens, HasFactory, Notifiable, HasRoles;

    /**
     * The attributes that are mass assignable.
     *
     * @var array<int, string>
     */
    protected $fillable = [
        'name',
        'surname',
        'tc_kimlik', // T.C Kimlik
        'job_start_date', // İşe başlama tarihi
        'job_quit_date', // İşten çıkış tarihi
        'email',
        'password',
        'country_code',
        'phone_number',                                                      
        'region', 
        'hospital',
        'job_title',
    ];

    public function roles()
    {
        return $this->belongsToMany(Role::class, 'role_user');
    }

    public function hasRole($role)
    {
        return $this->roles()->where('name', $role)->exists();
    }

    public function enrollments()
    {
        return $this->hasMany(Enrollment::class);
    }

    public function courses()
    {
        return $this->hasMany(Course::class);
    }

    public function exams()
    {
        return $this->hasMany(Exam::class);
    }

    public function examResults()
    {
        return $this->hasMany(ExamResult::class);
    }

    public function instructors()
    {
        return $this->hasOne(Instructor::class);
    }
    public function files()
    {
        return $this->hasMany(File::class);
    }
}
