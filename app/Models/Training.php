<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class Training extends Model
{
    use HasFactory;

    protected $fillable = [
        'title', 'description', 'date'
    ];

    public function attendees()
    {
        return $this->belongsToMany(Employee::class, 'training_attendance');
    }
}
