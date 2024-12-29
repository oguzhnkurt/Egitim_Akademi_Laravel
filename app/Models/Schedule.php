<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class Schedule extends Model
{
    use HasFactory;

    protected $fillable = ['day', 'time', 'link', 'instructor_id'];

    // Eğitmen ile ilişki
    public function instructor()
    {
        return $this->belongsTo(User::class, 'instructor_id');
    }
}