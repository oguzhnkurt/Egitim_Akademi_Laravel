<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class MeetingAttendance extends Model
{
    use HasFactory;

    protected $fillable = [
        'employee_id', 'meeting_id'
    ];

    public function employee()
    {
        return $this->belongsTo(Employee::class);
    }

    public function meeting()
    {
        return $this->belongsTo(Meeting::class);
    }
}