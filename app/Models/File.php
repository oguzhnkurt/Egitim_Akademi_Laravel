<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class File extends Model
{
    protected $fillable = ['title', 'file_path', 'department_id', 'type', 'user_id'];

    public function department()
    {
        return $this->belongsTo(Department::class);
    }
    public function user()
{
    return $this->belongsTo(User::class);
}
}
