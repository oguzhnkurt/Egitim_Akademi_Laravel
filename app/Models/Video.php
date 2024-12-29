<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Carbon\Carbon;

class Video extends Model
{
    use HasFactory;

    protected $fillable = ['title', 'filename', 'instructor_id', 'date', 'category_id', 'subcategory_id'];

    // Eğitmen ile olan ilişki
    public function instructor()
    {
        return $this->belongsTo(Instructor::class);
    }

    // Tarih formatını düzenle
    public function getDateAttribute($value)
    {
        Carbon::setLocale('tr');
        return Carbon::parse($value);
    }

    // Kategori ile ilişki
    // Video Modeli
    public function category()
    {
        return $this->belongsTo(Category::class, 'category_id');
    }

    public function subcategory()
    {
        return $this->belongsTo(Category::class, 'subcategory_id');
    }
}
