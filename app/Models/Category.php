<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Model;

class Category extends Model
{
    protected $fillable = ['name', 'parent_id'];

    // Top category'nin alt kategorileri
    public function subcategories()
    {
        return $this->hasMany(Category::class, 'parent_id');
    }

    // Mid category'nin üst kategorisi
    public function parent()
    {
        return $this->belongsTo(Category::class, 'parent_id');
    }
}
