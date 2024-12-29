<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use App\Models\Category;
use App\Models\Video;

class CategoryController extends Controller
{
    public function index()
    {
        
        $categories = Category::with('subcategories')->whereNull('parent_id')->get();
        return view('categories.index', compact('categories'));
    }

    // Alt kategoriye göre videoları getir
    public function getVideosBySubcategory($subcategoryId)
    {
        $subcategory = Category::with('videos')->find($subcategoryId);
        if (!$subcategory) {
            return response()->json(['error' => 'Subcategory not found'], 404);
        }

        return response()->json(['videos' => $subcategory->videos]);
    }

    // Alt kategorileri getir
    public function getSubcategories($categoryId)
    {
        $category = Category::with('subcategories')->find($categoryId);
        if (!$category) {
            return response()->json(['error' => 'Category not found'], 404);
        }

        return response()->json(['subcategories' => $category->subcategories]);
    }


    // Alt kategoriye göre videoları getir
    public function getVideos($subcategoryId)
    {
        $subcategory = Category::with('videos')->find($subcategoryId);
        if (!$subcategory) {
            return response()->json(['error' => 'Subcategory not found'], 404);
        }

        return response()->json(['videos' => $subcategory->videos]);
    }
}
