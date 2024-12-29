<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use App\Models\Video;
use App\Models\Category;

class UserVideoController extends Controller
{
    public function index(Request $request)
    {
        $query = Video::query();

        // Kategori id'sine göre filtreleme
        if ($request->has('category_id')) {
            $query->where('category_id', $request->input('category_id'));
        }

        // Alt kategori id'sine göre filtreleme
        if ($request->has('subcategory_id')) {
            $query->where('subcategory_id', $request->input('subcategory_id'));
        }

        // Arama işlemi
        if ($request->has('search')) {
            $query->where('title', 'like', '%' . $request->input('search') . '%');
        }

        // Sadece category_type = 0 olan videoları almak için kategori ilişkisine göre filtreleme yapıyoruz
        $categoryType = 0;
        $query->whereHas('category', function ($q) use ($categoryType) {
            $q->where('category_type', $categoryType);
        });

        // Videoları eğitmenle birlikte alıyoruz
        $videos = $query->with('instructor')->get();

        // Sadece category_type = 0 olan kategorileri alıyoruz
        $categories = Category::where('category_type', $categoryType)
            ->whereNull('parent_id')
            ->with('subcategories')
            ->get();

        return view('videos.index', compact('videos', 'categories'));
    }


    public function create()
    {
        // Ana kategorileri ve alt kategorileri alıyoruz (category_type = 0 olanlar)
        $categoryType = 0;

        $categories = Category::where('category_type', $categoryType)
            ->whereNull('parent_id')
            ->with('subcategories')
            ->get();

        return view('videos.create', compact('categories'));
    }

    public function store(Request $request)
    {
        // Gerekli validasyonlar
        $request->validate([
            'title' => 'required|string|max:255',
            'category_id' => 'required|exists:categories,id', // Ana kategori ID kontrolü
            'subcategory_id' => 'nullable|exists:categories,id', // Alt kategori ID kontrolü
            // Diğer validasyon kuralları...
        ]);

        // Video oluşturma işlemi
        Video::create([
            'title' => $request->title,
            'category_id' => $request->category_id, // Seçilen ana kategori
            'subcategory_id' => $request->subcategory_id, // Seçilen alt kategori
            // Diğer alanlar...
        ]);

        return redirect()->route('videos.index')->with('success', 'Video başarıyla eklendi.');
    }

    public function show($id)
    {
        // Video ve eğitmen bilgileri
        $video = Video::with('instructor')->findOrFail($id);

        return view('videos.show', compact('video'));
    }

    public function idariEgitimler(Request $request)
    {
        $query = Video::query();

        // İdari eğitimler için category_type = 1 olan kategorileri filtreliyoruz
        $categoryType = 1;

        // Sadece category_type = 1 olan videoları almak için kategori ilişkisine göre filtreleme yapıyoruz
        $query->whereHas('category', function ($q) use ($categoryType) {
            $q->where('category_type', $categoryType);
        });

        // Alt kategori id'sine göre filtreleme (opsiyonel)
        if ($request->has('subcategory_id')) {
            $query->where('subcategory_id', $request->input('subcategory_id'));
        }

        // Arama işlemi
        if ($request->has('search')) {
            $query->where('title', 'like', '%' . $request->input('search') . '%');
        }

        // Videoları eğitmenle birlikte alıyoruz
        $videos = $query->with('instructor')->get();

        // İdari Eğitimler Kategorileri ve Alt Kategorileri (category_type = 1)
        $categories = Category::where('category_type', $categoryType)
            ->whereNull('parent_id')
            ->with('subcategories')
            ->get();

        return view('videos.idari_egitimler', compact('videos', 'categories'));
    }
}
