<?php

namespace App\Http\Controllers\Admin;

use App\Http\Controllers\Controller;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Storage;
use App\Models\Video;
use App\Models\Instructor;
use App\Models\Category;
use App\Models\Subcategory;

class AdminVideoController extends Controller
{
    // Videoların listelendiği sayfa
    public function index(Request $request)
    {
        $category_id = $request->input('category_id');
        
        $videosQuery = Video::query();
        
        if ($category_id) {
            $videosQuery->whereHas('categories', function ($query) use ($category_id) {
                $query->where('category_id', $category_id);
            });
        }
        
        $videos = $videosQuery->get();
        
        return view('admin.videos.index', compact('videos'));
    }


    // İdari eğitim videolarını listeleme
    public function idariEgitimler()
    {
        $videos = Video::where('category', 'administrative')->get();
        return view('videos.idari_egitimler', compact('videos'));
    }

    // Video yükleme formu
    public function create()
    {
        $instructors = Instructor::all();
        $categories = Category::with('subcategories')->get(); // Kategorileri ve alt kategorileri al
        return view('admin.videos.create', compact('instructors', 'categories'));
    }



    // Yeni video kaydetme işlemi
    public function store(Request $request)
    {
        $request->validate([
            'title' => 'required',
            'video' => 'required|mimes:mp4,mov,avi,wmv|max:20480',
            'instructor_id' => 'required',
            'date' => 'required|date',
            'category_id' => 'required',
            'subcategory_id' => 'required', // Alt kategori doğrulaması
        ]);

        $video = new Video();
        $video->title = $request->title;
        $video->instructor_id = $request->instructor_id;
        $video->date = $request->date;
        $video->category_id = $request->category_id; // Kategori kaydı
        $video->subcategory_id = $request->subcategory_id; // Alt kategori kaydı

        if ($request->hasFile('video')) {
            $videoPath = $request->file('video')->store('videos', 'public');
            $video->filename = $videoPath;
        }

        $video->save();

        return redirect()->route('admin.videos.index')->with('success', 'Video başarıyla yüklendi.');
    }

    public function getSubcategories($categoryId)
    {
        // Kategoriye ait alt kategorileri getir
        $subcategories = Category::where('parent_id', $categoryId)->get();

        // Alt kategorileri JSON formatında döndür
        return response()->json($subcategories);
    }


    public function edit($id)
    {
        $video = Video::findOrFail($id);
        $instructors = Instructor::all();
        $categories = Category::with('subcategories')->get(); // Kategorileri ve alt kategorileri al
        return view('admin.videos.edit', compact('video', 'instructors', 'categories'));
    }

    public function update(Request $request, $id)
    {
        $request->validate([
            'title' => 'required|string|max:255',
            'video' => 'nullable|mimes:mp4,mov,ogg,qt|max:204800',
            'instructor_id' => 'required|exists:instructors,id',
            'date' => 'required|date',
            'category_id' => 'required',
            'subcategory_id' => 'required',
        ]);

        $video = Video::findOrFail($id);
        $video->title = $request->title;
        $video->instructor_id = $request->instructor_id;
        $video->date = $request->date;
        $video->category_id = $request->category_id;
        $video->subcategory_id = $request->subcategory_id;

        if ($request->hasFile('video')) {
            Storage::disk('public')->delete('videos/' . $video->filename);
            $filename = time() . '_' . $request->video->getClientOriginalName();
            $request->video->storeAs('videos', $filename, 'public');
            $video->filename = $filename;
        }

        $video->save();

        return redirect()->back()->with('success', 'Video başarıyla güncellendi.');
    }



    // Video silme işlemi
    public function destroy($id)
    {
        $video = Video::findOrFail($id);

        // Dosyayı sil
        Storage::disk('public')->delete('videos/' . $video->filename);

        // Veritabanından kaydı sil
        $video->delete();

        return redirect()->route('admin.videos.index')->with('success', 'Video başarıyla silindi.');
    }
}
