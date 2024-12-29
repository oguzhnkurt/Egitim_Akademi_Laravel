<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use App\Models\File;
use Illuminate\Support\Facades\Storage;

class ProceduresFormsController extends Controller
{
    public function __construct()
    {
        $this->middleware('role:merkez-ofis'); // Yalnızca merkez ofis rolündekiler erişebilir
    }

    public function index(Request $request)
{
    // Arama terimini al
    $search = $request->query('search');

    // Arama yapılacak sorguyu oluştur
    $files = File::with('user')
        ->when($search, function ($query, $search) {
            return $query->where('title', 'like', '%' . $search . '%');
        })
        ->paginate(12);

    return view('admin.procedures-forms.index', compact('files'));
}



    public function store(Request $request)
    {
        $request->validate([
            'file' => 'required|mimes:docx,pdf,xlsx|max:20480', // Maksimum 20MB
            'user_id' => 'required|exists:users,id',
        ]);

        $file = $request->file('file');
        $path = $file->store('files', 'public');

        File::create([
            'title' => $file->getClientOriginalName(),
            'file_path' => $path,
            'type' => $file->getClientOriginalExtension(),
            'user_id' => $request->input('user_id'),
        ]);

        return back()->with('success', 'Dosya başarıyla yüklendi.');
    }


    public function download($id)
    {
        $file = File::findOrFail($id);
        $filePath = public_path('storage/files/' . basename($file->file_path));

        return response()->download($filePath, basename($file->file_path));
    }
}
