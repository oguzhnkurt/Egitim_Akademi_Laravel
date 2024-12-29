<?php

namespace App\Http\Controllers;

use Illuminate\Support\Facades\Log;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Storage;
use PhpOffice\PhpWord\IOFactory;
use App\Models\File;
use App\Models\UserFile;
use App\Models\Department;
use PhpOffice\PhpWord\Element\Text;
use PhpOffice\PhpWord\Element\TextRun;
use PhpOffice\PhpWord\Element\Paragraph;
use PhpOffice\PhpWord\Shared\Html;


class FileController extends Controller
{
    public function showUploadForm()
    {
        $departments = Department::all();
        return view('files.upload', compact('departments'));
    }

    public function upload(Request $request)
    {
        // Validation
        $request->validate([
            'title' => 'required|string|max:255',
            'department_id' => 'required|exists:departments,id',
            'type' => 'required|in:prosedur,is_akisi,form,dilekce',
            'file' => 'required|file|mimes:pdf,xlsx,docx|max:20480',
        ]);

        // Dosya nesnesini al
        $file = $request->file('file');

        // Dosya adını oluştur ve dosyayı kaydet
        $filename = time() . '_' . $file->getClientOriginalName();
        $filePath = $file->storeAs('public/files', $filename);

        // Dosya yolunu veritabanına kaydet
        File::create([
            'title' => $request->title,
            'file_path' => 'storage/files/' . $filename, // Veritabanına kaydedilen yol
            'department_id' => $request->department_id,
            'type' => $request->type,
        ]);

        return redirect()->route('files.upload')->with('success', 'File uploaded successfully.');
    }

    public function showFiles($department, $type)
    {
        $departmentName = Department::findOrFail($department)->name;
        $files = File::where('department_id', $department)
            ->where('type', $type)
            ->get();

        return view('files.index', [
            'department' => $departmentName,
            'type' => $type,
            'files' => $files,
        ]);
    }

    public function download($id)
    {
        // Veritabanından dosya kaydını alın
        $file = File::findOrFail($id);

        // Eğer dosya türü "prosedur" ise indirmeyi engelleyin
        if ($file->type === 'prosedur') {
            return abort(403, 'Prosedür dosyaları indirilemez.');
        }

        // Dosya yolunu oluşturun
        $filePath = public_path('storage/files/' . basename($file->file_path));

        // Dosyanın var olup olmadığını kontrol edin
        if (!file_exists($filePath)) {
            return abort(404, 'File not found.');
        }

        // Dosyayı indirmek için geri döndürün
        return response()->download($filePath, basename($file->file_path));
    }

    public function openFile($id)
    {
        $file = File::findOrFail($id);

        // Eğer dosya türü 'prosedur' ise, sadece görüntüleme izin veriyoruz.
        if ($file->type === 'prosedur') {
            $filePath = public_path($file->file_path);
            return response()->file($filePath, [
                'Content-Type' => 'application/pdf',
                'Content-Disposition' => 'inline; filename="' . basename($filePath) . '"'
            ]);
        }

        // Diğer dosyalar için varsayılan şekilde açıyoruz.
        return response()->file(public_path($file->file_path));
    }
    public function view($id)
    {
        $file = File::findOrFail($id);

        // Eğer dosya türü prosedür ise, iframe ile sadece görüntüleme yapıyoruz
        if ($file->type === 'prosedur') {
            return view('files.iframe', ['file' => $file]);
        }

        // Diğer dosyalar için varsayılan indirme
        return response()->file(public_path($file->file_path));
    }

    public function edit($id)
    {
        // ID ile ilgili belgeyi veritabanından çek
        $file = File::findOrFail($id);

        // Depolanmış dosya yolunu al
        $docxFilePath = storage_path('app/files/' . $file->file_name); // Veritabanında saklanan dosya adıyla yolu birleştir

        // Belge türüne göre uygun şablon sayfasını döndür
        if ($file->type === 'dilekce' && $file->title === 'Sezin Zeyilname Dilekçesi') {
            return view('dilekce.sezin_Zeyilname_Dilekcesi', compact('file', 'docxFilePath'));
            
        } else {
            return redirect()->back()->with('error', 'Düzenlemek istediğiniz belge bulunamadı.');
        }
    }


    // Dosya düzenleme ve kaydetme işlemi
    public function update(Request $request, $id)
    {
        // Kullanıcıdan gelen verileri al
        $request->validate([
            'baslik' => 'required|string|max:255',
            'departman' => 'required|string|max:255',
            'tur' => 'required|string|max:255',
            'icerik' => 'required', // İlgili alanların doğrulaması
        ]);

        // Kullanıcı ve dosya bilgisini kaydet
        UserFile::create([
            'user_id' => auth()->user()->id, // Hangi kullanıcı düzenlemiş
            'file_id' => $id, // Hangi dosya düzenlenmiş
            'edited_at' => now(), // Düzenlenme tarihi
            'content' => $request->icerik // Yeni içerik
        ]);

        return redirect()->route('files.edit', $id)->with('success', 'Belge başarıyla güncellendi.');
    }
}
