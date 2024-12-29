<?php

namespace App\Http\Controllers\Admin;

use App\Http\Controllers\Controller;
use Illuminate\Http\Request;
use App\Models\Exam;
use App\Models\ExamResult;
use App\Models\User;
use App\Models\Course;
use App\Models\Question;
use Illuminate\Support\Facades\Storage;

class AdminExamController extends Controller
{
    public function __construct()
    {
        $this->middleware(['auth', 'role:admin']);
    }

    public function index()
    {
        $exams = Exam::all();
        return view('admin.exams.index', compact('exams'));
    }

    public function create()
    {
        $courses = Course::all();
        return view('admin.exams.create', compact('courses'));
    }

    public function store(Request $request)
    {
        $validated = $request->validate([
            'title' => 'required|string|max:255',
            'description' => 'required|string',
            'start_date' => 'required|date',
            'end_date' => 'required|date',
            'duration' => 'required|integer',
            'questions' => 'required|array',
            'questions.*.question_text' => 'required|string',
            'questions.*.max_score' => 'required|integer',
            'questions.*.options' => 'required|array',
            'questions.*.correct_option' => 'required|string',
            'questions.*.option_images.*' => 'nullable|image|mimes:jpeg,png,jpg,gif|max:2048' // Şık görsel validasyonu
        ]);

        $exam = new Exam();
        $exam->title = $request->title;
        $exam->description = $request->description;
        $exam->start_date = $request->start_date;
        $exam->end_date = $request->end_date;
        $exam->duration = $request->duration;
        $exam->save();

        if (is_array($request->questions)) {
            foreach ($request->questions as $index => $questionData) {
                $question = new Question();
                $question->exam_id = $exam->id;
                $question->question_text = $questionData['question_text'];
                $question->max_score = $questionData['max_score'];
                $question->options = json_encode($questionData['options']);
                $question->correct_option = $questionData['correct_option'];

                // Şıkların görsellerini kaydetme işlemi
                $optionImages = [];
                if (isset($questionData['option_images'])) {
                    foreach ($questionData['option_images'] as $optionIndex => $image) {
                        if ($image) {
                            $imagePath = $image->store('exam_images', 'public'); // Görseli kaydet
                            $optionImages[$optionIndex] = $imagePath; // Path'i JSON olarak sakla
                        }
                    }
                }
                $question->option_images = json_encode($optionImages); // Şık görsellerini JSON olarak sakla
                $question->save();
            }
        }

        return redirect()->route('admin.exams.index')->with('success', 'Sınav başarıyla oluşturuldu.');
    }


    public function show(Exam $exam)
    {
        return view('admin.exams.show', compact('exam'));
    }

    public function edit(Exam $exam)
    {
        $courses = Course::all();
        return view('admin.exams.edit', compact('exam', 'courses'));
    }

    public function update(Request $request, Exam $exam)
    {
        $request->validate([
            'title' => 'required|string|max:255',
            'description' => 'required|string',
            'start_date' => 'required|date',
            'end_date' => 'required|date|after_or_equal:start_date',
            'duration' => 'required|integer|min:1',
            'questions' => 'array',
            'questions.*.question_text' => 'required|string',
            'questions.*.max_score' => 'required|integer|min:1',
            'questions.*.options' => 'required|array',
            'questions.*.correct_option' => 'required|string',
            'questions.*.image' => 'nullable|file|image|max:2048', // Görsel yükleme doğrulaması
            'questions.*.correct_option.required' => 'Her soru için bir doğru cevap seçilmelidir.'
        ]);

        $exam->update($request->only(['title', 'description', 'start_date', 'end_date', 'duration']));

        foreach ($exam->questions as $question) {
            if (isset($request->questions[$question->id])) {
                $questionData = $request->questions[$question->id];
                $question->update([
                    'question_text' => $questionData['question_text'],
                    'max_score' => $questionData['max_score'],
                    'options' => json_encode($questionData['options']),
                    'correct_option' => $questionData['correct_option'],
                ]);

                // Eğer yeni bir görsel yüklenmişse
                if (isset($questionData['image']) && $questionData['image'] instanceof \Illuminate\Http\UploadedFile) {
                    // Eski görseli sil
                    if ($question->image) {
                        Storage::disk('public')->delete($question->image);
                    }

                    // Yeni görseli kaydet
                    $path = $questionData['image']->store('exam_images', 'public'); // Yeni görseli exam_images klasörüne kaydet
                    $question->image = $path;
                }

                $question->save();
            } else {
                $question->delete(); // Eğer soru güncellenmemişse, silinir
            }
        }

        // Yeni eklenen soruları ekle
        foreach ($request->questions as $key => $questionData) {
            if (strpos($key, 'new') !== false) {
                $this->createQuestion($exam, $questionData);
            }
        }

        return redirect()->route('admin.exams.index')->with('success', 'Sınav başarıyla güncellendi.');
    }

    public function destroy(Exam $exam)
    {
        foreach ($exam->questions as $question) {
            // Her soru için varsa görselleri sil
            if ($question->image) {
                Storage::disk('public')->delete($question->image);
            }
        }

        $exam->delete();

        return redirect()->route('admin.exams.index')->with('success', 'Sınav başarıyla silindi.');
    }

    private function createQuestion(Exam $exam, array $questionData)
    {
        $question = new Question();
        $question->exam_id = $exam->id;
        $question->question_text = $questionData['question_text'];
        $question->max_score = $questionData['max_score'];
        $question->options = json_encode($questionData['options']);
        $question->correct_option = $questionData['correct_option'];

        // Eğer bir görsel yüklenmişse
        if (isset($questionData['image']) && $questionData['image'] instanceof \Illuminate\Http\UploadedFile) {
            $path = $questionData['image']->store('exam_images', 'public'); // Görselleri exam_images klasörüne kaydet
            $question->image = $path;
        }

        $question->save();
    }

    public function uploadImage(Request $request)
    {
        if ($request->hasFile('image')) {
            // Görseli geçici olarak 'exam_images' klasörüne kaydet
            $image = $request->file('image');
            $path = $image->store('exam_images', 'public'); // Dosyaları exam_images klasörüne kaydet

            // Görsel yolunu döndür
            return response()->json(['path' => Storage::url($path), 'message' => 'Görsel başarıyla yüklendi.']);
        }

        return response()->json(['error' => 'Görsel yüklenemedi.'], 400);
    }

    public function end($id)
    {
        $exam = Exam::findOrFail($id);

        // Sınava giren kullanıcıların cevaplarını kontrol et
        $participants = $exam->participants; // Sınava katılan kullanıcılar
        foreach ($participants as $participant) {
            $this->evaluateUserExam($participant, $exam);
        }

        // Sınavı bitirme işlemi
        $exam->update(['status' => 'completed']);

        return redirect()->route('admin.exams.index')->with('success', 'Sınav başarıyla bitirildi.');
    }

    private function evaluateUserExam(User $user, Exam $exam)
    {
        $totalScore = 0;

        // Kullanıcının bu sınavdaki cevaplarını kontrol et
        $userAnswers = $user->answers()->where('exam_id', $exam->id)->get();

        foreach ($exam->questions as $question) {
            // Doğru cevabı ve kullanıcının işaretlediği cevabı kontrol et
            $userAnswer = $userAnswers->where('question_id', $question->id)->first();
            if ($userAnswer && $userAnswer->answer == $question->correct_option) {
                $totalScore += $question->max_score; // Doğru cevap, maksimum puanı ekle
            }
        }

        // Kullanıcının sınav sonucunu kaydet
        ExamResult::updateOrCreate(
            ['user_id' => $user->id, 'exam_id' => $exam->id],
            ['score' => $totalScore]
        );
    }
}
