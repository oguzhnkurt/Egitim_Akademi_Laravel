<?php

namespace App\Http\Controllers\Admin;

use App\Http\Controllers\Controller;
use App\Models\Schedule;
use Illuminate\Http\Request;
use App\Models\Instructor;
use App\Models\User;
use Illuminate\Support\Facades\Log;
use Exception;
use Illuminate\Support\Str;


class AdminScheduleController extends Controller
{
    public function index()
    {
        $schedules = Schedule::with('instructor')->get();
        return view('admin.schedule.index', compact('schedules'));
    }

    public function create()
    {
        $instructors = Instructor::with('user')->get();
        return view('admin.schedule.create', compact('instructors'));
    }

    public function store(Request $request)
{
    $request->validate([
        'day' => 'required|string',
        'time' => 'required|string',
        'link' => 'required|string',
        'instructor_selection' => 'required|string',
        'instructor_name' => 'nullable|string',
        'instructor_surname' => 'nullable|string',
        'instructor_email' => 'nullable|email',
        'instructor_type' => 'nullable|string'
    ]);

    if (Str::startsWith($request->instructor_selection, 'existing-')) {
        $instructorId = str_replace('existing-', '', $request->instructor_selection);
        
        $user = User::find($instructorId);
        
        if (!$user) {
            return redirect()->back()->withErrors('Seçilen kullanıcı mevcut değil.');
        }
    } else {
        // Yeni bir eğitmen yaratıldığında, kullanıcıyı da oluşturuyoruz
        $user = User::create([
            'name' => $request->instructor_name,
            'surname' => $request->instructor_surname,
            'email' => $request->instructor_email,
            'is_instructor' => $request->instructor_type === 'internal',
        ]);
    
        $instructor = Instructor::create([
            'user_id' => $user->id,
            'name' => $request->instructor_name,
            'surname' => $request->instructor_surname,
            'email' => $request->instructor_email,
            'external' => $request->instructor_type === 'external',
        ]);
    
        $instructorId = $user->id;
    }
    
    Schedule::create([
        'day' => $request->day,
        'time' => $request->time,
        'link' => $request->link,
        'instructor_id' => $instructorId, // Bu artık users tablosundaki ID'dir
    ]);
    
    return redirect()->route('schedule.index')->with('success', 'Eğitim başarıyla oluşturuldu.');
    
}
    public function destroy($id)
    {
        $schedule = Schedule::findOrFail($id);
        $schedule->delete();
        return redirect()->route('schedule.index')->with('success', 'Eğitim başarıyla silindi.');
    }

    public function edit($id)
    {
        $schedule = Schedule::with('instructor')->findOrFail($id);
        $instructors = Instructor::all();
        return view('admin.schedule.edit', compact('schedule', 'instructors'));
    }

    public function update(Request $request, $id)
    {
        $validatedData = $request->validate([
            'day' => 'required|string|max:255',
            'time' => 'required|string|max:255',
            'link' => 'required|url',
            'instructor_selection' => 'required|string|in:new,existing-*',
            'instructor_name' => 'required_if:instructor_selection,new',
            'instructor_surname' => 'required_if:instructor_selection,new',
            'instructor_email' => 'required_if:instructor_selection,new|email',
            'instructor_type' => 'required_if:instructor_selection,new'
        ]);

        $schedule = Schedule::findOrFail($id);
        $instructor_id = $this->handleInstructor($validatedData);

        try {
            $updated = $schedule->update([
                'day' => $request->day,
                'time' => $request->time,
                'link' => $request->link,
                'instructor_id' => $instructor_id,
            ]);

            Log::info('Eğitim Güncellendi:', $schedule->toArray());

            if ($updated) {
                return redirect()->back()->with('success', 'Eğitim başarıyla güncellendi.');
            } else {
                return redirect()->back()->with('error', 'Eğitim güncellenirken bir hata oluştu.');
            }
        } catch (Exception $e) {
            Log::error('Eğitim güncellenirken bir hata oluştu:', ['error' => $e->getMessage()]);
            return redirect()->back()->with('error', 'Eğitim güncellenirken bir hata oluştu.');
        }
    }

    /**
     * Handles instructor operations (adds a new instructor or selects an existing one).
     */
    private function handleInstructor(array $data)
    {
        if (isset($data['instructor_selection']) && $data['instructor_selection'] === 'new') {
            $instructor = Instructor::create([
                'name' => $data['instructor_name'],
                'surname' => $data['instructor_surname'],
                'email' => $data['instructor_email'],
                'external' => $data['instructor_type'] === 'external' ? 1 : 0,
            ]);
            return $instructor->id;
        } else {
            return explode('-', $data['instructor_selection'])[1];
        }
    }
}
