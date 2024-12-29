<?php

// app/Http/Controllers/TrainingAttendanceController.php

namespace App\Http\Controllers;

use App\Models\TrainingAttendance;
use Illuminate\Http\Request;

class TrainingAttendanceController extends Controller
{
    public function index()
    {
        $attendances = TrainingAttendance::with(['employee', 'training'])->get();
        return view('trainingAttendance.index', compact('attendances'));
    }

    public function create()
    {
        return view('trainingAttendance.create');
    }

    public function store(Request $request)
    {
        $request->validate([
            'employee_id' => 'required|exists:employees,id',
            'training_id' => 'required|exists:trainings,id',
        ]);

        TrainingAttendance::create($request->all());

        return redirect()->route('trainingAttendance.index')->with('success', 'Attendance recorded successfully.');
    }

    public function show(TrainingAttendance $attendance)
    {
        return view('trainingAttendance.show', compact('attendance'));
    }

    public function edit(TrainingAttendance $attendance)
    {
        return view('trainingAttendance.edit', compact('attendance'));
    }

    public function update(Request $request, TrainingAttendance $attendance)
    {
        $request->validate([
            'employee_id' => 'required|exists:employees,id',
            'training_id' => 'required|exists:trainings,id',
        ]);

        $attendance->update($request->all());

        return redirect()->route('trainingAttendance.index')->with('success', 'Attendance updated successfully.');
    }

    public function destroy(TrainingAttendance $attendance)
    {
        $attendance->delete();

        return redirect()->route('trainingAttendance.index')->with('success', 'Attendance deleted successfully.');
    }
}
