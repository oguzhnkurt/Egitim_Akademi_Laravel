<?php

// app/Http/Controllers/MeetingAttendanceController.php

namespace App\Http\Controllers;

use App\Models\MeetingAttendance;
use Illuminate\Http\Request;

class MeetingAttendanceController extends Controller
{
    public function index()
    {
        $attendances = MeetingAttendance::with(['employee', 'meeting'])->get();
        return view('meetingAttendance.index', compact('attendances'));
    }

    public function create()
    {
        return view('meetingAttendance.create');
    }

    public function store(Request $request)
    {
        $request->validate([
            'employee_id' => 'required|exists:employees,id',
            'meeting_id' => 'required|exists:meetings,id',
        ]);

        MeetingAttendance::create($request->all());

        return redirect()->route('meetingAttendance.index')->with('success', 'Attendance recorded successfully.');
    }

    public function show(MeetingAttendance $attendance)
    {
        return view('meetingAttendance.show', compact('attendance'));
    }

    public function edit(MeetingAttendance $attendance)
    {
        return view('meetingAttendance.edit', compact('attendance'));
    }

    public function update(Request $request, MeetingAttendance $attendance)
    {
        $request->validate([
            'employee_id' => 'required|exists:employees,id',
            'meeting_id' => 'required|exists:meetings,id',
        ]);

        $attendance->update($request->all());

        return redirect()->route('meetingAttendance.index')->with('success', 'Attendance updated successfully.');
    }

    public function destroy(MeetingAttendance $attendance)
    {
        $attendance->delete();

        return redirect()->route('meetingAttendance.index')->with('success', 'Attendance deleted successfully.');
    }
}
