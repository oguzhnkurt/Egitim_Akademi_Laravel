<?php

namespace App\Http\Controllers;

use App\Models\Department;
use Illuminate\Http\Request;

class DepartmentController extends Controller
{
    public function showFiles(Department $department, $type)
    {
        $files = $department->files()->where('type', $type)->get();
        return view('department.files', compact('files', 'department', 'type'));
    }
}
