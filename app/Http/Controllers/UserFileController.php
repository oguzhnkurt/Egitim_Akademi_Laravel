<?php

namespace App\Http\Controllers;

use App\Models\UserFile;

class UserFileController extends Controller
{
    public function index()
    {
        $userFiles = UserFile::with('user', 'file')->get();
        return view('user_files.index', compact('userFiles'));
    }
}
