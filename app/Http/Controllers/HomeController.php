<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use Illuminate\Support\Facades\Auth;

class HomeController extends Controller
{
    public function homepage()
    {
        return view('homepage');
    }

    public function home()
    {
        $user = Auth::user();
        return view('welcome', compact('user'));
    }
}
