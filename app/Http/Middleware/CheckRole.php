<?php

namespace App\Http\Middleware;

use Closure;
use Illuminate\Http\Request;
use Symfony\Component\HttpFoundation\Response;

class CheckRole
{
    public function handle(Request $request, Closure $next, ...$roles)
    {
        if (!$request->user() || !$request->user()->roles()->whereIn('name', $roles)->exists()) {
            return redirect()->route('home')->with('error', 'Bu sayfaya erişim izniniz yok.');
        }

        return $next($request);
    }
}
