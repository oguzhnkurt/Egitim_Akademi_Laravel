<?php

namespace App\Http\Controllers;

use App\Models\User;
use Illuminate\Http\Request;
use Spatie\Permission\Models\Role;

class UserController extends Controller
{
    // Kullanıcı rol atama formunu göster
    public function showAssignRoleList()
    {
        $users = User::all();
        $roles = Role::all();
        return view('users.assignroles', compact('users', 'roles'));
    }

    // Kullanıcıya rol atama işlemi
    public function assignRole(Request $request)
    {
        $userId = $request->input('user_id'); // Kullanıcı ID'sini al
        $role = $request->input('role'); // Rol adını al

        $user = User::findOrFail($userId); // Kullanıcıyı bul

        // Kullanıcının zaten rolü olup olmadığını kontrol et
        if ($user->hasRole($role)) {
            return redirect()->back()->with('error', 'Bu kullanıcı zaten bu rolü sahip.');
        }

        // Kullanıcıya rolü ata
        $user->assignRole($role);

        return redirect()->route('users.showAssignRoleList')
            ->with('success', 'Rol başarıyla atandı.');
    }
    public function showUsersWithRoles()
    {
        // Tüm kullanıcıları ve bunlara atanan rolleri çek
        $users = User::with('roles')->get();

        // Rol bilgilerini bir view'e gönder
        return view('users.roles', compact('users'));
    }
}
    