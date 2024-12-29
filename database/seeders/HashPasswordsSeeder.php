<?php

namespace Database\Seeders;

use Illuminate\Database\Seeder;
use App\Models\User;
use Illuminate\Support\Facades\Hash;

class HashPasswordsSeeder extends Seeder
{
    public function run()
    {
        // Tüm kullanıcıları al
        $users = User::all();

        foreach ($users as $user) {
            // Şifreyi hash'le ve veritabanında güncelle
            if ($user->password) {
                $user->password = Hash::make($user->password);
                $user->save();
            }
        }
    }
}
