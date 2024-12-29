<?php

namespace Database\Seeders;

use Illuminate\Database\Seeder;
use App\Models\User; // Admin yerine User modelini kullanmalısınız
use Illuminate\Support\Facades\Hash; // Hash sınıfını kullanmak için
use Spatie\Permission\Models\Role; // Role modeli için

class UsersTableSeeder extends Seeder
{
    public function run()
    {
        $adminRole = Role::where('name', 'admin')->first();
        $userRole = Role::where('name', 'user')->first();

    }
}