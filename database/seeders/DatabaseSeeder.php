<?php

namespace Database\Seeders;

use Hash;
use Illuminate\Database\Seeder;

class DatabaseSeeder extends Seeder
{
    public function run()
    {
        $this->call(RolesTableSeeder::class);
        $this->call(UsersTableSeeder::class);

        // admin rolüne sahip kullanıcı oluşturulur
        $admin = \App\Models\User::create([

            'name' => 'Admin',
            'surname' => 'Admin',
            'email' => 'admin@example.com',
            'password' => Hash::make('password'),
        ]);

        $admin->assignRole('admin');
    }
}
