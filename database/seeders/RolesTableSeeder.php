<?php

namespace Database\Seeders;

use Illuminate\Database\Seeder;
use Spatie\Permission\Models\Role;
use Carbon\Carbon;

class RolesTableSeeder extends Seeder
{
    public function run()
    {
        Role::firstOrCreate(
            ['name' => 'admin'],
            [
                'guard_name' => 'web',
                'created_at' => Carbon::now(),
                'updated_at' => Carbon::now(),
            ]
        );

        Role::firstOrCreate(
            ['name' => 'user'],
            [
                'guard_name' => 'web',
                'created_at' => Carbon::now(),
                'updated_at' => Carbon::now(),
            ]
        );

        Role::firstOrCreate(
            ['name' => 'merkez-ofis'],
            [
                'guard_name' => 'web',
                'created_at' => Carbon::now(),
                'updated_at' => Carbon::now(),
            ]
        );
    }
}
