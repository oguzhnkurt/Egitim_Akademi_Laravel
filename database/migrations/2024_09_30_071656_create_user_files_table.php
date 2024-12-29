<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

class CreateUserFilesTable extends Migration
{
    public function up()
    {
        Schema::create('user_files', function (Blueprint $table) {
            $table->id();
            $table->foreignId('user_id')->constrained()->onDelete('cascade'); // Hangi kullanıcı düzenlemiş
            $table->foreignId('file_id')->constrained()->onDelete('cascade'); // Hangi dosya düzenlenmiş
            $table->timestamp('edited_at'); // Düzenlenme tarihi
            $table->text('content'); // Dosyanın yeni içeriği
            $table->timestamps();
        });
    }

    public function down()
    {
        Schema::dropIfExists('user_files');
    }
}