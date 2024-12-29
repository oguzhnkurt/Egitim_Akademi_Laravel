<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

class CreateVideosTable extends Migration
{
    public function up()
    {
        Schema::create('videos', function (Blueprint $table) {
            $table->id();
            $table->string('title');
            $table->string('filename'); // 'path' yerine 'filename'
            $table->foreignId('instructor_id')->constrained()->onDelete('cascade'); // Eğitmen referansı
            $table->date('date')->nullable(); // Eğitim tarihi ekleyin
            $table->timestamps();
        });
    }

    public function down()
    {
        Schema::dropIfExists('videos');
    }
}
