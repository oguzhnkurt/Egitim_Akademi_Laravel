<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

return new class extends Migration
{
    /**
     * Run the migrations.
     */
    public function up()
    {
        Schema::create('exam_results', function (Blueprint $table) {
            $table->id();
            $table->unsignedBigInteger('user_id');
            $table->foreignId('exam_id')->constrained()->onDelete('cascade');
            $table->integer('score')->default(0);
            $table->string('reward')->default('Hayır');
            $table->string('job_title')->nullable(); // Görev alanı
            $table->string('hospital')->nullable(); // Hastane alanı
            $table->timestamps();

            // user_id için kısıtlama
            $table->foreign('user_id')->references('id')->on('users')->onDelete('cascade');
        });
    }

    /**
     * Reverse the migrations.
     */
    public function down(): void
    {   
        
        Schema::dropIfExists('exam_results');

    }
};
