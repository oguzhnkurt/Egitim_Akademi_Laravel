<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

class CreateUserAnswersTable extends Migration
{
    public function up()
    {
        Schema::create('user_answers', function (Blueprint $table) {
            $table->id();
            $table->foreignId('user_id')->constrained()->onDelete('cascade'); // Kullanıcı kimliği
            $table->foreignId('exam_id')->constrained()->onDelete('cascade'); // Sınav kimliği
            $table->foreignId('question_id')->constrained()->onDelete('cascade'); // Soru kimliği
            $table->string('answer'); // Kullanıcının verdiği cevap
            $table->timestamps();
        });
    }

    public function down()
    {
        Schema::dropIfExists('user_answers');
    }
}
