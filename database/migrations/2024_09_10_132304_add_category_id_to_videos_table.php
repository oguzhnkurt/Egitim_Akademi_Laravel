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
    Schema::table('videos', function (Blueprint $table) {
        $table->unsignedBigInteger('category_id')->nullable(); // Ana kategori
        $table->unsignedBigInteger('subcategory_id')->nullable(); // Alt kategori

        // Eğer `categories` tablosu ile ilişki kurmak isterseniz:
        $table->foreign('category_id')->references('id')->on('categories')->onDelete('cascade');
        $table->foreign('subcategory_id')->references('id')->on('categories')->onDelete('cascade');
    });
}

public function down()
{
    Schema::table('videos', function (Blueprint $table) {
        $table->dropForeign(['category_id']);
        $table->dropForeign(['subcategory_id']);
        $table->dropColumn('category_id');
        $table->dropColumn('subcategory_id');
    });
}

};
