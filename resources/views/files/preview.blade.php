@extends('layouts.master')

@section('content')
<div class="container">
    <h1>{{ $file->title }}</h1>

    <div class="pdf-container" style="position: relative;">
        <!-- PDF'i iframe içinde görüntüleme -->
        <iframe 
            src="{{ asset('storage/files/' . $file->filename) }}" 
            width="100%" 
            height="800px"
            style="border: none;"
            id="pdfFrame"
        ></iframe>
    </div>
</div>

<!-- Önizleme için JavaScript kodu -->
<script>
    // Sayfa açıldığında yazdırma işlemini engelle
    document.addEventListener('DOMContentLoaded', function() {
        // Sayfa açıldığında yazdırma işlemini engelle
        document.addEventListener('keydown', function(e) {
            if (e.ctrlKey && (e.key === 'p' || e.key === 's' || e.key === 'u')) {
                e.preventDefault();
                alert('Bu işlev devre dışı bırakıldı.');
            }
        });

        // Sağ tıklamayı engelle
        document.addEventListener('contextmenu', function(e) {
            e.preventDefault();
        });

        // Metin kopyalamayı engelle
        document.addEventListener('copy', function(e) {
            e.preventDefault();
        });
    });
</script>

<style>
    /* Kullanıcı seçimlerini devre dışı bırakma */
    body {
        -webkit-user-select: none;
        -moz-user-select: none;
        -ms-user-select: none;
        user-select: none;
    }
</style>
@endsection
