@extends('layouts.master')

@section('content')
<div class="container mt-5">
    <h3 class="mb-4">Arama Sonuçları</h3>

    @if($results->isEmpty())
    <div class="alert alert-warning">
        Sonuç bulunamadı. Farklı bir arama terimi deneyin.
    </div>
    @else
    <div class="list-group">
        @foreach ($results as $file)
        <a href="{{ route('department.files', ['department' => $file->department_id, 'type' => $file->type]) }}"
            class="list-group-item list-group-item-action">
            <h5 class="mb-1">{{ $file->title }} - <span class="badge bg-secondary">{{ ucfirst($file->type) }}</span>
            </h5>
            <p class="mb-1">
                {{ Str::limit($file->content, 150) }} <!-- İçerikten kısa bir özet gösteriyoruz -->
            </p>
            <small>
                İlgili departman:
                <span class="department-name department-{{ $file->department_id }}">{{ $file->department->name ??
                    'Bilinmiyor' }}</span> <!-- Departman ismi için dinamik sınıf atıyoruz -->
            </small>
        </a>
        @endforeach
    </div>
    @endif
</div>
@endsection

@section('nofooter')
<!-- Bu sayfada footer'ı göstermemek için bu section'ı ekliyoruz -->
@endsection

<style>
    .department-name {
        color: #FF5733;
        /* Departman isimlerine özel renk */
        font-weight: bold;
        /* İsterseniz kalın font kullanabilirsiniz */
    }
</style>