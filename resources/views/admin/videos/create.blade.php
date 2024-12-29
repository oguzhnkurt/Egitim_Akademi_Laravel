@extends('layouts.master')

@section('content')
<div class="container">
    <h2>Video Yükle</h2>
    <form action="{{ route('videos.store') }}" method="POST" enctype="multipart/form-data">
        @csrf
        <div class="form-group">
            <label for="title">Video Başlığı</label>
            <input type="text" name="title" class="form-control" required>
        </div>

        <div class="form-group">
            <label for="category">Ana Kategori</label>
            <select id="category" class="form-control" required>
                <option value="">Ana Kategori Seç</option>
                @foreach ($categories as $category)
                    <option value="{{ $category->id }}">{{ $category->name }}</option>
                @endforeach
            </select>
        </div>

        <div class="form-group">
            <label for="subcategory">Alt Kategori</label>
            <select name="category_id" id="subcategory" class="form-control" required>
                <option value="">Önce Ana Kategori Seç</option>
            </select>
        </div>

        <div class="form-group">
            <label for="video">Video Dosyası</label>
            <input type="file" name="video" class="form-control" required>
        </div>

        <button type="submit" class="btn btn-primary">Video Yükle</button>
    </form>
</div>
@endsection

@section('scripts')
<script>
    // Ana kategori seçildiğinde ilgili alt kategorileri getir
    document.getElementById('category').addEventListener('change', function () {
        var categoryId = this.value;
        var subcategorySelect = document.getElementById('subcategory');

        subcategorySelect.innerHTML = '<option value="">Alt Kategori Seç</option>'; // Önceki seçenekleri temizle

        if (categoryId) {
            fetch('/api/categories/' + categoryId + '/subcategories')
                .then(response => response.json())
                .then(data => {
                    data.forEach(function (subcategory) {
                        var option = document.createElement('option');
                        option.value = subcategory.id;
                        option.textContent = subcategory.name;
                        subcategorySelect.appendChild(option);
                    });
                });
        }
    });
</script>
@endsection
