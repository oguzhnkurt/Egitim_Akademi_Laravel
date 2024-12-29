@extends('layouts.master')

@section('content')
<div class="container mt-5">
    <h3>{{ $file->title }} Düzenle</h3>

    @if (session('success'))
    <div class="alert alert-success">
        {{ session('success') }}
    </div>
    @endif

    <form method="POST" action="{{ route('files.update', $file->id) }}">
        @csrf

        <div class="mb-3">
            <label for="baslik" class="form-label">Başlık</label>
            <input type="text" class="form-control" id="baslik" name="baslik" value="{{ old('baslik', $file->title) }}" required>
        </div>

        <div class="mb-3">
            <label for="departman" class="form-label">Departman</label>
            <input type="text" class="form-control" id="departman" name="departman" value="{{ old('departman', $file->department) }}" required>
        </div>

        <div class="mb-3">
            <label for="tur" class="form-label">Tür</label>
            <input type="text" class="form-control" id="tur" name="tur" value="{{ old('tur', $file->type) }}" required>
        </div>

        <div class="mb-3">
            <label for="icerik" class="form-label">İçerik</label>
            <textarea class="form-control" id="icerik" name="icerik" rows="10" required>{{ old('icerik', $file->content) }}</textarea>
        </div>

        <button type="submit" class="btn btn-primary">Kaydet</button>
    </form>
</div>
@endsection