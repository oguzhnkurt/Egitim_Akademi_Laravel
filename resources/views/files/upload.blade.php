@extends('layouts.master')

@section('content')
<div class="container">
    <div class="row">
        <div class="col-md-8 offset-md-2">
            <h1>Dosya Yükle</h1>
            <form action="{{ route('files.upload') }}" method="POST" enctype="multipart/form-data">
                @csrf
                <div class="mb-3">
                    <label for="title" class="form-label">Başlık</label>
                    <input type="text" class="form-control" id="title" name="title" required>
                </div>
                <div class="mb-3">
                    <label for="department" class="form-label">Departman</label>
                    <select class="form-select" id="department" name="department_id" required>
                        <option value="" disabled selected>Departman Seçin</option>
                        @foreach ($departments as $department)
                            <option value="{{ $department->id }}">{{ $department->name }}</option>
                        @endforeach
                    </select>
                </div>
                <div class="mb-3">
                    <label for="type" class="form-label">Tür</label>
                    <select class="form-select" id="type" name="type" required>
                        <option value="prosedur">Prosedür</option>
                        <option value="is_akisi">İş Akışı</option>
                        <option value="form">Formlar</option>
                        <option value="dilekce">Dilekçeler</option>
                    </select>
                </div>
                <div class="mb-3">
                    <label for="file" class="form-label">Dosya Yükle</label>
                    <input type="file" class="form-control" id="file" name="file" required>
                </div>
                <button type="submit" class="btn btn-primary">Yükle</button>
            </form>
        </div>
    </div>
</div>
@endsection
