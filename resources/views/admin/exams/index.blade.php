@extends('layouts.master')

@section('content')
<div class="container">
    <h1>Sınavlar</h1>
    <a href="{{ route('admin.exams.create') }}" class="btn btn-primary mb-3">Yeni Sınav Ekle</a>
    <table class="table table-bordered">
        <thead>
            <tr>
                <th>ID</th>
                <th>Başlık</th>
                <th>Açıklama</th>
                <th>Başlangıç Tarihi</th>
                <th>Bitiş Tarihi</th>
                <th>Süre</th>
                <th>İşlemler</th>
            </tr>
        </thead>
        <tbody>
            @foreach($exams as $exam)
            <tr>
                <td>{{ $exam->id }}</td>
                <td>{{ $exam->title }}</td>
                <td>{{ $exam->description }}</td>
                <td>{{ $exam->start_date }}</td>
                <td>{{ $exam->end_date }}</td>
                <td>{{ $exam->duration }}</td>
                <td>
                    <a href="{{ route('admin.exams.edit', $exam->id) }}" class="btn btn-warning">Düzenle</a>
                    <form action="{{ route('admin.exams.destroy', $exam->id) }}" method="POST" style="display:inline;">
                        @csrf
                        @method('DELETE')
                        <button type="submit" class="btn btn-danger">Sil</button>
                    </form>
                </td>
            </tr>
            @endforeach
        </tbody>
    </table>
</div>
@endsection
