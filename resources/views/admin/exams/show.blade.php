@extends('layouts.master')

@section('content')
<div class="container">
    <h1>Sınav Detayları</h1>
    <div class="form-group">
        <label for="title">Başlık</label>
        <input type="text" class="form-control" id="title" name="title" value="{{ $exam->title }}" readonly>
    </div>
    <div class="form-group">
        <label for="description">Açıklama</label>
        <textarea class="form-control" id="description" name="description" readonly>{{ $exam->description }}</textarea>
    </div>
    <div class="form-group">
        <label for="start_date">Başlangıç Tarihi</label>
        <input type="date" class="form-control" id="start_date" name="start_date" value="{{ $exam->start_date }}" readonly>
    </div>
    <div class="form-group">
        <label for="end_date">Bitiş Tarihi</label>
        <input type="date" class="form-control" id="end_date" name="end_date" value="{{ $exam->end_date }}" readonly>
    </div>
    <div class="form-group">
        <label for="duration">Süre (dakika)</label>
        <input type="number" class="form-control" id="duration" name="duration" value="{{ $exam->duration }}" readonly>
    </div>
    <a href="{{ route('admin.exams.index') }}" class="btn btn-primary">Geri Dön</a>
</div>
@endsection
