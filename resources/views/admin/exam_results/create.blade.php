@extends('layouts.app')

@section('content')
<div class="container mt-5">
    <h1>Yeni Sonuç Ekle</h1>

    <form method="POST" action="{{ route('exam.store-result', ['examId' => $exam->id]) }}">
        @csrf
        <div class="form-group mb-3">
            <label for="employee_id">Çalışan:</label>
            <select id="employee_id" name="employee_id" class="form-control" required>
                <option value="">Çalışan Seçin</option>
                @foreach($employees as $employee)
                    <option value="{{ $employee->id }}">{{ $employee->name }}</option>
                @endforeach
            </select>
            @error('employee_id')
                <div class="alert alert-danger mt-2">{{ $message }}</div>
            @enderror
        </div>

        <div class="form-group mb-3">
            <label for="exam_id">Sınav:</label>
            <select id="exam_id" name="exam_id" class="form-control" required>
                <option value="">Sınav Seçin</option>
                @foreach($exams as $exam)
                    <option value="{{ $exam->id }}">{{ $exam->title }}</option>
                @endforeach
            </select>
            @error('exam_id')
                <div class="alert alert-danger mt-2">{{ $message }}</div>
            @enderror
        </div>

        <div class="form-group mb-3">
            <label for="score">Puan:</label>
            <input type="number" id="score" name="score" class="form-control" required>
            @error('score')
                <div class="alert alert-danger mt-2">{{ $message }}</div>
            @enderror
        </div>

        <button type="submit" class="btn btn-primary">Kaydet</button>
    </form>
</div>
@endsection
