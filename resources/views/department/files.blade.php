@extends('layouts.app')

@section('content')
    <h1>{{ $department->name }} - {{ ucfirst($type) }}</h1>
    <ul>
        @foreach($files as $file)
            <li>{{ $file->title }}: {{ $file->content }}</li>
        @endforeach
    </ul>
@endsection
